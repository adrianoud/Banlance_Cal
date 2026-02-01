import tkinter as tk
from tkinter import ttk, messagebox, filedialog, PhotoImage
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.widgets import SpanSelector
import numpy as np
import csv
import json
import os
import sys
import shutil  # 添加缺失的shutil导入
from datetime import datetime, timedelta  # 添加对timedelta的导入


# 尝试导入openpyxl用于Excel导出1
try:
    import openpyxl
    from openpyxl import Workbook
except ImportError:
    openpyxl = None

# 设置matplotlib字体以支持中文显示
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'FangSong', 'Arial Unicode MS', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

class ProjectManager:
    """项目管理器"""
    def __init__(self, app_root_path):
        self.app_root_path = app_root_path
        self.projects_dir = os.path.join(app_root_path, "projects")
        self.ensure_projects_directory()
        
    def ensure_projects_directory(self):
        """确保项目目录存在"""
        if not os.path.exists(self.projects_dir):
            os.makedirs(self.projects_dir)

    def get_project_list(self):
        """获取项目列表"""
        projects = []
        if os.path.exists(self.projects_dir):
            for item in os.listdir(self.projects_dir):
                project_path = os.path.join(self.projects_dir, item)
                if os.path.isdir(project_path):
                    project_info_path = os.path.join(project_path, "project_info.json")
                    if os.path.exists(project_info_path):
                        try:
                            with open(project_info_path, 'r', encoding='utf-8') as f:
                                project_info = json.load(f)
                                projects.append({
                                    'id': item,
                                    'name': project_info.get('name', item),
                                    'created_time': project_info.get('created_time', ''),
                                    'modified_time': project_info.get('modified_time', ''),
                                    'path': project_path
                                })
                        except Exception as e:
                            print(f"读取项目信息失败: {e}")
        return sorted(projects, key=lambda x: x['modified_time'], reverse=True)
        
    def create_project(self, name, description=""):
        """创建新项目"""
        # 生成项目ID
        project_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        project_path = os.path.join(self.projects_dir, project_id)
        
        # 创建项目目录
        os.makedirs(project_path)
        
        # 创建项目信息文件
        project_info = {
            'name': name,
            'description': description,
            'created_time': datetime.now().isoformat(),
            'modified_time': datetime.now().isoformat()
        }
        
        with open(os.path.join(project_path, "project_info.json"), 'w', encoding='utf-8') as f:
            json.dump(project_info, f, ensure_ascii=False, indent=2)
            
        return {
            'id': project_id,
            'name': name,
            'created_time': project_info['created_time'],
            'modified_time': project_info['modified_time'],
            'path': project_path
        }
        
    def delete_project(self, project_id):
        """删除项目"""
        project_path = os.path.join(self.projects_dir, project_id)
        if os.path.exists(project_path):
            shutil.rmtree(project_path)
            return True
        return False
        
    def load_project_data(self, project_id):
        """加载项目数据"""
        project_path = os.path.join(self.projects_dir, project_id)
        data_file = os.path.join(project_path, "project_data.json")
        
        if os.path.exists(data_file):
            try:
                with open(data_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"加载项目数据失败: {e}")
        return None
        
    def save_project_data(self, project_id, data):
        """保存项目数据"""
        project_path = os.path.join(self.projects_dir, project_id)
        data_file = os.path.join(project_path, "project_data.json")
        
        # 更新项目修改时间
        info_file = os.path.join(project_path, "project_info.json")
        if os.path.exists(info_file):
            try:
                with open(info_file, 'r', encoding='utf-8') as f:
                    project_info = json.load(f)
                project_info['modified_time'] = datetime.now().isoformat()
                with open(info_file, 'w', encoding='utf-8') as f:
                    json.dump(project_info, f, ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"更新项目信息失败: {e}")
        
        # 保存项目数据
        try:
            with open(data_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存项目数据失败: {e}")
            return False

class EnergyDataModel:
    def __init__(self):
        # 时序数据存储 (8760小时)
        self.electric_load_hourly = [0.0] * 8760  # 电力负荷
        self.heat_load_hourly = [0.0] * 8760      # 热力负荷
        self.internal_electric_rate = 0.0         # 厂用电率
        self.solar_irradiance_hourly = [0.0] * 8760  # 光照强度
        self.wind_speed_hourly = [0.0] * 8760     # 风速
        self.grid_purchase_price_hourly = [0.0] * 8760  # 下网电价
        
        # 数据导入状态跟踪
        self.data_imported = {
            'electric': False,
            'heat': False,
            'solar': False,
            'wind': False,
            'grid_price': False
        }
        
        # 风机型号列表，每个元素是一个字典，包含型号名称、参数、数量和计算方法
        self.wind_turbine_models = [
            {
                'name': '默认型号',
                'params': {
                    'cut_in_wind': 3.0,
                    'rated_wind': 12.0,
                    'max_rated_wind': 18.0,
                    'cut_out_wind': 25.0,
                    'rated_power': 2000.0
                },
                'count': 10,
                'output_correction_factor': 1.0  # 出力修正系数，默认値为1
            }
        ]
        
        # 光伏型号列表，每个元素是一个字典，包含型号名称、参数、数量和计算方法
        self.pv_models = [
            {
                'name': '默认型号',
                'method': 'area_efficiency',  # 计算方法：area_efficiency(面积效率法) 或 installed_capacity(装机容量法)
                'params': {
                    'panel_efficiency': 0.2,
                    'panel_area': 1000.0
                },
                'count': 10,  # 光伏板数量
                'output_correction_factor': 1.0  # 出力修正系数，默认値为1
            }
        ]
        
        self.chp_electric_params = {
            'electric_heat_ratio': 0.5,
            'base_electric': 100.0
        }
        
        # 新增调峰机组最小出力参数
        self.peak_power_min_summer = 0.0  # 夏季最小出力
        self.peak_power_min_winter = 0.0  # 冬季最小出力
        self.peak_power_max = 2000.0
        
        # 最大电力负荷参数
        self.max_electric_load = 5000.0
        
        # 灵活负荷参数
        self.flexible_load_max = 0.0  # 最大灵活负荷
        self.flexible_load_min = 0.0  # 最小灵活负荷
        
        # 检修和投产计划数据
        self.maintenance_schedules = []  # 检修计划列表
        self.commissioning_schedules = []  # 投产计划列表
        
        # 出力限制计划数据
        self.output_limit_schedules = []  # 出力限制计划列表
        
        # 优化参数
        self.optimization_params = {
            'basic_load_revenue': 1.0,      # 基础负荷单位收益 (元/kWh)
            'flexible_load_revenue': 0.8,   # 灵活负荷单位收益 (元/kWh)
            'thermal_cost': 0.2,            # 火电发电单位成本 (元/kWh)
            'pv_cost': 0.05,               # 光伏发电单位成本 (元/kWh)
            'wind_cost': 0.05,              # 风机发电单位成本 (元/kWh)
            'load_change_rate_limit': 100000.0,  # 负荷最大变动率 (kW/Hr)
            'min_grid_load': 0.0            # 最小下网负荷 (kW)
        }
        
    def calculate_wind_total_capacity(self):
        """
        计算风机总装机容量
        风机总装机容量 = 每个型号的装机容量相加得到
        其中每个型号的装机容量 = 额定功率 乘以 风机数量
        """
        total_capacity = 0.0
        for model in self.wind_turbine_models:
            rated_power = model['params'].get('rated_power', 0.0)
            count = model.get('count', 0)
            total_capacity += rated_power * count
        return total_capacity
        
    def calculate_pv_total_capacity(self):
        """
        计算光伏总装机容量
        光伏总装机容量 = 每个型号的装机容量相加得到
        其中面积效率法型号的装机容量 = 光伏板面积 * 光伏板效率 * 光伏板数量
        其中装机容量法型号的装机容量 = 装机容量 * 光伏板数量
        """
        total_capacity = 0.0
        for model in self.pv_models:
            method = model.get('method', 'area_efficiency')
            count = model.get('count', 0)
            
            if method == 'area_efficiency':
                # 面积效率法
                panel_area = model['params'].get('panel_area', 0.0)
                panel_efficiency = model['params'].get('panel_efficiency', 0.0)
                capacity = panel_area * panel_efficiency * count
            elif method == 'installed_capacity':
                # 装机容量法
                installed_capacity = model['params'].get('installed_capacity', 0.0)
                capacity = installed_capacity * count
            else:
                # 默认使用面积效率法
                panel_area = model['params'].get('panel_area', 0.0)
                panel_efficiency = model['params'].get('panel_efficiency', 0.0)
                capacity = panel_area * panel_efficiency * count
                
            total_capacity += capacity
        return total_capacity
        

    def to_dict(self):
        """将数据模型转换为字典，用于保存"""
        data = {
            'electric_load_hourly': self.electric_load_hourly,
            'heat_load_hourly': self.heat_load_hourly,
            'internal_electric_rate': self.internal_electric_rate,
            'solar_irradiance_hourly': self.solar_irradiance_hourly,
            'wind_speed_hourly': self.wind_speed_hourly,
            'grid_purchase_price_hourly': self.grid_purchase_price_hourly,  # 下网电价
            'data_imported': self.data_imported,
            'wind_turbine_models': self.wind_turbine_models,
            'pv_models': self.pv_models,
            'chp_electric_params': self.chp_electric_params,
            'peak_power_min_summer': self.peak_power_min_summer,
            'peak_power_min_winter': self.peak_power_min_winter,
            'peak_power_max': self.peak_power_max,
            'max_electric_load': self.max_electric_load,
            'flexible_load_max': self.flexible_load_max,
            'flexible_load_min': self.flexible_load_min,
            'maintenance_schedules': self.maintenance_schedules,
            'commissioning_schedules': self.commissioning_schedules,
            'output_limit_schedules': self.output_limit_schedules,
            'optimization_params': self.optimization_params,
            'optimized_results': getattr(self, 'optimized_results', None)
        }
        return data
        
    def from_dict(self, data):
        """从字典加载数据模型"""
        self.electric_load_hourly = data.get('electric_load_hourly', [0.0] * 8760)
        self.heat_load_hourly = data.get('heat_load_hourly', [0.0] * 8760)
        self.internal_electric_rate = data.get('internal_electric_rate', 0.0)
        self.solar_irradiance_hourly = data.get('solar_irradiance_hourly', [0.0] * 8760)
        self.wind_speed_hourly = data.get('wind_speed_hourly', [0.0] * 8760)
        self.grid_purchase_price_hourly = data.get('grid_purchase_price_hourly', [0.0] * 8760)  # 下网电价
        self.data_imported = data.get('data_imported', {
            'electric': False,
            'heat': False,
            'solar': False,
            'wind': False,
            'grid_price': False
        })
        self.wind_turbine_models = data.get('wind_turbine_models', [
            {
                'name': '默认型号',
                'params': {
                    'cut_in_wind': 3.0,
                    'rated_wind': 12.0,
                    'max_rated_wind': 18.0,
                    'cut_out_wind': 25.0,
                    'rated_power': 2000.0
                },
                'count': 10,
                'output_correction_factor': 1.0  # 出力修正系数，默认値为1
            }
        ])
        self.pv_models = data.get('pv_models', [
            {
                'name': '默认型号',
                'method': 'area_efficiency',
                'params': {
                    'panel_efficiency': 0.2,
                    'panel_area': 1000.0
                },
                'count': 10,
                'output_correction_factor': 1.0  # 出力修正系数，默认値为1
            }
        ])
        self.chp_electric_params = data.get('chp_electric_params', {
            'electric_heat_ratio': 0.5,
            'base_electric': 100.0
        })
        self.peak_power_min_summer = data.get('peak_power_min_summer', 0.0)  # 夏季最小出力
        self.peak_power_min_winter = data.get('peak_power_min_winter', 0.0)  # 冬季最小出力
        self.peak_power_max = data.get('peak_power_max', 2000.0)
        self.max_electric_load = data.get('max_electric_load', 5000.0)
        self.flexible_load_max = data.get('flexible_load_max', 0.0)  # 最大灵活负荷
        self.flexible_load_min = data.get('flexible_load_min', 0.0)  # 最小灵活负荷
        
        # 加载检修和投产计划数据
        self.maintenance_schedules = data.get('maintenance_schedules', [])
        self.commissioning_schedules = data.get('commissioning_schedules', [])
        
        # 加载出力限制计划数据
        self.output_limit_schedules = data.get('output_limit_schedules', [])
        
        # 加载优化参数
        self.optimization_params = data.get('optimization_params', {
            'basic_load_revenue': 1.0,
            'flexible_load_revenue': 0.8,
            'thermal_cost': 0.2,
            'pv_cost': 0.05,
            'wind_cost': 0.05,
            'load_change_rate_limit': 100000.0,
            'min_grid_load': 0.0
        })
        
        # 加载优化结果（如果有）
        if 'optimized_results' in data and data['optimized_results'] is not None:
            self.optimized_results = data['optimized_results']
        
        # 加载计算结果（如果有）
        calculation_results = data.get('calculation_results')
        return calculation_results

# 风机出力与风速函数关系
def wind_power_function(wind_speed, params, correction_factor=1.0):
    """
    风机出力计算函数（修改版）
    1) 在切入风速和设计风速之间采用2次曲线拟合，并在接近设计风速时平滑过渡
    2) 增加一个最大额定风速，在额定风速和最大额定风速之间为额定功率，
       在最大额定风速和切出风速之间使用线性函数
    3) 应用出力修正系数
    """
    cut_in_wind = params.get('cut_in_wind', 3.0)           # 切入风速
    rated_wind = params.get('rated_wind', 12.0)             # 额定风速
    max_rated_wind = params.get('max_rated_wind', 18.0)     # 最大额定风速
    cut_out_wind = params.get('cut_out_wind', 25.0)         # 切出风速
    rated_power = params.get('rated_power', 2000.0)         # 额定功率
    
    # 风速低于切入风速或高于切出风速时，出力为0
    if wind_speed < cut_in_wind or wind_speed > cut_out_wind:
        return 0.0
    
    # 在切入风速和额定风速之间，采用二次曲线拟合
    elif wind_speed < rated_wind:
        # 二次曲线拟合，确保在额定风速处平滑过渡
        # 使用抛物线方程: P = a * (v - v_in)^2
        # 约束条件: 在额定风速处达到额定功率
        a = rated_power / ((rated_wind - cut_in_wind) ** 2)
        return a * (wind_speed - cut_in_wind) ** 2 * correction_factor
    
    # 在额定风速和最大额定风速之间，保持额定功率
    elif wind_speed < max_rated_wind:
        return rated_power * correction_factor
    
    # 在最大额定风速和切出风速之间，使用线性递减函数
    else:
        # 线性递减从额定功率到0
        slope = rated_power / (cut_out_wind - max_rated_wind)
        return (rated_power - slope * (wind_speed - max_rated_wind)) * correction_factor

def total_wind_power_function(wind_speed, turbine_models):
    """
    计算所有风机型号的总出力
    :param wind_speed: 风速
    :param turbine_models: 风机型号列表，每个元素包含参数和数量
    :return: 总出力
    """
    total_power = 0.0
    for model in turbine_models:
        # 计算单台风机出力，应用修正系数
        single_power = wind_power_function(wind_speed, model['params'], model.get('output_correction_factor', 1.0))
        # 乘以该型号风机数量
        total_power += single_power * model['count']
    return total_power

# 光伏出力与光照强度函数关系
def pv_power_function(irradiance, model):
    """
    光伏出力计算函数（支持两种计算方法）
    :param irradiance: 光照强度 (W/m²)
    :param model: 光伏型号参数
    :return: 光伏出力 (kW)
    """
    method = model.get('method', 'area_efficiency')
    params = model.get('params', {})
    correction_factor = model.get('output_correction_factor', 1.0)  # 获取出力修正系数，默认为1
    
    if method == 'area_efficiency':
        # 面积效率法
        panel_efficiency = params.get('panel_efficiency', 0.2)
        panel_area = params.get('panel_area', 1000.0)
        return panel_area * irradiance * panel_efficiency / 1000.0 * correction_factor
    elif method == 'installed_capacity':
        # 装机容量法
        installed_capacity = params.get('installed_capacity', 200.0)  # 装机容量 (kW)
        system_efficiency = params.get('system_efficiency', 0.9)    # 系统效率
        return irradiance / 1000.0 * installed_capacity * system_efficiency * correction_factor
    else:
        # 默认使用面积效率法
        panel_efficiency = params.get('panel_efficiency', 0.2)
        panel_area = params.get('panel_area', 1000.0)
        return panel_area * irradiance * panel_efficiency / 1000.0 * correction_factor

def total_pv_power_function(irradiance, pv_models):
    """
    计算所有光伏型号的总出力
    :param irradiance: 光照强度 (W/m²)
    :param pv_models: 光伏型号列表，每个元素包含参数和数量
    :return: 总出力 (kW)
    """
    total_power = 0.0
    for model in pv_models:
        # 计算单个光伏型号出力
        single_power = pv_power_function(irradiance, model)
        # 乘以该型号光伏板数量
        total_power += single_power * model.get('count', 1)
    return total_power

# 热电联产电出力与供热热负荷函数关系
def chp_electric_power(heat_load, params):
    """
    热电联产电出力计算函数
    """
    electric_heat_ratio = params.get('electric_heat_ratio', 0.5)
    base_electric = params.get('base_electric', 100.0)
    
    return base_electric + heat_load * electric_heat_ratio

class AnnualBalanceCalculator:
    def __init__(self, data_model):
        self.data_model = data_model
        
    def is_date_in_range(self, date_str, start_date_str, end_date_str):
        """
        检查给定日期是否在指定范围内（包含起止日期）
        :param date_str: 检查的日期字符串 (YYYY-MM-DD)
        :param start_date_str: 起始日期字符串 (YYYY-MM-DD)
        :param end_date_str: 结束日期字符串 (YYYY-MM-DD)
        :return: 是否在范围内
        """
        try:
            from datetime import datetime
            date = datetime.strptime(date_str, "%Y-%m-%d")
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
            return start_date <= date <= end_date
        except ValueError:
            # 如果日期解析失败，返回False
            return False
    
    def calculate_interpolation_factor(self, hour, start_date_str, end_date_str):
        """
        计算线性插值因子
        :param hour: 小时索引 (0-8759)
        :param start_date_str: 起始日期字符串 (YYYY-MM-DD)
        :param end_date_str: 结束日期字符串 (YYYY-MM-DD)
        :return: 插值因子 (0.0-1.0)
        """
        try:
            from datetime import datetime, timedelta
            # 计算当前日期
            base_date = datetime(2024, 1, 1)
            current_date = base_date + timedelta(hours=hour)
            current_date_str = current_date.strftime("%Y-%m-%d")
            
            # 如果在开始日期之前，因子为0
            if current_date < datetime.strptime(start_date_str, "%Y-%m-%d"):
                return 0.0
                
            # 如果在结束日期之后，因子为1
            if current_date > datetime.strptime(end_date_str, "%Y-%m-%d"):
                return 1.0
                
            # 计算在范围内的位置
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
            
            total_days = (end_date - start_date).days
            elapsed_days = (current_date - start_date).days
            
            if total_days == 0:
                return 1.0
                
            return elapsed_days / total_days
            
        except ValueError:
            # 如果日期解析失败，返回0
            return 0.0
    
    def get_active_maintenance_schedules(self, hour):
        """
        获取在指定小时处于活动状态的检修计划
        :param hour: 小时索引 (0-8759)
        :return: 活动的检修计划列表
        """
        # 计算日期 (假设从2024年1月1日开始)
        from datetime import datetime, timedelta
        base_date = datetime(2024, 1, 1)
        current_date = base_date + timedelta(hours=hour)
        current_date_str = current_date.strftime("%Y-%m-%d")
        
        active_schedules = []
        for schedule in self.data_model.maintenance_schedules:
            if self.is_date_in_range(current_date_str, 
                                   schedule.get('start_date', ''), 
                                   schedule.get('end_date', '')):
                active_schedules.append(schedule)
                
        return active_schedules
    
    def get_active_commissioning_schedules(self, hour):
        """
        获取在指定小时处于活动状态的投产计划
        对于投产计划，不仅在计划日期范围内需要激活，在计划开始日期之前也需要激活（用于投产前修正）
        :param hour: 小时索引 (0-8759)
        :return: 活动的投产计划列表
        """
        # 计算日期 (假设从2024年1月1日开始)
        from datetime import datetime, timedelta
        base_date = datetime(2024, 1, 1)
        current_date = base_date + timedelta(hours=hour)
        current_date_str = current_date.strftime("%Y-%m-%d")
        
        active_schedules = []
        for schedule in self.data_model.commissioning_schedules:
            start_date_str = schedule.get('start_date', '')
            end_date_str = schedule.get('end_date', '')
            
            # 检查当前日期是否在投产计划的起止日期范围内，或者在起始日期之前
            # 如果在开始日期之前，也要激活计划（用于投产前修正）
            try:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
                if current_date <= start_date or self.is_date_in_range(current_date_str, start_date_str, end_date_str):
                    active_schedules.append(schedule)
            except ValueError:
                # 如果日期解析失败，使用原来的逻辑
                if self.is_date_in_range(current_date_str, start_date_str, end_date_str):
                    active_schedules.append(schedule)
                
        return active_schedules
    
    def get_active_output_limit_schedules(self, hour, limit_type=None):
        """
        获取在指定小时处于活动状态的出力限制计划
        :param hour: 小时索引 (0-8759)
        :param limit_type: 限制类型，如果为None则返回所有类型的限制计划
        :return: 活动的出力限制计划列表
        """
        # 计算日期 (假设从2024年1月1日开始)
        from datetime import datetime
        base_date = datetime(2024, 1, 1)
        current_date = base_date + timedelta(hours=hour)
        current_date_str = current_date.strftime("%Y-%m-%d")
        
        active_schedules = []
        for schedule in self.data_model.output_limit_schedules:
            # 如果指定了限制类型，则只返回该类型的计划
            if limit_type and schedule.get('limit_type', '') != limit_type:
                continue
                
            if self.is_date_in_range(current_date_str, 
                                   schedule.get('start_date', ''), 
                                   schedule.get('end_date', '')):
                active_schedules.append(schedule)
                
        return active_schedules
    
    def calculate_annual_balance(self):
        """
        计算年度8760小时的能源平衡
        根据新公式计算各项参数
        """
        results = {
            'hourly_internal_electric_load': [0.0] * 8760,  # 厂用电负荷
            'hourly_total_load': [0.0] * 8760,              # 总负荷
            'hourly_chp_output': [0.0] * 8760,              # 热定电机组出力
            'hourly_pv_output': [0.0] * 8760,               # 光伏最大出力
            'hourly_wind_output': [0.0] * 8760,             # 风机最大出力
            'hourly_peak_pending_output': [0.0] * 8760,     # 调峰机组待定出力
            'hourly_peak_output': [0.0] * 8760,             # 调峰机组出力
            'hourly_thermal_output': [0.0] * 8760,          # 火电出力
            'hourly_generation': [0.0] * 8760,              # 总出力
            'hourly_wind_pv_abandon': [0.0] * 8760,         # 风机光伏放弃出力
            'hourly_wind_pv_actual': [0.0] * 8760,          # 风机光伏实际出力
            'hourly_grid_load': [0.0] * 8760,               # 下网负荷
            'hourly_abandon_rate': [0.0] * 8760,            # 弃光风率
            'hourly_corrected_electric_load': [0.0] * 8760,  # 修正后电力负荷
            'hourly_flexible_load_consumption': [0.0] * 8760  # 灵活负荷消纳量
        }
        
        # 保存原始参数用于检修修正计算
        original_peak_power_max = self.data_model.peak_power_max
        original_peak_power_min_summer = self.data_model.peak_power_min_summer  # 夏季最小出力
        original_peak_power_min_winter = self.data_model.peak_power_min_winter  # 冬季最小出力
        original_electric_load = list(self.data_model.electric_load_hourly)  # 复制一份
        
        for hour in range(8760):
            # 获取当前小时的活动检修计划和投产计划
            active_maintenance_schedules = self.get_active_maintenance_schedules(hour)
            active_commissioning_schedules = self.get_active_commissioning_schedules(hour)
            
            # 初始化当前小时的参数
            current_peak_power_max = original_peak_power_max
            
            # 根据月份确定当前应该使用的最小出力
            from datetime import datetime, timedelta
            base_date = datetime(2024, 1, 1)
            current_date = base_date + timedelta(hours=hour)
            current_month = current_date.month
            
            if 5 <= current_month <= 9:  # 夏季：5-9月
                current_peak_power_min = original_peak_power_min_summer
            else:  # 冬季：10-12月和1-4月
                current_peak_power_min = original_peak_power_min_winter
                
            current_electric_load = original_electric_load[hour]
            current_max_load = max(original_electric_load) if original_electric_load else 1.0  # 防止除零错误
            
            # 计算检修计划和投运计划对用电负荷的总影响
            maintenance_load_reduction = 0.0  # 检修计划对用电负荷的减少量
            commissioning_load_reduction = 0.0  # 投运计划对用电负荷的减少量（开始前）或增加量（结束后）
            
            # 收集检修计划对用电负荷的影响
            for schedule in active_maintenance_schedules:
                power_type = schedule.get('power_type', '')
                power_size = schedule.get('power_size', 0.0)
                
                if power_type == '调峰机组出力':
                    # 新最大负荷 = 原最大负荷 - 影响负荷出力大小
                    current_peak_power_max = current_peak_power_max - power_size
                    # 新最小负荷 = 原最小负荷 * （新最大负荷/原最大负荷）
                    if original_peak_power_max > 0:
                        # 分别调整夏季和冬季最小出力
                        if 5 <= current_date.month <= 9:  # 夏季
                            current_peak_power_min = original_peak_power_min_summer * (current_peak_power_max / original_peak_power_max)
                        else:  # 冬季
                            current_peak_power_min = original_peak_power_min_winter * (current_peak_power_max / original_peak_power_max)
                        
                elif power_type == '用电负荷':
                    maintenance_load_reduction += power_size
            
            # 收集投运计划对用电负荷的影响（线性变化）
            for schedule in active_commissioning_schedules:
                power_type = schedule.get('power_type', '')
                power_size = schedule.get('power_size', 0.0)
                start_date = schedule.get('start_date', '')
                end_date = schedule.get('end_date', '')
                
                # 计算线性插值因子
                interpolation_factor = self.calculate_interpolation_factor(hour, start_date, end_date)
                
                if power_type == '光伏出力':
                    # 在投产开始日期前，其最大出力修正为原最大出力减去设置的影响出力负荷大小
                    # 在投产结束日后，最大出力即为原最大出力
                    # 在投产期间，采用线性变化
                    pass  # 光伏出力的修正将在计算光伏出力时处理
                    
                elif power_type == '风机出力':
                    # 在投产开始日期前，其最大出力修正为原最大出力减去设置的影响出力负荷大小
                    # 在投产结束日后，最大出力即为原最大出力
                    # 在投产期间，采用线性变化
                    pass  # 风机出力的修正将在计算风机出力时处理
                    
                elif power_type == '调峰机组最大出力':
                    # 在投产开始日期前，修正的调峰机组最大负荷为原调峰机组最大负荷减去设置的影响出力负荷大小
                    # 在投产结束日后，修正的调峰机组最大负荷即为原调峰机组最大负荷
                    # 在投产期间，采用线性变化。即最大负荷随着投产进行逐步增大
                    adjusted_power_size = power_size * interpolation_factor
                    current_peak_power_max = current_peak_power_max - (power_size - adjusted_power_size)
                    
                elif power_type == '调峰机组夏季最小出力':
                    # 在投产开始日期前，修正的调峰机组夏季最小负荷为原调峰机组夏季最小负荷
                    # 在投产结束日后，修正的调峰机组夏季最小负荷即为原调峰机组夏季最小负荷减去设置的影响出力负荷大小
                    # 在投产期间，采用线性变化。即最小负荷随着投产进行逐步减小
                    adjusted_power_size = power_size * interpolation_factor
                    if 5 <= current_date.month <= 9:  # 夏季：5-9月
                        current_peak_power_min = original_peak_power_min_summer - adjusted_power_size
                elif power_type == '调峰机组冬季最小出力':
                    # 在投产开始日期前，修正的调峰机组冬季最小负荷为原调峰机组冬季最小负荷
                    # 在投产结束日后，修正的调峰机组冬季最小负荷即为原调峰机组冬季最小负荷减去设置的影响出力负荷大小
                    # 在投产期间，采用线性变化。即最小负荷随着投产进行逐步减小
                    adjusted_power_size = power_size * interpolation_factor
                    if current_date.month < 5 or current_date.month > 9:  # 冬季：10-12月和1-4月
                        current_peak_power_min = original_peak_power_min_winter - adjusted_power_size
                elif power_type == '调峰机组最小出力':
                    # 为了向后兼容，仍然支持原有的调峰机组最小出力设置
                    # 在投产开始日期前，修正的调峰机组最小负荷为原调峰机组最小负荷
                    # 在投产结束日后，修正的调峰机组最小负荷即为原调峰机组最小负荷减去设置的影响出力负荷大小
                    # 在投产期间，采用线性变化。即最小负荷随着投产进行逐步减小
                    adjusted_power_size = power_size * interpolation_factor
                    if 5 <= current_date.month <= 9:  # 夏季
                        current_peak_power_min = original_peak_power_min_summer - adjusted_power_size
                    else:  # 冬季
                        current_peak_power_min = original_peak_power_min_winter - adjusted_power_size
                    
                elif power_type == '用电负荷':
                    # 投运计划对用电负荷的影响是渐变的：开始前是全部影响，结束后是无影响，中间线性变化
                    adjusted_power_size = power_size * (1 - interpolation_factor)  # 1-interpolation_factor 表示剩余影响
                    commissioning_load_reduction += adjusted_power_size
            
            # 应用检修计划和投运计划对用电负荷的综合影响
            # 用电负荷 = 原始用电负荷 / 最大电力负荷 * (最大电力负荷 - 检修影响 - 投运影响)
            if current_max_load > 0:
                current_electric_load = current_electric_load / current_max_load * (current_max_load - maintenance_load_reduction - commissioning_load_reduction)
            
            # 保存修正后的电力负荷，以便在输出中显示
            results['hourly_corrected_electric_load'][hour] = current_electric_load
            
            # 使用迭代方法计算厂用电负荷和总负荷，解决循环依赖问题
            # 初始化：先使用传统方法计算初始值
            internal_electric_load = current_electric_load * self.data_model.internal_electric_rate
            total_load = current_electric_load + internal_electric_load
            
            # 迭代计算，直到收敛或达到最大迭代次数
            max_iterations = 10  # 最大迭代次数
            tolerance = 1e-6     # 收敛阈值
            
            for iteration in range(max_iterations):
                # 保存上一次的值用于比较
                prev_internal_electric_load = internal_electric_load
                prev_total_load = total_load
                
                # 3) 热定电机组出力 = 热力负荷 * 电热比
                chp_output = (
                    self.data_model.heat_load_hourly[hour] * 
                    self.data_model.chp_electric_params['electric_heat_ratio']
                )
                
                # 4) 光伏最大出力 = 根据所有光伏型号计算总出力
                pv_output = total_pv_power_function(
                    self.data_model.solar_irradiance_hourly[hour],
                    self.data_model.pv_models
                )
                
                # 收集所有光伏出力相关的投产计划
                pv_schedules = []
                for schedule in active_commissioning_schedules:
                    if schedule.get('power_type', '') == '光伏出力':
                        pv_schedules.append(schedule)
                
                # 应用出力限制计划对光伏出力的修正
                active_pv_limit_schedules = self.get_active_output_limit_schedules(hour, '光伏最大出力限制')
                if active_pv_limit_schedules:
                    # 如果有多个限制计划，取最小的限制值
                    min_limit = min([schedule.get('power_size', float('inf')) for schedule in active_pv_limit_schedules])
                    # 限制光伏出力不超过设定值
                    pv_output = min(pv_output, min_limit)
                
                # 应用投产计划对光伏出力的修正（支持多个计划叠加）
                if pv_schedules:
                    # 获取光伏总装机容量
                    total_pv_capacity = self.data_model.calculate_pv_total_capacity()
                    if total_pv_capacity > 0:
                        # 计算累计影响
                        cumulative_impact = 0.0
                        for schedule in pv_schedules:
                            power_size = schedule.get('power_size', 0.0)
                            start_date = schedule.get('start_date', '')
                            end_date = schedule.get('end_date', '')
                            
                            # 计算线性插值因子
                            interpolation_factor = self.calculate_interpolation_factor(hour, start_date, end_date)
                            
                            # 累计影响（投产前为全部影响，投产后为0影响）
                            if interpolation_factor < 1:  # 投产前或投产中
                                # 投产前影响是全部，投产中是部分影响
                                impact = power_size * (1 - interpolation_factor)
                                cumulative_impact += impact
                        
                        # 应用累计修正
                        adjusted_output_factor = (total_pv_capacity - cumulative_impact) / total_pv_capacity
                        pv_output = pv_output * adjusted_output_factor
                
                # 5) 风机最大出力 = 根据所有风机型号计算总出力
                wind_output = total_wind_power_function(
                    self.data_model.wind_speed_hourly[hour],
                    self.data_model.wind_turbine_models
                )
                
                # 收集所有风机出力相关的投产计划
                wind_schedules = []
                for schedule in active_commissioning_schedules:
                    if schedule.get('power_type', '') == '风机出力':
                        wind_schedules.append(schedule)
                
                # 应用出力限制计划对风机出力的修正
                active_wind_limit_schedules = self.get_active_output_limit_schedules(hour, '风机最大出力限制')
                if active_wind_limit_schedules:
                    # 如果有多个限制计划，取最小的限制值
                    min_limit = min([schedule.get('power_size', float('inf')) for schedule in active_wind_limit_schedules])
                    # 限制风机出力不超过设定值
                    wind_output = min(wind_output, min_limit)
                
                # 应用投产计划对风机出力的修正（支持多个计划叠加）
                if wind_schedules:
                    # 获取风机总装机容量
                    total_wind_capacity = self.data_model.calculate_wind_total_capacity()
                    if total_wind_capacity > 0:
                        # 计算累计影响
                        cumulative_impact = 0.0
                        for schedule in wind_schedules:
                            power_size = schedule.get('power_size', 0.0)
                            start_date = schedule.get('start_date', '')
                            end_date = schedule.get('end_date', '')
                            
                            # 计算线性插值因子
                            interpolation_factor = self.calculate_interpolation_factor(hour, start_date, end_date)
                            
                            # 累计影响（投产前为全部影响，投产后为0影响）
                            if interpolation_factor < 1:  # 投产前或投产中
                                # 投产前影响是全部，投产中是部分影响
                                impact = power_size * (1 - interpolation_factor)
                                cumulative_impact += impact
                        
                        # 应用累计修正
                        adjusted_output_factor = (total_wind_capacity - cumulative_impact) / total_wind_capacity
                        wind_output = wind_output * adjusted_output_factor
                
                # 6) 调峰机组待定出力 = 总负荷 - 热定电机组出力 - 光伏最大出力 - 风机最大出力
                peak_pending_output = total_load - chp_output - pv_output - wind_output
                
                # 7) 调峰机组出力 = max(min(调峰机组待定出力, 调峰机组最大出力），调峰机组最小出力）
                peak_output = max(
                    min(peak_pending_output, current_peak_power_max),
                    current_peak_power_min
                )
                
                # 8) 火电出力 = 热定电机组出力 + 调峰机组出力
                thermal_output = chp_output + peak_output
                
                # 使用新的公式计算厂用电负荷: 厂用电负荷 = 火电出力 * 厂用电率
                internal_electric_load = thermal_output * self.data_model.internal_electric_rate
                
                # 重新计算总负荷，使用修正后的电力负荷
                total_load = current_electric_load + internal_electric_load
                
                # 检查收敛性
                load_diff = abs(total_load - prev_total_load)
                if load_diff < tolerance:
                    # 已收敛，跳出迭代循环
                    break
            
            # 9) 风机光伏放弃出力 = max（当前月份对应的调峰机组最小出力 - 调峰机组待定出力，0）
            wind_pv_abandon = max(current_peak_power_min - peak_pending_output, 0)
            
            # 10) 应用灵活负荷消纳逻辑
            # 2-1）当光伏及风力放弃出力小于最小灵活负荷时，灵活负荷不启动，不消纳；
            # 2-2）当光伏及风力放弃出力大于等于最小灵活负荷时，小于等于最大灵活负荷时，所有放弃出力都被灵活负荷消纳。
            # 2-3) 当光伏及风力放弃出力大于最大灵活负荷时，新消纳的出力等于最大灵活负荷，剩余部分出力继续放弃。
            flexible_load_max = self.data_model.flexible_load_max
            flexible_load_min = self.data_model.flexible_load_min
            
            if wind_pv_abandon >= flexible_load_min:
                # 当放弃出力大于等于最小灵活负荷时，启动灵活负荷消纳
                if wind_pv_abandon <= flexible_load_max:
                    # 当放弃出力小于等于最大灵活负荷时，全部消纳
                    actual_flexible_load = wind_pv_abandon
                else:
                    # 当放弃出力大于最大灵活负荷时，只消纳最大灵活负荷
                    actual_flexible_load = flexible_load_max
            else:
                # 当放弃出力小于最小灵活负荷时，不启动灵活负荷
                actual_flexible_load = 0.0
            
            # 11) 修正计算结果中风机光伏放弃出力
            # 新的风机光伏放弃出力 = 原风机光伏放弃出力 - 新增的灵活负荷消纳出力
            corrected_wind_pv_abandon = max(wind_pv_abandon - actual_flexible_load, 0)  # 确保不为负数
            
            # 10) 新的光伏风电实际出力 = （光伏最大出力+风机最大出力）- 新的风机光伏放弃出力
            wind_pv_actual = (pv_output + wind_output) - corrected_wind_pv_abandon
            
            # 11) 总出力 = 光伏出力 + 风机出力 + 火电出力 - 新的风机光伏放弃出力
            generation = pv_output + wind_output + thermal_output - corrected_wind_pv_abandon
            
            # 12) 弃光风率 = 新的风机光伏放弃出力 / (光伏最大出力 + 风机最大出力)，避免除零错误
            abandon_rate = 0
            if (pv_output + wind_output) > 0:
                abandon_rate = corrected_wind_pv_abandon / (pv_output + wind_output)
            else:
                abandon_rate = 0
            
            # 13) 修正下网负荷计算
            # 下网负荷 = 总负荷 + 新增的灵活负荷消纳出力 - 总出力
            grid_load = total_load + actual_flexible_load - generation
            
            # 记录灵活负荷消纳量
            results['hourly_flexible_load_consumption'][hour] = actual_flexible_load
            
            # 存储修正后的风机光伏放弃出力和实际出力
            results['hourly_wind_pv_abandon'][hour] = corrected_wind_pv_abandon
            results['hourly_wind_pv_actual'][hour] = wind_pv_actual
            
            # 存储结果
            results['hourly_internal_electric_load'][hour] = internal_electric_load
            results['hourly_total_load'][hour] = total_load
            results['hourly_chp_output'][hour] = chp_output
            results['hourly_pv_output'][hour] = pv_output
            results['hourly_wind_output'][hour] = wind_output
            results['hourly_peak_pending_output'][hour] = peak_pending_output
            results['hourly_peak_output'][hour] = peak_output
            results['hourly_thermal_output'][hour] = thermal_output
            results['hourly_generation'][hour] = generation
            results['hourly_wind_pv_abandon'][hour] = wind_pv_abandon
            results['hourly_wind_pv_actual'][hour] = wind_pv_actual
            results['hourly_grid_load'][hour] = grid_load
            results['hourly_abandon_rate'][hour] = abandon_rate
            
        return results

class EnergyBalanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("园区用电用热负荷与出力平衡计算系统")
        # 设置默认最大化显示，保留标题栏
        self.root.state('zoomed')  # Windows系统下的最大化方法
        self.root.bind('<Escape>', self.exit_fullscreen)  # 绑定ESC键退出全屏
        
        # 初始化项目管理器
        # 检查是否在打包环境中运行
        if getattr(sys, 'frozen', False):
            # 如果是打包后的可执行文件，使用可执行文件所在目录
            self.app_path = os.path.dirname(sys.executable)
        else:
            # 如果是直接运行Python脚本，使用脚本所在目录
            self.app_path = os.path.dirname(os.path.abspath(__file__))
        
        self.project_manager = ProjectManager(self.app_path)
        self.current_project = None
        
        # 初始化数据模型
        self.data_model = EnergyDataModel()
        self.calculator = AnnualBalanceCalculator(self.data_model)
        self.results = None
        
        # 初始化图表交互变量
        self.pan_mode = False
        self.zoom_mode = False
        
        # 创建UI
        self.create_project_management_ui()
        
    def exit_fullscreen(self, event=None):
        """退出全屏模式"""
        self.root.state('normal')
        
    def create_project_management_ui(self):
        """创建项目管理界面"""
        # 清除主窗口内容
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # 创建项目管理界面
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="项目管理", font=("微软雅黑", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # 新建项目按钮
        new_project_btn = ttk.Button(main_frame, text="新建项目", command=self.create_new_project)
        new_project_btn.pack(pady=(0, 20))
        
        # 项目列表
        projects_frame = ttk.LabelFrame(main_frame, text="项目列表", padding="10")
        projects_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建表格来显示项目列表
        columns = ('项目名称', '创建时间', '修改时间')
        self.projects_tree = ttk.Treeview(projects_frame, columns=columns, show='headings', height=15)
        
        # 定义列标题
        self.projects_tree.heading('项目名称', text='项目名称')
        self.projects_tree.heading('创建时间', text='创建时间')
        self.projects_tree.heading('修改时间', text='修改时间')
        
        # 定义列宽度
        self.projects_tree.column('项目名称', width=200)
        self.projects_tree.column('创建时间', width=150)
        self.projects_tree.column('修改时间', width=150)
        
        # 添加滚动条
        scrollbar_y = ttk.Scrollbar(projects_frame, orient=tk.VERTICAL, command=self.projects_tree.yview)
        scrollbar_x = ttk.Scrollbar(projects_frame, orient=tk.HORIZONTAL, command=self.projects_tree.xview)
        self.projects_tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # 布局
        self.projects_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 操作按钮
        btn_frame = ttk.Frame(projects_frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(btn_frame, text="打开项目", command=self.open_selected_project).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="删除项目", command=self.delete_selected_project).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="复制项目", command=self.copy_selected_project).pack(side=tk.LEFT, padx=5)
        
        # 配置网格权重
        projects_frame.columnconfigure(0, weight=1)
        projects_frame.rowconfigure(0, weight=1)
        
        # 加载项目列表
        self.load_project_list()
        
    def load_project_list(self):
        """加载项目列表"""
        # 清空现有项目列表
        for item in self.projects_tree.get_children():
            self.projects_tree.delete(item)
        
        # 获取项目列表
        projects = self.project_manager.get_project_list()
        
        # 添加到Treeview
        for project in projects:
            created_time = datetime.fromisoformat(project['created_time']).strftime('%Y-%m-%d %H:%M')
            modified_time = datetime.fromisoformat(project['modified_time']).strftime('%Y-%m-%d %H:%M')
            self.projects_tree.insert('', tk.END, values=(project['name'], created_time, modified_time),
                                     tags=(project['id'],))
        
        # 绑定双击事件到项目名称列，用于编辑项目名称
        self.projects_tree.bind('<Double-1>', self.on_project_name_double_click)
        
    def on_project_name_double_click(self, event):
        """
        处理项目名称列双击事件，用于编辑项目名称
        """
        # 获取被点击的项目行
        item = self.projects_tree.identify_row(event.y)
        column = self.projects_tree.identify_column(event.x)
        
        # 只有在双击项目名称列时才执行编辑
        if item and column == '#1':  # 第一列是项目名称列
            # 获取项目信息
            item_values = self.projects_tree.item(item, 'values')
            project_id = self.projects_tree.item(item, 'tags')[0]
            project_name = item_values[0]
            
            # 执行编辑项目名称
            self.edit_project_name(item, project_id, project_name)
        
    def edit_project_name(self, item, project_id, current_name):
        """
        编辑项目名称
        """
        # 创建Entry控件用于编辑
        entry = ttk.Entry(self.projects_tree)
        entry.insert(0, current_name)
        entry.select_range(0, tk.END)
        entry.focus()
        
        # 获取项目名称列的位置
        bbox = self.projects_tree.bbox(item, '#1')
        if bbox:
            x, y, width, height = bbox
            entry.place(x=x, y=y, width=width)
        
        def save_name():
            new_name = entry.get().strip()
            entry.destroy()
            
            if new_name and new_name != current_name:
                # 检查新名称是否与其他项目重名
                all_projects = self.project_manager.get_project_list()
                for proj in all_projects:
                    if proj['name'] == new_name and proj['id'] != project_id:
                        messagebox.showwarning("警告", f"项目名称 '{new_name}' 已存在，请选择其他名称！")
                        return
                
                # 更新项目信息
                try:
                    # 加载项目信息文件
                    project_info_path = os.path.join(self.project_manager.projects_dir, project_id, "project_info.json")
                    if os.path.exists(project_info_path):
                        with open(project_info_path, 'r', encoding='utf-8') as f:
                            project_info = json.load(f)
                        
                        # 更新项目名称
                        project_info['name'] = new_name
                        
                        # 保存项目信息文件
                        with open(project_info_path, 'w', encoding='utf-8') as f:
                            json.dump(project_info, f, ensure_ascii=False, indent=2)
                        
                        # 更新Treeview显示
                        values = list(self.projects_tree.item(item, 'values'))
                        values[0] = new_name
                        self.projects_tree.item(item, values=values)
                        
                        messagebox.showinfo("成功", "项目名称已更新！")
                    else:
                        messagebox.showerror("错误", "项目信息文件不存在！")
                except Exception as e:
                    messagebox.showerror("错误", f"更新项目名称时发生错误:\n{str(e)}")
            
        def cancel_edit():
            entry.destroy()
        
        # 绑定回车键和失去焦点事件
        entry.bind('<Return>', lambda e: save_name())
        entry.bind('<FocusOut>', lambda e: save_name())  # 当Entry失去焦点时也保存
        entry.bind('<Escape>', lambda e: cancel_edit())
        
    def create_new_project(self):
        """创建新项目"""
        # 弹出对话框输入项目名称
        dialog = tk.Toplevel(self.root)
        dialog.title("新建项目")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="项目名称:").pack(pady=10)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        def confirm():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("警告", "请输入项目名称！")
                return
            
            # 创建项目
            project = self.project_manager.create_project(name)
            self.current_project = project
            
            # 关闭对话框
            dialog.destroy()
            
            # 进入主应用界面
            self.enter_main_app()
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="确定", command=confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键
        name_entry.bind('<Return>', lambda e: confirm())
        
    def open_selected_project(self):
        """打开选中的项目"""
        selection = self.projects_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个项目！")
            return
        
        # 获取选中项目的ID
        item = self.projects_tree.item(selection[0])
        project_id = self.projects_tree.item(selection[0], 'tags')[0]
        
        # 加载项目
        project_data = self.project_manager.load_project_data(project_id)
        if project_data is not None:  # 修改判断条件，允许空数据
            # 更新当前项目
            self.current_project = {
                'id': project_id,
                'name': item['values'][0],
                'path': os.path.join(self.project_manager.projects_dir, project_id)
            }
            
            # 从项目数据恢复数据模型
            self.data_model.from_dict(project_data)
            
            # 进入主应用界面
            self.enter_main_app()
        else:
            messagebox.showerror("错误", "无法加载项目数据！")
            
    def delete_selected_project(self):
        """删除选中的项目"""
        selection = self.projects_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个项目！")
            return
        
        # 获取选中项目的信息
        item = self.projects_tree.item(selection[0])
        project_name = item['values'][0]
        project_id = self.projects_tree.item(selection[0], 'tags')[0]
        
        # 确认删除
        if messagebox.askyesno("确认删除", f"确定要删除项目 '{project_name}' 吗？\n此操作不可撤销！"):
            try:
                if self.project_manager.delete_project(project_id):
                    messagebox.showinfo("成功", f"项目 '{project_name}' 已删除！")
                    # 刷新项目列表
                    self.load_project_list()
                else:
                    messagebox.showerror("错误", "删除项目失败！项目可能不存在。")
            except Exception as e:
                messagebox.showerror("错误", f"删除项目时发生异常:\n{str(e)}")
                
    def copy_selected_project(self):
        """复制选中的项目"""
        selection = self.projects_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个项目！")
            return
        
        # 获取选中项目的信息
        item = self.projects_tree.item(selection[0])
        original_project_name = item['values'][0]
        original_project_id = self.projects_tree.item(selection[0], 'tags')[0]
        
        # 弹出对话框输入新项目名称
        dialog = tk.Toplevel(self.root)
        dialog.title("复制项目")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text=f"复制项目: {original_project_name}\n请输入新项目名称:").pack(pady=10)
        name_var = tk.StringVar(value=f"{original_project_name}_副本")
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        def confirm():
            new_name = name_var.get().strip()
            if not new_name:
                messagebox.showwarning("警告", "请输入新项目名称！")
                return
            
            # 检查新项目名称是否已存在
            projects = self.project_manager.get_project_list()
            if any(p['name'] == new_name for p in projects):
                messagebox.showwarning("警告", f"项目名称 '{new_name}' 已存在，请选择其他名称！")
                return
            
            # 复制项目
            try:
                # 加载原项目数据
                original_data = self.project_manager.load_project_data(original_project_id)
                if original_data is not None:
                    # 创建新项目
                    new_project = self.project_manager.create_project(new_name)
                    # 保存原数据到新项目
                    self.project_manager.save_project_data(new_project['id'], original_data)
                    
                    messagebox.showinfo("成功", f"项目 '{original_project_name}' 已复制为 '{new_name}'！")
                    # 刷新项目列表
                    self.load_project_list()
                else:
                    messagebox.showerror("错误", "原项目没有可复制的数据！")
            except Exception as e:
                messagebox.showerror("错误", f"复制项目时发生异常:\n{str(e)}")
            
            # 关闭对话框
            dialog.destroy()
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="确定", command=confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键
        name_entry.bind('<Return>', lambda e: confirm())
                
    def enter_main_app(self):
        """进入主应用程序界面"""
        # 清除项目管理界面
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # 创建主应用界面
        self.create_widgets()
        
        # 初始化导入数据趋势图
        self.initialize_data_plot()
        
        # 如果有当前项目，加载项目数据
        if self.current_project:
            project_data = self.project_manager.load_project_data(self.current_project['id'])
            if project_data is not None:  # 修改判断条件，允许空数据
                calculation_results = self.data_model.from_dict(project_data)
                if calculation_results:
                    self.results = calculation_results
                    # 显示结果
                    self.display_results()
                # 更新数据统计和趋势图
                self.update_statistics()
                self.update_imported_data_plot()
                
        # 刷新风机和光伏型号列表，确保默认选中第一个型号
        self.root.after(100, self.refresh_wind_model_list)
        self.root.after(100, self.refresh_pv_model_list)
        
        # 加载检修和投产计划数据
        self.root.after(100, self.load_maintenance_schedules)
                
        # 添加返回项目列表按钮
        back_btn = ttk.Button(self.root, text="返回项目列表", command=self.return_to_project_list)
        back_btn.place(relx=1.0, rely=0.0, anchor="ne", x=-10, y=10)
        
        # 检查并自动加载已存在的数据或计算结果
        self.auto_load_existing_data()
        
    def auto_load_existing_data(self):
        """
        自动加载已存在的数据或计算结果
        """
        if self.current_project:
            project_data = self.project_manager.load_project_data(self.current_project['id'])
            if project_data is not None:
                # 检查是否已存在导入的数据
                has_imported_data = any([
                    any(self.data_model.electric_load_hourly),
                    any(self.data_model.heat_load_hourly),
                    any(self.data_model.solar_irradiance_hourly),
                    any(self.data_model.wind_speed_hourly)
                ])
                
                if has_imported_data:
                    # 更新数据统计和趋势图
                    self.update_statistics()
                    self.update_imported_data_plot()
                    
                # 检查是否已存在计算结果
                if 'calculation_results' in project_data and project_data['calculation_results']:
                    self.results = project_data['calculation_results']
                    self.display_results()
                    self.update_plot()
        
    def return_to_project_list(self):
        """返回项目列表界面"""
        # 保存当前项目数据
        if self.current_project:
            self.save_current_project()
            
        # 清除当前项目
        self.current_project = None
        
        # 返回项目管理界面
        self.create_project_management_ui()
                
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Notebook for different sections
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 数据导入标签页
        self.create_data_import_tab(notebook)
        
        # 函数设置标签页
        self.create_function_settings_tab(notebook)
        
        # 检修和投产计划标签页
        self.create_maintenance_schedule_tab(notebook)
        
        # 计算与结果标签页
        self.create_calculation_tab(notebook)
                
        # 优化标签页
        self.create_optimization_tab(notebook)
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # 配置notebook的权重，确保标签页内容可以正确缩放
        notebook.columnconfigure(0, weight=1)
        notebook.rowconfigure(0, weight=1)
        
    def create_data_import_tab(self, notebook):
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="📊 数据导入")  # 添加数据图标
        
        # 添加返回项目列表按钮
        back_btn = ttk.Button(tab, text="保存并返回项目列表", command=self.save_and_return_to_project_list)
        back_btn.grid(row=0, column=3, sticky=tk.E, padx=5, pady=5)
        
        # 数据说明
        info_label = ttk.Label(tab, text="请导入包含8760小时数据的CSV文件\n"
                                         "文件应包含列: 时间, 电力负荷(kW), 热力负荷(kW), 光照强度(W/m²), 风速(m/s)")
        info_label.grid(row=1, column=0, columnspan=4, pady=(0, 20), sticky=tk.W)
        
        # 添加下载模板按钮
        ttk.Button(tab, text="下载CSV模板", command=self.download_template).grid(row=2, column=0, pady=5, sticky=tk.W)
        
    def initialize_data_plot(self):
        """
        初始化导入数据趋势图
        """
        # 检查是否已经创建了图形对象，如果没有则创建
        if not hasattr(self, 'data_ax'):
            self.data_figure = Figure(figsize=(10, 6), dpi=100)  # 增加高度
            self.data_ax = self.data_figure.add_subplot(111)
        
        self.data_ax.clear()
        self.data_ax.text(0.5, 0.5, '暂无数据\n请导入数据后查看趋势图', 
                          horizontalalignment='center', verticalalignment='center',
                          transform=self.data_ax.transAxes, fontsize=12)
        self.data_ax.set_title('已导入数据趋势图')
        
        # 如果画布不存在，则创建
        if not hasattr(self, 'data_canvas'):
            self.data_canvas = FigureCanvasTkAgg(self.data_figure, self.root)  # 临时挂载到root
            
        self.data_canvas.draw()
        
        # 启用matplotlib交互功能
        self.data_figure.tight_layout()
        # 为图例预留额外空间
        self.data_figure.subplots_adjust(right=0.85)
        
    def update_imported_data_plot(self):
        """
        更新已导入数据的趋势图
        """
        # 清除之前的图表
        self.data_ax.clear()
        
        # 解析时间段
        try:
            start_date = datetime.strptime(self.start_date_var.get(), "%Y-%m-%d")
            end_date = datetime.strptime(self.end_date_var.get(), "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("错误", "日期格式不正确，请使用 YYYY-MM-DD 格式")
            return
        
        # 计算时间段对应的小时索引
        start_hour = self.date_to_hour(start_date)
        end_hour = self.date_to_hour(end_date)
        
        if start_hour > end_hour:
            messagebox.showerror("错误", "开始日期不能晚于结束日期")
            return
        
        if start_hour < 0 or end_hour >= 8760:
            messagebox.showerror("错误", "日期超出范围，应在2025-01-01至2025-12-31之间")
            return
        
        # 获取时间段内的数据
        hours = list(range(start_hour, end_hour + 1))
        
        # 将小时转换为日期格式
        dates = [datetime(2025, 1, 1) + timedelta(hours=h) for h in hours]
        date_labels = [d.strftime('%m-%d') for d in dates]
        
        # 绘制已导入的数据
        lines = []  # 存储所有绘制的线条
        labels = []  # 存储所有标签
        
        if self.data_model.data_imported['electric']:
            electric_data = [self.data_model.electric_load_hourly[i] for i in hours]
            line, = self.data_ax.plot(dates, electric_data, label='电力负荷(kW)', linewidth=0.5)
            lines.append(line)
            labels.append('电力负荷(kW)')
            
        if self.data_model.data_imported['heat']:
            heat_data = [self.data_model.heat_load_hourly[i] for i in hours]
            line, = self.data_ax.plot(dates, heat_data, label='热力负荷(kW)', linewidth=0.5)
            lines.append(line)
            labels.append('热力负荷(kW)')
            
        if self.data_model.data_imported['solar']:
            solar_data = [self.data_model.solar_irradiance_hourly[i] for i in hours]
            line, = self.data_ax.plot(dates, solar_data, label='光照强度(W/m²)', linewidth=0.5)
            lines.append(line)
            labels.append('光照强度(W/m²)')
            
        if self.data_model.data_imported['wind']:
            wind_data = [self.data_model.wind_speed_hourly[i] for i in hours]
            line, = self.data_ax.plot(dates, wind_data, label='风速(m/s)', linewidth=0.5)
            lines.append(line)
            labels.append('风速(m/s)')
        
        if self.data_model.data_imported['grid_price']:
            grid_price_data = [self.data_model.grid_purchase_price_hourly[i] for i in hours]
            line, = self.data_ax.plot(dates, grid_price_data, label='下网电价(元/kWh)', linewidth=0.5)
            lines.append(line)
            labels.append('下网电价(元/kWh)')
        
        self.data_ax.set_xlabel('日期 (MM-DD)')
        self.data_ax.set_ylabel('数值')
        self.data_ax.set_title(f'已导入数据趋势图 ({self.start_date_var.get()} 至 {self.end_date_var.get()})')
        
        # 设置x轴日期格式
        self.data_ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%m-%d'))
        
        # 根据时间跨度自动选择适当的日期定位器
        date_span = (dates[-1] - dates[0]).days
        if date_span <= 31:  # 一个月内，使用周定位器
            self.data_ax.xaxis.set_major_locator(plt.matplotlib.dates.WeekdayLocator(interval=1))
        elif date_span <= 180:  # 6个月内，使用双周定位器
            self.data_ax.xaxis.set_major_locator(plt.matplotlib.dates.WeekdayLocator(interval=2))
        else:  # 超过6个月，使用月定位器
            self.data_ax.xaxis.set_major_locator(plt.matplotlib.dates.MonthLocator())
        
        # 旋转x轴标签以更好地显示
        plt.setp(self.data_ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
        
        # 创建图例并启用点击功能
        legend = self.data_ax.legend(loc='upper left', bbox_to_anchor=(1, 1))  # 固定图例位置
        
        # 启用图例点击交互
        lined = {}
        for legline, origline in zip(legend.get_lines(), lines):
            legline.set_picker(True)  # Enable picking on the legend line
            lined[legline] = origline
        
        # 为图例中的文本也启用点击
        for legtext, origline in zip(legend.get_texts(), lines):
            legtext.set_picker(True)  # Enable picking on the legend text
            lined[legtext] = origline
        
        # 保存线条和图例映射关系
        self.lined_data = lined
        
        # 连接点击事件
        self.data_canvas.mpl_connect('pick_event', self.on_legend_click_data)
        
        # 连接鼠标移动事件以实现悬浮功能
        self.data_canvas.mpl_connect('motion_notify_event', self.on_data_hover)
        
        self.data_ax.grid(True, alpha=0.3)
        
        # 设置x轴范围
        self.data_ax.set_xlim(dates[0], dates[-1])
        
        # 调整子图参数以确保图例和标签完全显示
        self.data_figure.tight_layout()
        # 为图例预留额外空间
        self.data_figure.subplots_adjust(right=0.85)
        
        # 刷新画布
        self.data_canvas.draw()
        
    def create_function_settings_tab(self, notebook):
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="⚙️ 机组设置")  # 修改标签名称为"设置"并添加齿轮图标
        
    def create_data_import_tab(self, notebook):
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="📊 数据导入")  # 添加数据图标
        
        # 添加返回项目列表按钮
        back_btn = ttk.Button(tab, text="保存并返回项目列表", command=self.save_and_return_to_project_list)
        back_btn.grid(row=0, column=3, sticky=tk.E, padx=5, pady=5)
        
        # 数据说明
        info_label = ttk.Label(tab, text="请导入包含8760小时数据的CSV文件\n"
                                         "文件应包含列: 时间, 电力负荷(kW), 热力负荷(kW), 光照强度(W/m²), 风速(m/s), 下网电价(元/kWh)")
        info_label.grid(row=1, column=0, columnspan=4, pady=(0, 20), sticky=tk.W)
        
        # 添加下载模板按钮
        ttk.Button(tab, text="下载CSV模板", command=self.download_template).grid(row=2, column=0, pady=5, sticky=tk.W)
        
        # 单一文件导入控件
        ttk.Label(tab, text="统一数据文件:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.single_file_path = tk.StringVar()
        # 初始化单文件模式变量
        self.single_file_mode = tk.BooleanVar(value=True)  # 默认使用单文件模式
        self.single_file_entry = ttk.Entry(tab, textvariable=self.single_file_path, width=50)
        self.single_file_entry.grid(row=3, column=1, padx=5, pady=5)
        self.single_file_button = ttk.Button(
            tab, 
            text="浏览...", 
            command=lambda: self.browse_file(self.single_file_path)
        )
        self.single_file_button.grid(row=3, column=2, pady=5)
        
        # 厂用电率设置
        ttk.Label(tab, text="厂用电率:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.internal_rate_var = tk.DoubleVar(value=0.05)
        ttk.Entry(tab, textvariable=self.internal_rate_var, width=20).grid(row=4, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(tab, text="(小数形式, 如0.05表示5%)").grid(row=4, column=1, sticky=tk.E, padx=5, pady=5)
        
        # 导入按钮
        ttk.Button(tab, text="导入数据", command=self.import_all_data).grid(row=5, column=0, columnspan=3, pady=20)
        
        # 刷新图表按钮
        ttk.Button(tab, text="刷新趋势图", command=self.update_imported_data_plot).grid(row=5, column=2, pady=20, sticky=tk.E)
        
        # 数据统计
        stats_frame = ttk.LabelFrame(tab, text="数据统计", padding="10")
        stats_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.stats_text = tk.Text(stats_frame, height=6, width=80)  # 增加高度
        scrollbar = ttk.Scrollbar(stats_frame, orient=tk.VERTICAL, command=self.stats_text.yview)
        self.stats_text.configure(yscrollcommand=scrollbar.set)
        
        self.stats_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 时间段选择区域
        time_range_frame = ttk.LabelFrame(tab, text="时间段选择", padding="10")
        time_range_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(time_range_frame, text="开始日期:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.start_date_var = tk.StringVar(value="2025-01-01")
        self.start_date_entry = ttk.Entry(time_range_frame, textvariable=self.start_date_var, width=12)
        self.start_date_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(time_range_frame, text="结束日期:").grid(row=0, column=2, sticky=tk.W, padx=(10, 5))
        self.end_date_var = tk.StringVar(value="2025-12-31")
        self.end_date_entry = ttk.Entry(time_range_frame, textvariable=self.end_date_var, width=12)
        self.end_date_entry.grid(row=0, column=3, padx=5)
        
        ttk.Button(time_range_frame, text="更新图表", command=self.update_imported_data_plot).grid(row=0, column=4, padx=(10, 0))
        
        # 图表展示（用于显示导入数据的趋势）
        plot_frame = ttk.LabelFrame(tab, text="已导入数据趋势图", padding="10")
        plot_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # 创建matplotlib图形
        self.data_figure = Figure(figsize=(10, 6), dpi=100)  # 增加高度
        self.data_ax = self.data_figure.add_subplot(111)
        self.data_canvas = FigureCanvasTkAgg(self.data_figure, plot_frame)
        self.data_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 配置权重
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(8, weight=1)  # 给图表区域分配更多空间
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)
        time_range_frame.columnconfigure(5, weight=1)
        plot_frame.columnconfigure(0, weight=1)
        plot_frame.rowconfigure(0, weight=1)

    def create_function_data_tab(self, notebook):
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="📊 数据导入")  # 修改为正确的标签名称
        
        # 统计信息区域
        stats_frame = ttk.LabelFrame(tab, text="统计信息", padding="10")
        stats_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(stats_frame, text="总电力消耗 (kWh):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.total_consumption_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.total_consumption_var, width=20).grid(row=0, column=1, pady=2)
        
        ttk.Label(stats_frame, text="平均电力负荷 (kW):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.avg_load_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.avg_load_var, width=20).grid(row=1, column=1, pady=2)
        
        ttk.Label(stats_frame, text="最大电力负荷 (kW):").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.max_load_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.max_load_var, width=20).grid(row=2, column=1, pady=2)
        
        ttk.Label(stats_frame, text="最小电力负荷 (kW):").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.min_load_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.min_load_var, width=20).grid(row=3, column=1, pady=2)
        
        ttk.Label(stats_frame, text="运行时间 (小时):").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.runtime_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.runtime_var, width=20).grid(row=4, column=1, pady=2)
        
        ttk.Label(stats_frame, text="停机时间 (小时):").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.downtime_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.downtime_var, width=20).grid(row=5, column=1, pady=2)
        
        ttk.Label(stats_frame, text="总成本 (元):").grid(row=6, column=0, sticky=tk.W, pady=2)
        self.total_cost_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.total_cost_var, width=20).grid(row=6, column=1, pady=2)
        
        # 添加时间段选择区域（与数据导入tab保持一致）
        time_range_frame = ttk.LabelFrame(tab, text="时间段选择", padding="10")
        time_range_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(time_range_frame, text="开始日期:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.start_date_var = tk.StringVar(value="2025-01-01")
        self.start_date_entry = ttk.Entry(time_range_frame, textvariable=self.start_date_var, width=12)
        self.start_date_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(time_range_frame, text="结束日期:").grid(row=0, column=2, sticky=tk.W, padx=(10, 5))
        self.end_date_var = tk.StringVar(value="2025-12-31")
        self.end_date_entry = ttk.Entry(time_range_frame, textvariable=self.end_date_var, width=12)
        self.end_date_entry.grid(row=0, column=3, padx=5)
        
        ttk.Button(time_range_frame, text="更新图表", command=self.update_imported_data_plot).grid(row=0, column=4, padx=(10, 0))
        
        # 图表展示（用于显示导入数据的趋势）
        plot_frame = ttk.LabelFrame(tab, text="已导入数据趋势图", padding="10")
        plot_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # 创建matplotlib图形
        self.data_figure = Figure(figsize=(10, 6), dpi=100)  # 增加高度
        self.data_ax = self.data_figure.add_subplot(111)
        self.data_canvas = FigureCanvasTkAgg(self.data_figure, plot_frame)
        self.data_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 配置权重
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)  # 给图表区域分配更多空间
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)
        time_range_frame.columnconfigure(5, weight=1)
        plot_frame.columnconfigure(0, weight=1)
        plot_frame.rowconfigure(0, weight=1)
        
    def create_function_settings_tab(self, notebook):
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="⚙️ 机组设置")  # 修改标签名称为"设置"并添加齿轮图标
        
        # 按钮区域（放在同一行）
        buttons_frame = ttk.Frame(tab)
        buttons_frame.grid(row=0, column=0, columnspan=2, sticky=tk.E, padx=5, pady=5)
        
        # 保存函数参数按钮
        save_params_btn = ttk.Button(buttons_frame, text="保存函数参数", command=self.save_function_parameters)
        save_params_btn.pack(side=tk.RIGHT, padx=5)
        
        # 保存并返回项目列表按钮
        back_btn = ttk.Button(buttons_frame, text="保存并返回项目列表", command=self.save_and_return_to_project_list)
        back_btn.pack(side=tk.RIGHT, padx=5)
        
        # 负荷设置区域
        load_frame = ttk.LabelFrame(tab, text="负荷设置", padding="10")
        load_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(load_frame, text="最大电力负荷 (kW):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.max_load_var = tk.DoubleVar(value=self.data_model.max_electric_load)  # 使用数据模型中的值
        ttk.Entry(load_frame, textvariable=self.max_load_var, width=20).grid(row=0, column=1, pady=2)
        
        ttk.Label(load_frame, text="最大灵活负荷 (kW):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.flexible_load_max_var = tk.DoubleVar(value=self.data_model.flexible_load_max)  # 使用数据模型中的值
        ttk.Entry(load_frame, textvariable=self.flexible_load_max_var, width=20).grid(row=1, column=1, pady=2)
        
        ttk.Label(load_frame, text="最小灵活负荷 (kW):").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.flexible_load_min_var = tk.DoubleVar(value=self.data_model.flexible_load_min)  # 使用数据模型中的值
        ttk.Entry(load_frame, textvariable=self.flexible_load_min_var, width=20).grid(row=2, column=1, pady=2)
        
        # 风机和光伏型号管理区域（放在同一行）
        models_frame = ttk.Frame(tab)
        models_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        
    def create_function_data_tab(self, notebook):
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="📊 数据导入")  # 修改为正确的标签名称
        
        # 统计信息区域
        stats_frame = ttk.LabelFrame(tab, text="统计信息", padding="10")
        stats_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(stats_frame, text="总电力消耗 (kWh):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.total_consumption_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.total_consumption_var, width=20).grid(row=0, column=1, pady=2)
        
        ttk.Label(stats_frame, text="平均电力负荷 (kW):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.avg_load_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.avg_load_var, width=20).grid(row=1, column=1, pady=2)
        
        ttk.Label(stats_frame, text="最大电力负荷 (kW):").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.max_load_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.max_load_var, width=20).grid(row=2, column=1, pady=2)
        
        ttk.Label(stats_frame, text="最小电力负荷 (kW):").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.min_load_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.min_load_var, width=20).grid(row=3, column=1, pady=2)
        
        ttk.Label(stats_frame, text="运行时间 (小时):").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.runtime_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.runtime_var, width=20).grid(row=4, column=1, pady=2)
        
        ttk.Label(stats_frame, text="停机时间 (小时):").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.downtime_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.downtime_var, width=20).grid(row=5, column=1, pady=2)
        
        ttk.Label(stats_frame, text="总成本 (元):").grid(row=6, column=0, sticky=tk.W, pady=2)
        self.total_cost_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.total_cost_var, width=20).grid(row=6, column=1, pady=2)
        
        ttk.Label(stats_frame, text="平均成本 (元/kWh):").grid(row=7, column=0, sticky=tk.W, pady=2)
        self.avg_cost_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.avg_cost_var, width=20).grid(row=7, column=1, pady=2)
        
        # 时间范围选择区域
        time_range_frame = ttk.LabelFrame(tab, text="时间范围", padding="10")
        time_range_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(time_range_frame, text="开始时间:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.start_time_var = tk.StringVar(value="2023-01-01 00:00:00")
        ttk.Entry(time_range_frame, textvariable=self.start_time_var, width=20).grid(row=0, column=1, pady=2)
        
        ttk.Label(time_range_frame, text="结束时间:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.end_time_var = tk.StringVar(value="2023-12-31 23:59:59")
        ttk.Entry(time_range_frame, textvariable=self.end_time_var, width=20).grid(row=1, column=1, pady=2)
        
        # 图表区域
        plot_frame = ttk.Frame(tab)
        plot_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.data_figure = Figure(figsize=(10, 6), dpi=100)  # 增加高度
        self.data_ax = self.data_figure.add_subplot(111)
        self.data_canvas = FigureCanvasTkAgg(self.data_figure, plot_frame)
        self.data_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 配置权重
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(8, weight=1)  # 给图表区域分配更多空间
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)
        time_range_frame.columnconfigure(5, weight=1)
        plot_frame.columnconfigure(0, weight=1)
        plot_frame.rowconfigure(0, weight=1)
        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        

        
    def create_function_settings_tab(self, notebook):
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="⚙️ 机组设置")  # 修改标签名称为"设置"并添加齿轮图标
        
        # 按钮区域（放在同一行）
        buttons_frame = ttk.Frame(tab)
        buttons_frame.grid(row=0, column=0, columnspan=2, sticky=tk.E, padx=5, pady=5)
        
        # 保存函数参数按钮
        save_params_btn = ttk.Button(buttons_frame, text="保存函数参数", command=self.save_function_parameters)
        save_params_btn.pack(side=tk.RIGHT, padx=5)
        
        # 保存并返回项目列表按钮
        back_btn = ttk.Button(buttons_frame, text="保存并返回项目列表", command=self.save_and_return_to_project_list)
        back_btn.pack(side=tk.RIGHT, padx=5)
        
    def create_function_data_tab(self, notebook):
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="📊 数据导入")  # 修改为正确的标签名称
        
        # 统计信息区域
        stats_frame = ttk.LabelFrame(tab, text="统计信息", padding="10")
        stats_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(stats_frame, text="总电力消耗 (kWh):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.total_consumption_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.total_consumption_var, width=20).grid(row=0, column=1, pady=2)
        
        ttk.Label(stats_frame, text="平均电力负荷 (kW):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.avg_load_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.avg_load_var, width=20).grid(row=1, column=1, pady=2)
        
        ttk.Label(stats_frame, text="最大电力负荷 (kW):").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.max_load_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.max_load_var, width=20).grid(row=2, column=1, pady=2)
        
        ttk.Label(stats_frame, text="最小电力负荷 (kW):").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.min_load_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.min_load_var, width=20).grid(row=3, column=1, pady=2)
        
        ttk.Label(stats_frame, text="运行时间 (小时):").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.runtime_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.runtime_var, width=20).grid(row=4, column=1, pady=2)
        
        ttk.Label(stats_frame, text="停机时间 (小时):").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.downtime_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.downtime_var, width=20).grid(row=5, column=1, pady=2)
        
        ttk.Label(stats_frame, text="总成本 (元):").grid(row=6, column=0, sticky=tk.W, pady=2)
        self.total_cost_var = tk.StringVar(value="0.0")
        ttk.Entry(stats_frame, textvariable=self.total_cost_var, width=20).grid(row=6, column=1, pady=2)
        
        stats_frame.rowconfigure(0, weight=1)
        
    def create_function_settings_tab(self, notebook):
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="⚙️ 机组设置")  # 修改标签名称为"设置"并添加齿轮图标
        
        # 按钮区域（放在同一行）
        buttons_frame = ttk.Frame(tab)
        buttons_frame.grid(row=0, column=0, columnspan=2, sticky=tk.E, padx=5, pady=5)
        
        # 保存函数参数按钮
        save_params_btn = ttk.Button(buttons_frame, text="保存函数参数", command=self.save_function_parameters)
        save_params_btn.pack(side=tk.RIGHT, padx=5)
        
        # 保存并返回项目列表按钮
        back_btn = ttk.Button(buttons_frame, text="保存并返回项目列表", command=self.save_and_return_to_project_list)
        back_btn.pack(side=tk.RIGHT, padx=5)
        
        # 负荷设置区域
        load_frame = ttk.LabelFrame(tab, text="负荷设置", padding="10")
        load_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(load_frame, text="最大电力负荷 (kW):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.max_load_var = tk.DoubleVar(value=self.data_model.max_electric_load)  # 使用数据模型中的值
        ttk.Entry(load_frame, textvariable=self.max_load_var, width=20).grid(row=0, column=1, pady=2)
        
        ttk.Label(load_frame, text="最大灵活负荷 (kW):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.flexible_load_max_var = tk.DoubleVar(value=self.data_model.flexible_load_max)  # 使用数据模型中的值
        ttk.Entry(load_frame, textvariable=self.flexible_load_max_var, width=20).grid(row=1, column=1, pady=2)
        
        ttk.Label(load_frame, text="最小灵活负荷 (kW):").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.flexible_load_min_var = tk.DoubleVar(value=self.data_model.flexible_load_min)  # 使用数据模型中的值
        ttk.Entry(load_frame, textvariable=self.flexible_load_min_var, width=20).grid(row=2, column=1, pady=2)
        
        # 风机和光伏型号管理区域（放在同一行）
        models_frame = ttk.Frame(tab)
        models_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        
        # 风机型号管理
        wind_models_frame = ttk.LabelFrame(models_frame, text="风机型号管理", padding="10")
        wind_models_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # 风机总装机容量显示
        self.wind_total_capacity_label = ttk.Label(wind_models_frame, text="总装机容量: 0 kW")
        self.wind_total_capacity_label.grid(row=5, column=0, columnspan=2, pady=5, sticky=tk.W)
        
        # 风机型号列表
        self.wind_model_listbox = tk.Listbox(wind_models_frame, height=8)
        self.wind_model_listbox.grid(row=0, column=0, rowspan=5, padx=(0, 10))
        self.wind_model_listbox.bind('<<ListboxSelect>>', self.on_wind_model_select)
        self.wind_model_listbox.bind('<Double-Button-1>', self.edit_wind_model)
        
        # 风机型号操作按钮
        ttk.Button(wind_models_frame, text="添加型号", command=self.add_wind_model).grid(row=0, column=1, pady=2, sticky=tk.W+tk.E)
        ttk.Button(wind_models_frame, text="删除型号", command=self.delete_wind_model).grid(row=1, column=1, pady=2, sticky=tk.W+tk.E)
        ttk.Button(wind_models_frame, text="编辑型号", command=self.edit_wind_model).grid(row=2, column=1, pady=2, sticky=tk.W+tk.E)
        
        # 风机型号详情编辑区域
        self.wind_model_detail_frame = ttk.LabelFrame(wind_models_frame, text="型号详情", padding="10")
        self.wind_model_detail_frame.grid(row=0, column=2, rowspan=5, padx=(10, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 当前编辑的型号索引
        self.current_editing_index = None
        
        ttk.Label(self.wind_model_detail_frame, text="型号名称:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.wind_model_name = tk.StringVar()
        ttk.Entry(self.wind_model_detail_frame, textvariable=self.wind_model_name, width=20).grid(row=0, column=1, pady=2)
        
        # 将切入风速和额定风速放在同一行
        wind_speed_frame = ttk.Frame(self.wind_model_detail_frame)
        wind_speed_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=2)
        
        ttk.Label(wind_speed_frame, text="切入风速:").grid(row=0, column=0, sticky=tk.W)
        self.wind_model_cut_in = tk.DoubleVar()
        ttk.Entry(wind_speed_frame, textvariable=self.wind_model_cut_in, width=10).grid(row=0, column=1, padx=(0, 5))
        
        ttk.Label(wind_speed_frame, text="额定风速:").grid(row=0, column=2, sticky=tk.W, padx=(10, 0))
        self.wind_model_rated = tk.DoubleVar()
        ttk.Entry(wind_speed_frame, textvariable=self.wind_model_rated, width=10).grid(row=0, column=3)
        
        # 将最大额定风速和切出风速放在同一行
        wind_speed_frame2 = ttk.Frame(self.wind_model_detail_frame)
        wind_speed_frame2.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=2)
        
        ttk.Label(wind_speed_frame2, text="最大额定风速:").grid(row=0, column=0, sticky=tk.W)
        self.wind_model_max_rated = tk.DoubleVar()
        ttk.Entry(wind_speed_frame2, textvariable=self.wind_model_max_rated, width=10).grid(row=0, column=1, padx=(0, 5))
        
        ttk.Label(wind_speed_frame2, text="切出风速:").grid(row=0, column=2, sticky=tk.W, padx=(10, 0))
        self.wind_model_cut_out = tk.DoubleVar()
        ttk.Entry(wind_speed_frame2, textvariable=self.wind_model_cut_out, width=10).grid(row=0, column=3)
        
        ttk.Label(self.wind_model_detail_frame, text="额定功率 (kW):").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.wind_model_rated_power = tk.DoubleVar()
        ttk.Entry(self.wind_model_detail_frame, textvariable=self.wind_model_rated_power, width=20).grid(row=3, column=1, pady=2)
        
        ttk.Label(self.wind_model_detail_frame, text="风机数量:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.wind_model_count = tk.IntVar()
        ttk.Entry(self.wind_model_detail_frame, textvariable=self.wind_model_count, width=20).grid(row=4, column=1, pady=2)
        
        # 添加出力修正系数输入框
        ttk.Label(self.wind_model_detail_frame, text="出力修正系数:").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.wind_model_correction_factor = tk.DoubleVar(value=1.0)
        ttk.Entry(self.wind_model_detail_frame, textvariable=self.wind_model_correction_factor, width=20).grid(row=5, column=1, pady=2)
        
        ttk.Button(self.wind_model_detail_frame, text="保存型号", command=self.save_wind_model).grid(row=6, column=0, columnspan=2, pady=5, sticky=tk.W+tk.E)
        
        # 风机函数图像显示
        self.wind_function_frame = ttk.LabelFrame(wind_models_frame, text="风机函数图像", padding="10")
        self.wind_function_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建matplotlib图形用于显示风机函数关系
        self.wind_figure, self.wind_ax = plt.subplots(figsize=(6, 3))  # 调整大小适应界面
        self.wind_function_canvas = FigureCanvasTkAgg(self.wind_figure, self.wind_function_frame)
        self.wind_function_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 光伏型号管理
        pv_models_frame = ttk.LabelFrame(models_frame, text="光伏型号管理", padding="10")
        pv_models_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        
        # 光伏总装机容量显示
        self.pv_total_capacity_label = ttk.Label(pv_models_frame, text="总装机容量: 0 kW")
        self.pv_total_capacity_label.grid(row=5, column=0, columnspan=2, pady=5, sticky=tk.W)
        
        # 光伏型号列表
        self.pv_model_listbox = tk.Listbox(pv_models_frame, height=8)
        self.pv_model_listbox.grid(row=0, column=0, rowspan=5, padx=(0, 10))
        self.pv_model_listbox.bind('<<ListboxSelect>>', self.on_pv_model_select)
        self.pv_model_listbox.bind('<Double-Button-1>', self.edit_pv_model)
        
        # 光伏型号操作按钮
        ttk.Button(pv_models_frame, text="添加型号", command=self.add_pv_model).grid(row=0, column=1, pady=2, sticky=tk.W+tk.E)
        ttk.Button(pv_models_frame, text="删除型号", command=self.delete_pv_model).grid(row=1, column=1, pady=2, sticky=tk.W+tk.E)
        ttk.Button(pv_models_frame, text="编辑型号", command=self.edit_pv_model).grid(row=2, column=1, pady=2, sticky=tk.W+tk.E)
        
        # 光伏型号详情编辑区域
        self.pv_model_detail_frame = ttk.LabelFrame(pv_models_frame, text="型号详情", padding="10")
        self.pv_model_detail_frame.grid(row=0, column=2, rowspan=5, padx=(10, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 当前编辑的型号索引
        self.current_pv_editing_index = None
        
        ttk.Label(self.pv_model_detail_frame, text="型号名称:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.pv_model_name = tk.StringVar()
        ttk.Entry(self.pv_model_detail_frame, textvariable=self.pv_model_name, width=20).grid(row=0, column=1, pady=2)
        
        ttk.Label(self.pv_model_detail_frame, text="计算方法:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.pv_model_method = tk.StringVar(value="area_efficiency")
        method_frame = ttk.Frame(self.pv_model_detail_frame)
        method_frame.grid(row=1, column=1, pady=2, sticky=tk.W)
        
        self.area_efficiency_radio = ttk.Radiobutton(method_frame, text="面积效率法", variable=self.pv_model_method, 
                                                    value="area_efficiency", command=self.on_pv_method_change)
        self.area_efficiency_radio.pack(side=tk.LEFT)
        
        self.installed_capacity_radio = ttk.Radiobutton(method_frame, text="装机容量法", variable=self.pv_model_method, 
                                                      value="installed_capacity", command=self.on_pv_method_change)
        self.installed_capacity_radio.pack(side=tk.LEFT)
        
        # 面积效率法参数
        self.area_efficiency_frame = ttk.Frame(self.pv_model_detail_frame)
        self.area_efficiency_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(self.area_efficiency_frame, text="光伏板效率:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.pv_panel_efficiency = tk.DoubleVar()
        ttk.Entry(self.area_efficiency_frame, textvariable=self.pv_panel_efficiency, width=20).grid(row=0, column=1, pady=2)
        
        ttk.Label(self.area_efficiency_frame, text="光伏板面积 (m²):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.pv_panel_area = tk.DoubleVar()
        ttk.Entry(self.area_efficiency_frame, textvariable=self.pv_panel_area, width=20).grid(row=1, column=1, pady=2)
        
        # 装机容量法参数
        self.installed_capacity_frame = ttk.Frame(self.pv_model_detail_frame)
        self.installed_capacity_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(self.installed_capacity_frame, text="装机容量 (kW):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.pv_installed_capacity = tk.DoubleVar()
        ttk.Entry(self.installed_capacity_frame, textvariable=self.pv_installed_capacity, width=20).grid(row=0, column=1, pady=2)
        
        ttk.Label(self.installed_capacity_frame, text="系统效率:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.pv_system_efficiency = tk.DoubleVar(value=0.9)
        ttk.Entry(self.installed_capacity_frame, textvariable=self.pv_system_efficiency, width=20).grid(row=1, column=1, pady=2)
        
        # 公共参数
        ttk.Label(self.pv_model_detail_frame, text="光伏板数量:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.pv_model_count = tk.IntVar()
        ttk.Entry(self.pv_model_detail_frame, textvariable=self.pv_model_count, width=20).grid(row=4, column=1, pady=2)
        
        # 添加出力修正系数输入框
        ttk.Label(self.pv_model_detail_frame, text="出力修正系数:").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.pv_model_correction_factor = tk.DoubleVar(value=1.0)
        ttk.Entry(self.pv_model_detail_frame, textvariable=self.pv_model_correction_factor, width=20).grid(row=5, column=1, pady=2)
        
        ttk.Button(self.pv_model_detail_frame, text="保存型号", command=self.save_pv_model).grid(row=6, column=0, columnspan=2, pady=5, sticky=tk.W+tk.E)
        
        # 光伏函数图像显示
        self.pv_function_frame = ttk.LabelFrame(pv_models_frame, text="光伏函数图像", padding="10")
        self.pv_function_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建matplotlib图形用于显示光伏函数关系
        self.pv_figure, self.pv_ax = plt.subplots(figsize=(6, 3))  # 调整大小适应界面
        self.pv_function_canvas = FigureCanvasTkAgg(self.pv_figure, self.pv_function_frame)
        self.pv_function_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 创建一个主框架来容纳热电联产和调峰机组设置
        settings_main_frame = ttk.Frame(tab)
        settings_main_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 0))
        
        # 热电联产设置区域
        chp_frame = ttk.LabelFrame(settings_main_frame, text="热电联产设置", padding="10")
        chp_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        ttk.Label(chp_frame, text="电热比:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.electric_heat_ratio = tk.DoubleVar(value=self.data_model.chp_electric_params['electric_heat_ratio'])
        ttk.Entry(chp_frame, textvariable=self.electric_heat_ratio, width=20).grid(row=0, column=1, pady=2)
        
        ttk.Label(chp_frame, text="基础发电量 (kW):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.base_electric = tk.DoubleVar(value=self.data_model.chp_electric_params['base_electric'])
        ttk.Entry(chp_frame, textvariable=self.base_electric, width=20).grid(row=1, column=1, pady=2)
        
        # 调峰机组设置区域
        peak_frame = ttk.LabelFrame(settings_main_frame, text="调峰机组设置", padding="10")
        peak_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        ttk.Label(peak_frame, text="调峰机组最大出力 (kW):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.peak_power_max = tk.DoubleVar(value=self.data_model.peak_power_max)
        ttk.Entry(peak_frame, textvariable=self.peak_power_max, width=20).grid(row=0, column=1, pady=2)
        
        ttk.Label(peak_frame, text="调峰机组夏季最小出力 (kW):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.peak_power_min_summer = tk.DoubleVar(value=self.data_model.peak_power_min_summer)
        ttk.Entry(peak_frame, textvariable=self.peak_power_min_summer, width=20).grid(row=1, column=1, pady=2)
        
        ttk.Label(peak_frame, text="调峰机组冬季最小出力 (kW):").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.peak_power_min_winter = tk.DoubleVar(value=self.data_model.peak_power_min_winter)
        ttk.Entry(peak_frame, textvariable=self.peak_power_min_winter, width=20).grid(row=2, column=1, pady=2)
        
        # 配置权重，使两个设置区域平分空间
        settings_main_frame.columnconfigure(0, weight=1)
        settings_main_frame.columnconfigure(1, weight=1)
        settings_main_frame.rowconfigure(0, weight=1)
        
        # 配置权重
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(2, weight=1)  # 型号管理区域可以扩展
        models_frame.columnconfigure(0, weight=1)
        models_frame.columnconfigure(1, weight=1)
        models_frame.rowconfigure(0, weight=1)
        
        # 初始化风机和光伏型号列表
        self.root.after(100, self.refresh_wind_model_list)
        self.root.after(100, self.refresh_pv_model_list)
        

        
            



        
            



        
            



        
            



        
            



        
            



        
            

            

            

    def delete_wind_model(self):
        """
        删除选中的风机型号
        """
        selection = self.wind_model_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个风机型号！")
            return
        
        index = selection[0]
        model_name = self.data_model.wind_turbine_models[index]['name']
        
        # 确认删除
        if messagebox.askyesno("确认删除", f"确定要删除风机型号 '{model_name}' 吗？"):
            # 从数据模型中删除
            del self.data_model.wind_turbine_models[index]
            # 刷新列表
            self.refresh_wind_model_list()
            # 清空详情区域
            self.clear_wind_model_details()
            # 更新风机总装机容量显示
            self.update_wind_total_capacity()

    def delete_pv_model(self):
        """
        删除选中的光伏型号
        """
        selection = self.pv_model_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个光伏型号！")
            return
        
        index = selection[0]
        model_name = self.data_model.pv_models[index]['name']
        
        # 确认删除
        if messagebox.askyesno("确认删除", f"确定要删除光伏型号 '{model_name}' 吗？"):
            # 从数据模型中删除
            del self.data_model.pv_models[index]
            # 刷新列表
            self.refresh_pv_model_list()
            # 清空详情区域
            self.clear_pv_model_details()
            # 更新光伏总装机容量显示
            self.update_pv_total_capacity()
            
    def clear_wind_model_details(self):
        """
        清空风机型号详情区域
        """
        self.wind_model_name.set("")
        self.wind_model_cut_in.set(0.0)
        self.wind_model_rated.set(0.0)
        self.wind_model_max_rated.set(0.0)
        self.wind_model_cut_out.set(0.0)
        self.wind_model_rated_power.set(0.0)
        self.wind_model_count.set(0)
        
        # 清空风机函数图像
        self.wind_ax.clear()
        self.wind_ax.text(0.5, 0.5, '请选择或添加风机型号', 
                         horizontalalignment='center', verticalalignment='center',
                         transform=self.wind_ax.transAxes, fontsize=12)
        self.wind_ax.set_title('风机出力函数')
        self.wind_function_canvas.draw()
        
    def clear_pv_model_details(self):
        """
        清空光伏型号详情区域
        """
        self.pv_model_name.set("")
        self.pv_model_method.set("area_efficiency")
        self.pv_panel_efficiency.set(0.0)
        self.pv_panel_area.set(0.0)
        self.pv_installed_capacity.set(0.0)
        self.pv_system_efficiency.set(0.9)
        self.pv_model_count.set(0)
        
        # 清空光伏函数图像
        self.pv_ax.clear()
        self.pv_ax.text(0.5, 0.5, '请选择或添加光伏型号', 
                       horizontalalignment='center', verticalalignment='center',
                       transform=self.pv_ax.transAxes, fontsize=12)
        self.pv_ax.set_title('光伏出力函数')
        self.pv_function_canvas.draw()
        
    def update_wind_total_capacity(self):
        """
        更新风机总装机容量显示
        """
        total_capacity = self.data_model.calculate_wind_total_capacity()
        self.wind_total_capacity_label.config(text=f"总装机容量: {total_capacity:.2f} kW")
        
    def update_pv_total_capacity(self):
        """
        更新光伏总装机容量显示
        """
        total_capacity = self.data_model.calculate_pv_total_capacity()
        self.pv_total_capacity_label.config(text=f"总装机容量: {total_capacity:.2f} kW")
        
    def edit_wind_model(self, event=None):
        """
        编辑选中的风机型号
        """
        selection = self.wind_model_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个风机型号！")
            return
            
        # 设置当前编辑索引
        self.current_editing_index = selection[0]
        
        # 触发选择事件来填充编辑区域
        self.wind_model_listbox.event_generate("<<ListboxSelect>>")
        
    def edit_pv_model(self, event=None):
        """
        编辑选中的光伏型号
        """
        selection = self.pv_model_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个光伏型号！")
            return
            
        # 设置当前编辑索引
        self.current_pv_editing_index = selection[0]
        
        # 触发选择事件来填充编辑区域
        self.pv_model_listbox.event_generate("<<ListboxSelect>>")
        
    def on_wind_model_select(self, event):
        """
        当选择风机型号时的回调函数
        """
        selection = self.wind_model_listbox.curselection()
        if selection:
            index = selection[0]
            self.current_editing_index = index
            
            model = self.data_model.wind_turbine_models[index]
            
            # 填充详情字段
            self.wind_model_name.set(model['name'])
            self.wind_model_cut_in.set(model['params']['cut_in_wind'])
            self.wind_model_rated.set(model['params']['rated_wind'])
            self.wind_model_max_rated.set(model['params']['max_rated_wind'])
            self.wind_model_cut_out.set(model['params']['cut_out_wind'])
            self.wind_model_rated_power.set(model['params']['rated_power'])
            self.wind_model_count.set(model['count'])
            # 设置出力修正系数
            self.wind_model_correction_factor.set(model.get('output_correction_factor', 1.0))
            
            # 绘制当前选中风机型号的函数曲线
            self.plot_single_wind_curve(model)
    
    def on_pv_model_select(self, event):
        """
        当选择光伏型号时的回调函数
        """
        selection = self.pv_model_listbox.curselection()
        if selection:
            index = selection[0]
            self.current_pv_editing_index = index
            
            model = self.data_model.pv_models[index]
            
            # 填充详情字段
            self.pv_model_name.set(model['name'])
            self.pv_model_method.set(model['method'])
            self.pv_model_count.set(model['count'])
            # 设置出力修正系数
            self.pv_model_correction_factor.set(model.get('output_correction_factor', 1.0))
            
            # 根据计算方法填充参数
            params = model['params']
            if model['method'] == 'area_efficiency':
                self.pv_panel_efficiency.set(params.get('panel_efficiency', 0.2))
                self.pv_panel_area.set(params.get('panel_area', 1000.0))
            elif model['method'] == 'installed_capacity':
                self.pv_installed_capacity.set(params.get('installed_capacity', 200.0))
                self.pv_system_efficiency.set(params.get('system_efficiency', 0.9))
            
            # 更新界面显示
            self.on_pv_method_change()
            
            # 绘制当前选中光伏型号的函数曲线
            self.plot_single_pv_curve(model)

    def save_wind_model(self):
        """
        保存当前编辑的风机型号
        """
        # 获取输入的参数
        name = self.wind_model_name.get().strip()
        if not name:
            messagebox.showwarning("警告", "请输入型号名称！")
            return
            
        cut_in_wind = self.wind_model_cut_in.get()
        rated_wind = self.wind_model_rated.get()
        max_rated_wind = self.wind_model_max_rated.get()
        cut_out_wind = self.wind_model_cut_out.get()
        rated_power = self.wind_model_rated_power.get()
        count = self.wind_model_count.get()
        # 获取出力修正系数
        correction_factor = self.wind_model_correction_factor.get()
        
        # 参数校验
        if cut_in_wind >= rated_wind or rated_wind >= max_rated_wind or max_rated_wind >= cut_out_wind:
            messagebox.showerror("错误", "风速参数不符合逻辑关系！\n应满足：切入风速 < 额定风速 < 最大额定风速 < 切出风速")
            return
            
        if rated_power <= 0 or count <= 0:
            messagebox.showerror("错误", "额定功率和数量必须大于0！")
            return
            
        if correction_factor < 0:
            messagebox.showerror("错误", "出力修正系数不能为负数！")
            return
            
        # 创建型号数据
        model_data = {
            'name': name,
            'params': {
                'cut_in_wind': cut_in_wind,
                'rated_wind': rated_wind,
                'max_rated_wind': max_rated_wind,
                'cut_out_wind': cut_out_wind,
                'rated_power': rated_power
            },
            'count': count,
            'output_correction_factor': correction_factor  # 添加出力修正系数
        }
        
        # 如果是编辑现有型号
        if self.current_editing_index is not None and self.current_editing_index < len(self.data_model.wind_turbine_models):
            self.data_model.wind_turbine_models[self.current_editing_index] = model_data
        # 如果是新增型号（理论上不会走到这里，因为新增型号会直接添加）
        else:
            self.data_model.wind_turbine_models.append(model_data)
            
        # 刷新列表
        self.refresh_wind_model_list()
        
        # 更新风机总装机容量显示
        self.update_wind_total_capacity()
        
        # 显示成功消息
        messagebox.showinfo("成功", f"风机型号 '{name}' 已保存！")
        
        # 自动保存当前项目以确保修改被持久化
        self.save_current_project()
        
    def save_pv_model(self):
        """
        保存当前编辑的光伏型号
        """
        # 获取输入的参数
        name = self.pv_model_name.get().strip()
        if not name:
            messagebox.showwarning("警告", "请输入型号名称！")
            return
            
        method = self.pv_model_method.get()
        count = self.pv_model_count.get()
        # 获取出力修正系数
        correction_factor = self.pv_model_correction_factor.get()
        
        # 参数校验
        if count <= 0:
            messagebox.showerror("错误", "数量必须大于0！")
            return
            
        if correction_factor < 0:
            messagebox.showerror("错误", "出力修正系数不能为负数！")
            return
            
        # 创建型号数据
        model_data = {
            'name': name,
            'method': method,
            'count': count,
            'output_correction_factor': correction_factor  # 添加出力修正系数
        }
        
        # 根据不同方法设置参数
        if method == 'area_efficiency':
            panel_efficiency = self.pv_panel_efficiency.get()
            panel_area = self.pv_panel_area.get()
            
            if panel_efficiency <= 0 or panel_area <= 0:
                messagebox.showerror("错误", "光伏板效率和面积必须大于0！")
                return
                
            model_data['params'] = {
                'panel_efficiency': panel_efficiency,
                'panel_area': panel_area
            }
        elif method == 'installed_capacity':
            installed_capacity = self.pv_installed_capacity.get()
            system_efficiency = self.pv_system_efficiency.get()
            
            if installed_capacity <= 0 or system_efficiency <= 0:
                messagebox.showerror("错误", "装机容量和系统效率必须大于0！")
                return
                
            model_data['params'] = {
                'installed_capacity': installed_capacity,
                'system_efficiency': system_efficiency
            }
        
        # 如果是编辑现有型号
        if self.current_pv_editing_index is not None and self.current_pv_editing_index < len(self.data_model.pv_models):
            self.data_model.pv_models[self.current_pv_editing_index] = model_data
        # 如果是新增型号（理论上不会走到这里，因为新增型号会直接添加）
        else:
            self.data_model.pv_models.append(model_data)
            
        # 刷新列表
        self.refresh_pv_model_list()
        
        # 更新光伏总装机容量显示
        self.update_pv_total_capacity()
        
        # 显示成功消息
        messagebox.showinfo("成功", f"光伏型号 '{name}' 已保存！")
        
        # 自动保存当前项目以确保修改被持久化
        self.save_current_project()
        
    def add_wind_model(self):
        """
        添加新的风机型号
        """
        # 创建一个新的默认风机型号
        new_model = {
            'name': '新风机型号',
            'params': {
                'cut_in_wind': 3.0,
                'rated_wind': 12.0,
                'max_rated_wind': 18.0,
                'cut_out_wind': 25.0,
                'rated_power': 2000.0
            },
            'count': 1,
            'output_correction_factor': 1.0  # 添加出力修正系数，默认为1.0
        }
        
        # 添加到数据模型
        self.data_model.wind_turbine_models.append(new_model)
        
        # 刷新列表
        self.refresh_wind_model_list()
        
        # 选中新添加的型号
        self.wind_model_listbox.selection_clear(0, tk.END)
        self.wind_model_listbox.selection_set(tk.END)
        self.wind_model_listbox.activate(tk.END)
        self.wind_model_listbox.event_generate("<<ListboxSelect>>")
        
        # 绘制新添加型号的函数曲线
        self.plot_single_wind_curve(new_model)
        
        # 更新风机总装机容量显示
        self.update_wind_total_capacity()
        
    def add_pv_model(self):
        """
        添加新的光伏型号
        """
        # 创建一个新的默认光伏型号
        new_model = {
            'name': '新光伏型号',
            'method': 'area_efficiency',
            'params': {
                'panel_efficiency': 0.2,
                'panel_area': 1000.0
            },
            'count': 10,
            'output_correction_factor': 1.0  # 添加出力修正系数，默认为1.0
        }
        
        # 添加到数据模型
        self.data_model.pv_models.append(new_model)
        
        # 刷新列表
        self.refresh_pv_model_list()
        
        # 选中新添加的型号
        self.pv_model_listbox.selection_clear(0, tk.END)
        self.pv_model_listbox.selection_set(tk.END)
        self.pv_model_listbox.activate(tk.END)
        self.pv_model_listbox.event_generate("<<ListboxSelect>>")
        
        # 绘制新添加型号的函数曲线
        self.plot_single_pv_curve(new_model)
        
        # 更新光伏总装机容量显示
        self.update_pv_total_capacity()
        
    def save_function_parameters(self):
        """
        保存函数参数设置
        """
        try:
            # 保存热电联产参数
            self.data_model.chp_electric_params['electric_heat_ratio'] = self.electric_heat_ratio.get()
            self.data_model.chp_electric_params['base_electric'] = self.base_electric.get()
            
            # 保存调峰机组参数
            self.data_model.peak_power_max = self.peak_power_max.get()
            self.data_model.peak_power_min_summer = self.peak_power_min_summer.get()
            self.data_model.peak_power_min_winter = self.peak_power_min_winter.get()
            
            # 保存负荷设置
            self.data_model.max_electric_load = self.max_load_var.get()
            self.data_model.flexible_load_max = self.flexible_load_max_var.get()
            self.data_model.flexible_load_min = self.flexible_load_min_var.get()
            
            # 显示成功消息
            messagebox.showinfo("成功", "所有函数参数已保存！")
        except Exception as e:
            messagebox.showerror("错误", f"保存参数时发生错误：{str(e)}")
            
    def save_optimization_params(self):
        """
        保存优化参数设置
        """
        try:
            # 保存优化参数
            self.data_model.optimization_params['basic_load_revenue'] = self.basic_load_revenue.get()
            self.data_model.optimization_params['flexible_load_revenue'] = self.flexible_load_revenue.get()
            self.data_model.optimization_params['thermal_cost'] = self.thermal_cost.get()
            self.data_model.optimization_params['pv_cost'] = self.pv_cost.get()
            self.data_model.optimization_params['wind_cost'] = self.wind_cost.get()
            self.data_model.optimization_params['load_change_rate_limit'] = self.load_change_rate_limit.get()
            self.data_model.optimization_params['min_grid_load'] = self.min_grid_load.get()
            
            # 显示成功消息
            messagebox.showinfo("成功", "优化参数已保存！")
        except Exception as e:
            messagebox.showerror("错误", f"保存优化参数时发生错误：{str(e)}")
            
    def save_and_return_to_project_list(self):
        """
        保存当前项目并返回项目列表
        """
        self.save_current_project()
        self.return_to_project_list()
        
    def save_current_project(self):
        """
        保存当前项目数据
        """
        if self.current_project:
            # 准备项目数据，包含优化结果（如果存在）
            project_data = self.data_model.to_dict()
            
            # 如果存在优化结果，将其添加到项目数据中
            if hasattr(self, 'optimized_results') and self.optimized_results is not None:
                project_data['optimized_results'] = self.optimized_results
            
            # 保存数据模型到项目文件
            success = self.project_manager.save_project_data(
                self.current_project['id'], 
                project_data
            )
            
            if success:
                print(f"项目 '{self.current_project['name']}' 已保存")
            else:
                messagebox.showerror("错误", "保存项目数据失败！")

    def create_calculation_tab(self, notebook):
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="📈 平衡计算")  # 添加结果图标
        
        # 添加返回项目列表按钮
        back_btn = ttk.Button(tab, text="保存并返回项目列表", command=self.save_and_return_to_project_list)
        back_btn.grid(row=0, column=3, sticky=tk.E, padx=5, pady=5)
        
        # 计算控制
        control_frame = ttk.LabelFrame(tab, text="计算控制", padding="10")
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        ttk.Button(control_frame, text="开始年度平衡计算", command=self.start_calculation).grid(row=0, column=0, pady=10, padx=(0, 10))
        
        # 添加导出结果按钮
        ttk.Button(control_frame, text="导出计算结果", command=self.export_results).grid(row=0, column=1, pady=10, padx=(0, 10))
        
        # 进度条
        self.progress = ttk.Progressbar(control_frame, mode='determinate')
        self.progress.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.progress_label = ttk.Label(control_frame, text="准备就绪")
        self.progress_label.grid(row=2, column=0, columnspan=2, pady=5)
        
        # 结果展示
        result_frame = ttk.LabelFrame(tab, text="计算结果", padding="10")
        result_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        self.result_text = tk.Text(result_frame, height=8, width=100)
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 时间段选择区域
        time_range_frame = ttk.LabelFrame(tab, text="时间段选择", padding="10")
        time_range_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(time_range_frame, text="开始日期:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.result_start_date_var = tk.StringVar(value="2025-01-01")
        self.result_start_date_entry = ttk.Entry(time_range_frame, textvariable=self.result_start_date_var, width=12)
        self.result_start_date_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(time_range_frame, text="结束日期:").grid(row=0, column=2, sticky=tk.W, padx=(10, 5))
        self.result_end_date_var = tk.StringVar(value="2025-12-31")
        self.result_end_date_entry = ttk.Entry(time_range_frame, textvariable=self.result_end_date_var, width=12)
        self.result_end_date_entry.grid(row=0, column=3, padx=5)
        
        ttk.Button(time_range_frame, text="更新图表", command=self.update_plot).grid(row=0, column=4, padx=(10, 0))
        
        # 图表展示
        plot_frame = ttk.LabelFrame(tab, text="可视化展示", padding="10")
        plot_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 创建matplotlib图形
        self.figure = Figure(figsize=(10, 6), dpi=100)  # 增加高度
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, plot_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 配置权重
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(4, weight=1)  # 给图表区域分配更多空间
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        time_range_frame.columnconfigure(5, weight=1)
        plot_frame.columnconfigure(0, weight=1)
        plot_frame.rowconfigure(0, weight=1)
        
    def browse_file(self, var):
        filename = filedialog.askopenfilename(
            title="选择CSV文件",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            var.set(filename)
            
    def download_template(self):
        """
        下载CSV模板文件
        """
        try:
            # 获取当前工作目录
            import os
            template_path = os.path.join(os.path.dirname(__file__), "data_template.csv")
            
            # 检查模板文件是否存在
            if os.path.exists(template_path):
                # 询问用户保存位置
                save_path = filedialog.asksaveasfilename(
                    title="保存模板文件",
                    defaultextension=".csv",
                    filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                    initialfile="energy_data_template.csv"
                )
                
                if save_path:
                    # 复制模板文件到用户指定位置
                    shutil.copy2(template_path, save_path)
                    messagebox.showinfo("成功", f"模板文件已保存至:\n{save_path}")
                else:
                    # 用户取消操作
                    pass
            else:
                # 如果模板文件不存在，则动态创建一个
                self.create_template_file()
                # 再次检查模板文件是否存在
                if os.path.exists(template_path):
                    save_path = filedialog.asksaveasfilename(
                        title="保存模板文件",
                        defaultextension=".csv",
                        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                        initialfile="energy_data_template.csv"
                    )
                    
                    if save_path:
                        shutil.copy2(template_path, save_path)
                        messagebox.showinfo("成功", f"模板文件已保存至:\n{save_path}")
                else:
                    messagebox.showerror("错误", "无法创建或找到模板文件 data_template.csv")
        except Exception as e:
            messagebox.showerror("错误", f"下载模板文件时出错:\n{str(e)}")
    
    def create_template_file(self):
        """
        动态创建CSV模板文件
        """
        try:
            import csv
            import os
            
            # 定义文件路径
            template_path = os.path.join(os.path.dirname(__file__), "data_template.csv")
            
            # 表头 - 现在包含下网电价列
            headers = ['时间', '电力负荷(kW)', '热力负荷(kW)', '光照强度(W/m²)', '风速(m/s)', '下网电价(元/kWh)']
            
            # 创建8760小时的示例行（只显示前几行和后几行）
            rows = []
            # 添加前24小时示例
            for i in range(24):  
                time_str = f"2024-01-01 {i:02d}:00"
                row = [time_str, "0.0", "0.0", "0.0", "0.0", "0.0"]
                rows.append(row)
            
            # 添加说明文字
            rows.append(["..."] * 6)  # 占位符表示中间省略的行
            
            # 添加最后几行示例
            for i in range(24):  
                time_str = f"2024-12-31 {i:02d}:00"
                row = [time_str, "0.0", "0.0", "0.0", "0.0", "0.0"]
                rows.append(row)
            
            # 写入CSV文件，使用UTF-8编码并添加BOM
            with open(template_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(headers)
                writer.writerows(rows)
                
        except Exception as e:
            raise Exception(f"创建模板文件失败: {str(e)}")
            
    def import_all_data(self):
        try:
            # 更新厂用电率
            self.data_model.internal_electric_rate = self.internal_rate_var.get()
            
            # 清空之前的数据
            self.data_model.electric_load_hourly = [0.0] * 8760
            self.data_model.heat_load_hourly = [0.0] * 8760
            self.data_model.solar_irradiance_hourly = [0.0] * 8760
            self.data_model.wind_speed_hourly = [0.0] * 8760
            self.data_model.grid_purchase_price_hourly = [0.0] * 8760  # 同时清空下网电价数据
            
            # 使用单一文件导入模式
            if not self.single_file_path.get():
                messagebox.showerror("错误", "请选择统一数据文件!")
                return
            self.import_single_file_data()
            
            # 更新统计信息
            self.update_statistics()
            
            # 保存项目数据
            self.save_current_project()
            
            messagebox.showinfo("成功", "数据导入完成！")
            
        except Exception as e:
            messagebox.showerror("错误", f"数据导入失败: {str(e)}")
            
    def import_single_file_data(self):
        """
        从单一文件导入所有数据
        """
        import csv
        
        with open(self.single_file_path.get(), 'r', encoding='utf-8-sig') as csvfile:
            reader = csv.reader(csvfile)
            headers = next(reader)
            
            # 检查表头是否正确 - 现在支持包含下网电价的格式
            expected_headers_basic = ['时间', '电力负荷(kW)', '热力负荷(kW)', '光照强度(W/m²)', '风速(m/s)']
            expected_headers_with_price = ['时间', '电力负荷(kW)', '热力负荷(kW)', '光照强度(W/m²)', '风速(m/s)', '下网电价(元/kWh)']
            
            if headers != expected_headers_basic and headers != expected_headers_with_price:
                messagebox.showerror("错误", "文件表头不正确！请使用模板文件格式。\n期望格式: ['时间', '电力负荷(kW)', '热力负荷(kW)', '光照强度(W/m²)', '风速(m/s)'[, '下网电价(元/kWh)']]")
                return
            
            # 读取数据
            for i, row in enumerate(reader):
                if i >= 8760:
                    break
                self.data_model.electric_load_hourly[i] = float(row[1])
                self.data_model.heat_load_hourly[i] = float(row[2])
                self.data_model.solar_irradiance_hourly[i] = float(row[3])
                self.data_model.wind_speed_hourly[i] = float(row[4])
                
                # 如果文件包含下网电价数据，则导入
                if len(row) > 5:
                    self.data_model.grid_purchase_price_hourly[i] = float(row[5])
                    self.data_model.data_imported['grid_price'] = True  # 标记下网电价数据已导入
                else:
                    self.data_model.data_imported['grid_price'] = False  # 标记下网电价数据未导入
                
            # 标记所有数据类型为已导入（只要有数据就标记为导入）
            self.data_model.data_imported['electric'] = any(self.data_model.electric_load_hourly)
            self.data_model.data_imported['heat'] = any(self.data_model.heat_load_hourly)
            self.data_model.data_imported['solar'] = any(self.data_model.solar_irradiance_hourly)
            self.data_model.data_imported['wind'] = any(self.data_model.wind_speed_hourly)
                

                
    def refresh_wind_model_list(self):
        """
        刷新风机型号列表
        """
        self.wind_model_listbox.delete(0, tk.END)
        for model in self.data_model.wind_turbine_models:
            self.wind_model_listbox.insert(tk.END, model['name'])
        
        if self.current_editing_index is not None and self.current_editing_index < len(self.data_model.wind_turbine_models):
            self.wind_model_listbox.selection_clear(0, tk.END)
            self.wind_model_listbox.selection_set(self.current_editing_index)
            self.wind_model_listbox.activate(self.current_editing_index)
            # 更新函数图像
            self.plot_single_wind_curve(self.data_model.wind_turbine_models[self.current_editing_index])
        elif len(self.data_model.wind_turbine_models) > 0:
            # 如果没有选中项但有型号，则默认选中第一个
            self.wind_model_listbox.selection_clear(0, tk.END)
            self.wind_model_listbox.selection_set(0)
            self.wind_model_listbox.activate(0)
            self.wind_model_listbox.event_generate("<<ListboxSelect>>")
        
        # 更新风机总装机容量显示
        self.update_wind_total_capacity()
        
    def refresh_pv_model_list(self):
        """
        刷新光伏型号列表
        """
        self.pv_model_listbox.delete(0, tk.END)
        for model in self.data_model.pv_models:
            self.pv_model_listbox.insert(tk.END, model['name'])
        
        # 如果之前有选中的项，尝试重新选中它
        if self.current_pv_editing_index is not None and self.current_pv_editing_index < len(self.data_model.pv_models):
            self.pv_model_listbox.selection_clear(0, tk.END)
            self.pv_model_listbox.selection_set(self.current_pv_editing_index)
            self.pv_model_listbox.activate(self.current_pv_editing_index)
            # 更新函数图像
            self.plot_single_pv_curve(self.data_model.pv_models[self.current_pv_editing_index])
        elif len(self.data_model.pv_models) > 0:
            # 如果没有选中项但有型号，则默认选中第一个
            self.pv_model_listbox.selection_clear(0, tk.END)
            self.pv_model_listbox.selection_set(0)
            self.pv_model_listbox.activate(0)
            self.pv_model_listbox.event_generate("<<ListboxSelect>>")
        
        # 更新光伏总装机容量显示
        self.update_pv_total_capacity()
        
    def on_wind_model_select(self, event):
        """
        当选择风机型号时的回调函数
        """
        selection = self.wind_model_listbox.curselection()
        if selection:
            index = selection[0]
            self.current_editing_index = index
            
            model = self.data_model.wind_turbine_models[index]
            
            # 填充详情字段
            self.wind_model_name.set(model['name'])
            self.wind_model_cut_in.set(model['params']['cut_in_wind'])
            self.wind_model_rated.set(model['params']['rated_wind'])
            self.wind_model_max_rated.set(model['params']['max_rated_wind'])
            self.wind_model_cut_out.set(model['params']['cut_out_wind'])
            self.wind_model_rated_power.set(model['params']['rated_power'])
            self.wind_model_count.set(model['count'])
            
            # 设置出力修正系数
            self.wind_model_correction_factor.set(model.get('output_correction_factor', 1.0))
            
            # 绘制当前选中风机型号的函数曲线
            self.plot_single_wind_curve(model)
    
    def plot_single_wind_curve(self, model):
        """
        绘制单个风机型号的函数曲线
        """
        # 清除之前的图表
        self.wind_ax.clear()
        
        # 绘制风机出力函数曲线
        wind_speeds = np.linspace(0, 30, 300)  # 风速范围 0-30 m/s
        wind_powers = []
        
        for wind_speed in wind_speeds:
            power = wind_power_function(wind_speed, model['params'])
            wind_powers.append(power)
        
        # 绘制曲线
        self.wind_ax.plot(wind_speeds, wind_powers, '-', linewidth=2, color='blue')
        self.wind_ax.set_xlabel('风速 (m/s)')
        self.wind_ax.set_ylabel('出力 (kW)')
        self.wind_ax.set_title(f'{model["name"]} 出力函数')
        self.wind_ax.grid(True, alpha=0.3)
        
        # 刷新画布
        self.wind_function_canvas.draw()

    def on_pv_model_select(self, event):
        """
        当选择光伏型号时的回调函数
        """
        selection = self.pv_model_listbox.curselection()
        if selection:
            index = selection[0]
            self.current_pv_editing_index = index
            
            model = self.data_model.pv_models[index]
            
            # 填充详情字段
            self.pv_model_name.set(model['name'])
            self.pv_model_method.set(model['method'])
            self.pv_model_count.set(model['count'])
            
            # 根据计算方法填充参数
            params = model['params']
            if model['method'] == 'area_efficiency':
                self.pv_panel_efficiency.set(params.get('panel_efficiency', 0.2))
                self.pv_panel_area.set(params.get('panel_area', 1000.0))
            elif model['method'] == 'installed_capacity':
                self.pv_installed_capacity.set(params.get('installed_capacity', 200.0))
                self.pv_system_efficiency.set(params.get('system_efficiency', 0.9))
            
            # 更新界面显示
            self.on_pv_method_change()
            
            # 设置出力修正系数
            self.pv_model_correction_factor.set(model.get('output_correction_factor', 1.0))
            
            # 绘制当前选中光伏型号的函数曲线
            self.plot_single_pv_curve(model)
    
    def plot_single_pv_curve(self, model):
        """
        绘制单个光伏型号的函数曲线
        """
        # 清除之前的图表
        self.pv_ax.clear()
        
        # 绘制光伏出力函数曲线
        irradiances = np.linspace(0, 1200, 300)  # 光照强度范围 0-1200 W/m²
        pv_powers = []
        
        for irradiance in irradiances:
            power = pv_power_function(irradiance, model)
            pv_powers.append(power)
        
        # 绘制曲线
        self.pv_ax.plot(irradiances, pv_powers, '-', linewidth=2, color='orange')
        self.pv_ax.set_xlabel('光照强度 (W/m²)')
        self.pv_ax.set_ylabel('出力 (kW)')
        self.pv_ax.set_title(f'{model["name"]} 出力函数')
        self.pv_ax.grid(True, alpha=0.3)
        
        # 刷新画布
        self.pv_function_canvas.draw()

    def on_pv_method_change(self):
        """
        当光伏计算方法改变时的回调函数
        """
        method = self.pv_model_method.get()
        if method == 'area_efficiency':
            self.area_efficiency_frame.grid()
            self.installed_capacity_frame.grid_remove()
        elif method == 'installed_capacity':
            self.area_efficiency_frame.grid_remove()
            self.installed_capacity_frame.grid()
            
    def add_wind_model(self):
        """
        添加新的风机型号
        """
        # 创建一个新的默认风机型号
        new_model = {
            'name': '新风机型号',
            'params': {
                'cut_in_wind': 3.0,
                'rated_wind': 12.0,
                'max_rated_wind': 18.0,
                'cut_out_wind': 25.0,
                'rated_power': 2000.0
            },
            'count': 1,
            'output_correction_factor': 1.0  # 添加出力修正系数，默认为1.0
        }
        
        # 添加到数据模型
        self.data_model.wind_turbine_models.append(new_model)
        
        # 刷新列表
        self.refresh_wind_model_list()
        
        # 选中新添加的型号
        self.wind_model_listbox.selection_clear(0, tk.END)
        self.wind_model_listbox.selection_set(tk.END)
        self.wind_model_listbox.activate(tk.END)
        self.wind_model_listbox.event_generate("<<ListboxSelect>>")
        
        # 绘制新添加型号的函数曲线
        self.plot_single_wind_curve(new_model)
        
        # 更新风机总装机容量显示
        self.update_wind_total_capacity()
        
    def add_pv_model(self):
        """
        添加新的光伏型号
        """
        # 创建一个新的默认光伏型号
        new_model = {
            'name': '新光伏型号',
            'method': 'area_efficiency',
            'params': {
                'panel_efficiency': 0.2,
                'panel_area': 1000.0
            },
            'count': 10,
            'output_correction_factor': 1.0  # 添加出力修正系数，默认为1.0
        }
        
        # 添加到数据模型
        self.data_model.pv_models.append(new_model)
        
        # 刷新列表
        self.refresh_pv_model_list()
        
        # 选中新添加的型号
        self.pv_model_listbox.selection_clear(0, tk.END)
        self.pv_model_listbox.selection_set(tk.END)
        self.pv_model_listbox.activate(tk.END)
        self.pv_model_listbox.event_generate("<<ListboxSelect>>")
        
        # 绘制新添加型号的函数曲线
        self.plot_single_pv_curve(new_model)
        
        # 更新光伏总装机容量显示
        self.update_pv_total_capacity()

    def import_all_data(self):
        try:
            # 更新厂用电率
            self.data_model.internal_electric_rate = self.internal_rate_var.get()
            
            # 清空之前的数据
            self.data_model.electric_load_hourly = [0.0] * 8760
            self.data_model.heat_load_hourly = [0.0] * 8760
            self.data_model.solar_irradiance_hourly = [0.0] * 8760
            self.data_model.wind_speed_hourly = [0.0] * 8760
            self.data_model.grid_purchase_price_hourly = [0.0] * 8760  # 清空下网电价数据
            
            # 检查使用哪种导入模式
            if self.single_file_mode.get():
                # 使用单一文件导入模式
                if not self.single_file_path.get():
                    messagebox.showerror("错误", "请选择统一数据文件!")
                    return
                self.import_single_file_data()
            else:
                # 使用多文件导入模式
                self.import_multiple_files_data()
            
            # 更新统计信息
            self.update_statistics()
            
            # 保存项目数据
            self.save_current_project()
            
            messagebox.showinfo("成功", "数据导入完成！")
            
        except Exception as e:
            messagebox.showerror("错误", f"数据导入失败: {str(e)}")
            
    def import_single_file_data(self):
        """
        从单一文件导入所有数据
        """
        import csv
        
        with open(self.single_file_path.get(), 'r', encoding='utf-8-sig') as csvfile:
            reader = csv.reader(csvfile)
            headers = next(reader)
            
            # 检查表头是否正确 - 现在支持包含下网电价的表头
            expected_headers_basic = ['时间', '电力负荷(kW)', '热力负荷(kW)', '光照强度(W/m²)', '风速(m/s)']
            expected_headers_with_price = ['时间', '电力负荷(kW)', '热力负荷(kW)', '光照强度(W/m²)', '风速(m/s)', '下网电价(元/kWh)']
            
            if headers != expected_headers_basic and headers != expected_headers_with_price:
                messagebox.showerror("错误", "文件表头不正确！请使用模板文件格式。")
                return
            
            # 读取数据
            for i, row in enumerate(reader):
                if i >= 8760:
                    break
                self.data_model.electric_load_hourly[i] = float(row[1])
                self.data_model.heat_load_hourly[i] = float(row[2])
                self.data_model.solar_irradiance_hourly[i] = float(row[3])
                self.data_model.wind_speed_hourly[i] = float(row[4])
                
                # 如果表头包含下网电价列，则导入该数据
                if len(row) > 5 and headers == expected_headers_with_price:
                    self.data_model.grid_purchase_price_hourly[i] = float(row[5])
                    self.data_model.data_imported['grid_price'] = True
                else:
                    # 如果没有下网电价列，默认为0，标记为未导入
                    self.data_model.data_imported['grid_price'] = False
                
    def import_multiple_files_data(self):
        """
        从多个文件分别导入数据
        """
        try:
            # 检查是否所有必需的文件都已选择
            files_selected = [
                self.elec_load_file.get(),
                self.heat_load_file.get(),
                self.solar_file.get(),
                self.wind_file.get()
            ]
            
            selected_count = sum(1 for f in files_selected if f)
            
            if selected_count == 0:
                raise Exception("请至少选择一个数据文件!")
            
            # 分别读取各个文件（根据表头自动识别数据类型）
            if self.elec_load_file.get():
                self.read_csv_data(self.elec_load_file.get())
            if self.heat_load_file.get():
                self.read_csv_data(self.heat_load_file.get())
            if self.solar_file.get():
                self.read_csv_data(self.solar_file.get())
            if self.wind_file.get():
                self.read_csv_data(self.wind_file.get())
                
        except Exception as e:
            raise Exception(f"从多个文件导入数据时出错: {str(e)}")
    
    def read_csv_data(self, file_path, data_type=None, single_file=False):
        """
        读取CSV数据文件
        :param file_path: 文件路径
        :param data_type: 数据类型 (electric, heat, solar, wind)
        :param single_file: 是否为单一文件模式
        """
        try:
            import csv
            with open(file_path, 'r', encoding='utf-8-sig') as csvfile:
                reader = csv.reader(csvfile)
                headers = next(reader)  # 读取表头
                
                # 验证表头
                expected_headers = ['时间', '电力负荷(kW)', '热力负荷(kW)', '光照强度(W/m²)', '风速(m/s)']
                if single_file:
                    if headers != expected_headers:
                        raise Exception(f"文件列标题不匹配!\n期望: {expected_headers}\n实际: {headers}")
                
                # 读取数据行
                for i, row in enumerate(reader):
                    if i >= 8760:  # 最多读取8760小时的数据
                        break
                        
                    # 检查行是否有足够的列数
                    if len(row) < 2:
                        raise Exception(f"数据格式错误，第{i+2}行缺少数值列")
                    
                    value = float(row[1])
                    
                    # 根据表头自动判断数据类型
                    if data_type is None:
                        header = headers[1] if len(headers) > 1 else ""  # 获取数值列的表头
                        if "电力负荷" in header:
                            data_type = "electric"
                            self.data_model.data_imported['electric'] = True
                        elif "热力负荷" in header:
                            data_type = "heat"
                            self.data_model.data_imported['heat'] = True
                        elif "光照强度" in header:
                            data_type = "solar"
                            self.data_model.data_imported['solar'] = True
                        elif "风速" in header:
                            data_type = "wind"
                            self.data_model.data_imported['wind'] = True
                        else:
                            raise Exception(f"无法识别的数据类型: {header}")
                    else:
                        # 标记对应类型数据已导入
                        if data_type == "electric":
                            self.data_model.data_imported['electric'] = True
                        elif data_type == "heat":
                            self.data_model.data_imported['heat'] = True
                        elif data_type == "solar":
                            self.data_model.data_imported['solar'] = True
                        elif data_type == "wind":
                            self.data_model.data_imported['wind'] = True
                    
                    # 根据数据类型分配数据
                    if data_type == "electric":
                        self.data_model.electric_load_hourly[i] = value
                    elif data_type == "heat":
                        self.data_model.heat_load_hourly[i] = value
                    elif data_type == "solar":
                        self.data_model.solar_irradiance_hourly[i] = value
                    elif data_type == "wind":
                        self.data_model.wind_speed_hourly[i] = value
                            
        except FileNotFoundError:
            raise Exception(f"文件未找到: {file_path}")
        except ValueError as e:
            raise Exception(f"数据格式错误，请检查第{i+2}行: {str(e)}")
        except Exception as e:
            raise Exception(f"读取文件时出错: {str(e)}")
            
    def update_statistics(self):
        """更新数据统计信息"""
        # 检查哪些数据已导入
        imported_data = []
        if self.data_model.data_imported['electric']:
            imported_data.append("电力负荷")
        if self.data_model.data_imported['heat']:
            imported_data.append("热力负荷")
        if self.data_model.data_imported['solar']:
            imported_data.append("光照强度")
        if self.data_model.data_imported['wind']:
            imported_data.append("风速")
        if self.data_model.data_imported['grid_price']:
            imported_data.append("下网电价")
        
        imported_info = "已导入数据: " + ", ".join(imported_data) if imported_data else "未导入任何数据"
        
        stats = f"""数据统计信息:
{imported_info}

电力负荷: 最小 {min(self.data_model.electric_load_hourly):.2f} kW, 
          最大 {max(self.data_model.electric_load_hourly):.2f} kW, 
          平均 {np.mean(self.data_model.electric_load_hourly):.2f} kW

热力负荷: 最小 {min(self.data_model.heat_load_hourly):.2f} kW, 
          最大 {max(self.data_model.heat_load_hourly):.2f} kW, 
          平均 {np.mean(self.data_model.heat_load_hourly):.2f} kW

光照强度: 最小 {min(self.data_model.solar_irradiance_hourly):.2f} W/m², 
          最大 {max(self.data_model.solar_irradiance_hourly):.2f} W/m², 
          平均 {np.mean(self.data_model.solar_irradiance_hourly):.2f} W/m²

风速:     最小 {min(self.data_model.wind_speed_hourly):.2f} m/s, 
          最大 {max(self.data_model.wind_speed_hourly):.2f} m/s, 
          平均 {np.mean(self.data_model.wind_speed_hourly):.2f} m/s

下网电价: 最小 {min(self.data_model.grid_purchase_price_hourly):.2f} 元/kWh, 
          最大 {max(self.data_model.grid_purchase_price_hourly):.2f} 元/kWh, 
          平均 {np.mean(self.data_model.grid_purchase_price_hourly):.2f} 元/kWh

厂用电率: {self.data_model.internal_electric_rate*100:.2f}%
"""
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(tk.END, stats)
        
        # 更新导入数据趋势图
        self.update_imported_data_plot()
        
    def start_calculation(self):
        try:
            # 更新参数
            self.save_function_parameters()
            
            # 执行计算
            self.progress_label.config(text="正在计算...")
            self.progress["value"] = 0
            self.root.update_idletasks()
            
            # 实际计算过程
            self.results = self.calculator.calculate_annual_balance()
            
            # 更新进度
            self.progress["value"] = 100
            self.progress_label.config(text="计算完成")
            
            # 显示结果
            self.display_results()
            
            # 保存项目数据（包括计算结果）
            self.save_current_project()
            
            messagebox.showinfo("成功", "年度平衡计算完成！")
            
        except Exception as e:
            self.progress_label.config(text="计算失败")
            messagebox.showerror("错误", f"计算过程中出现错误: {str(e)}")
            
    def display_results(self):
        if not self.results:
            return
            
        # 计算统计信息
        internal_electric_load = self.results['hourly_internal_electric_load']
        total_load = self.results['hourly_total_load']
        chp_output = self.results['hourly_chp_output']
        pv_output = self.results['hourly_pv_output']
        wind_output = self.results['hourly_wind_output']
        # 修复 KeyError: 'hourly_peak_pending_output'
        peak_pending_output = self.results.get('hourly_peak_pending_output', [0.0] * 8760)
        peak_output = self.results['hourly_peak_output']
        thermal_output = self.results['hourly_thermal_output']
        generation = self.results['hourly_generation']
        wind_pv_abandon = self.results['hourly_wind_pv_abandon']
        grid_load = self.results['hourly_grid_load']
        abandon_rate = self.results['hourly_abandon_rate']
        corrected_electric_load = self.results['hourly_corrected_electric_load']
        
        # 计算各种统计数据
        grid_load_positive_hours = sum(1 for x in grid_load if x > 0)  # 需要下网的小时数
        grid_load_negative_hours = sum(1 for x in grid_load if x < 0)  # 可以上网的小时数
        
        total_grid_load_positive = sum(x for x in grid_load if x > 0)  # 总下网电量
        total_grid_load_negative = sum(abs(x) for x in grid_load if x < 0)  # 总上网电量
        total_wind_pv_abandon = sum(x for x in wind_pv_abandon)  # 总弃光弃风量
        total_pv_wind_output = sum(x for x in pv_output) + sum(x for x in wind_output)  # 总风光发电量
        avg_abandon_rate = np.mean(abandon_rate) * 100  # 平均弃光风率转为百分比
        
        # 计算弃光风率（按总电量计算）
        overall_abandon_rate = 0
        if total_pv_wind_output > 0:
            overall_abandon_rate = abs(total_wind_pv_abandon) / total_pv_wind_output * 100
            
        avg_internal_electric_load = np.mean(internal_electric_load)
        avg_total_load = np.mean(total_load)
        avg_chp_output = np.mean(chp_output)
        avg_pv_output = np.mean(pv_output)
        avg_wind_output = np.mean(wind_output)
        avg_peak_pending_output = np.mean(peak_pending_output)
        avg_peak_output = np.mean(peak_output)
        avg_thermal_output = np.mean(thermal_output)
        avg_generation = np.mean(generation)
        avg_wind_pv_abandon = np.mean(wind_pv_abandon)
        avg_grid_load = np.mean(grid_load)
        avg_corrected_electric_load = np.mean(corrected_electric_load)
        
        # 计算风机光伏实际出力的平均值
        wind_pv_actual = self.results['hourly_wind_pv_actual']
        avg_wind_pv_actual = np.mean(wind_pv_actual)
        
        result_text = f"""年度计算结果:

负荷分析:
  平均电力负荷: {np.mean(self.data_model.electric_load_hourly):.2f} kW
  平均修正后电力负荷: {avg_corrected_electric_load:.2f} kW
  平均厂用电负荷: {avg_internal_electric_load:.2f} kW
  平均总负荷: {avg_total_load:.2f} kW

发电出力分析:
  热定电机组平均出力: {avg_chp_output:.2f} kW
  光伏最大平均出力: {avg_pv_output:.2f} kW
  风机最大平均出力: {avg_wind_output:.2f} kW
  风机光伏实际平均出力: {avg_wind_pv_actual:.2f} kW
  调峰机组待定平均出力: {avg_peak_pending_output:.2f} kW
  调峰机组平均出力: {avg_peak_output:.2f} kW
  火电平均出力: {avg_thermal_output:.2f} kW
  总平均发电出力: {avg_generation:.2f} kW

弃光弃风分析:
  总弃光弃风量: {abs(total_wind_pv_abandon):.2f} kWh
  总风光发电量: {total_pv_wind_output:.2f} kWh
  总弃光风率: {overall_abandon_rate:.2f}%
  平均弃光风率: {avg_abandon_rate:.2f}%

供需平衡分析:
  需要下网小时数: {grid_load_positive_hours} 小时 ({grid_load_positive_hours/8760*100:.2f}%)
  可以上网小时数: {grid_load_negative_hours} 小时 ({grid_load_negative_hours/8760*100:.2f}%)
  总下网电量: {total_grid_load_positive:.2f} kWh
  总上网电量: {total_grid_load_negative:.2f} kWh
  平均下网负荷: {avg_grid_load:+.2f} kW
"""
        
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, result_text)
        
        # 更新图表
        self.update_plot()
        
    def update_plot(self):
        if not self.results:
            return
            
        # 清除之前的图表
        self.ax.clear()
        
        # 解析时间段
        try:
            start_date = datetime.strptime(self.result_start_date_var.get(), "%Y-%m-%d")
            end_date = datetime.strptime(self.result_end_date_var.get(), "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("错误", "日期格式不正确，请使用 YYYY-MM-DD 格式")
            return
        
        # 计算时间段对应的小时索引
        start_hour = self.date_to_hour(start_date)
        end_hour = self.date_to_hour(end_date)
        
        if start_hour > end_hour:
            messagebox.showerror("错误", "开始日期不能晚于结束日期")
            return
        
        if start_hour < 0 or end_hour >= 8760:
            messagebox.showerror("错误", "日期超出范围，应在2025-01-01至2025-12-31之间")
            return
        
        # 获取时间段内的数据
        hours = list(range(start_hour, end_hour + 1))
        total_load = [self.results['hourly_total_load'][i] for i in hours]
        generation = [self.results['hourly_generation'][i] for i in hours]
        grid_load = [self.results['hourly_grid_load'][i] for i in hours]
        
        # 处理可能缺失的新字段
        pv_output = [self.results.get('hourly_pv_output', [0.0] * 8760)[i] for i in hours]
        wind_output = [self.results.get('hourly_wind_output', [0.0] * 8760)[i] for i in hours]
        chp_output = [self.results.get('hourly_chp_output', [0.0] * 8760)[i] for i in hours]
        # 修复 KeyError: 'hourly_peak_pending_output'
        peak_pending_output = [self.results.get('hourly_peak_pending_output', [0.0] * 8760)[i] for i in hours]
        peak_output = [self.results.get('hourly_peak_output', [0.0] * 8760)[i] for i in hours]
        
        # 将小时转换为日期格式
        dates = [datetime(2025, 1, 1) + timedelta(hours=h) for h in hours]
        
        # 绘制各类出力组成
        line_total_load, = self.ax.plot(dates, total_load, label='总负荷', linewidth=0.5, color='blue')
        line_generation, = self.ax.plot(dates, generation, label='总出力', linewidth=0.5, color='green')
        line_grid_load, = self.ax.plot(dates, grid_load, label='下网负荷', linewidth=0.5, color='red')
        fill_pv = self.ax.fill_between(dates, [0]*len(dates), pv_output, label='光伏出力', alpha=0.3, color='orange')
        fill_wind = self.ax.fill_between(dates, [0]*len(dates), wind_output, label='风机出力', alpha=0.3, color='purple')
        fill_chp = self.ax.fill_between(dates, [0]*len(dates), chp_output, label='热电联产出力', alpha=0.3, color='brown')
        fill_peak = self.ax.fill_between(dates, [0]*len(dates), peak_output, label='调峰机组出力', alpha=0.3, color='cyan')
        
        self.ax.set_xlabel('日期 (MM-DD)')
        self.ax.set_ylabel('功率 (kW)')
        self.ax.set_title(f'能源供需趋势 ({self.result_start_date_var.get()} 至 {self.result_end_date_var.get()})')
        
        # 设置x轴日期格式
        self.ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%m-%d'))
        
        # 根据时间跨度自动选择适当的日期定位器
        date_span = (dates[-1] - dates[0]).days
        if date_span <= 31:  # 一个月内，使用周定位器
            self.ax.xaxis.set_major_locator(plt.matplotlib.dates.WeekdayLocator(interval=1))
        elif date_span <= 180:  # 6个月内，使用双周定位器
            self.ax.xaxis.set_major_locator(plt.matplotlib.dates.WeekdayLocator(interval=2))
        else:  # 超过6个月，使用月定位器
            self.ax.xaxis.set_major_locator(plt.matplotlib.dates.MonthLocator())
        
        # 旋转x轴标签以更好地显示
        plt.setp(self.ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
        
        # 创建图例并启用点击功能
        legend = self.ax.legend(loc='upper left', bbox_to_anchor=(1, 1))  # 固定图例位置
        
        # 启用图例点击交互
        lines = [line_total_load, line_generation, line_grid_load]
        fills = [fill_pv, fill_wind, fill_chp, fill_peak]
        
        # 为线条图例启用点击
        lined = {}
        for legline, origline in zip(legend.get_lines(), lines):
            legline.set_picker(True)  # Enable picking on the legend line
            lined[legline] = origline
        
        # 为填充图例启用点击
        for legline, origfill in zip(legend.get_patches(), fills):
            legline.set_picker(True)  # Enable picking on the legend patch
            lined[legline] = origfill
        
        # 为图例中的文本也启用点击
        for i, (legtext, origline) in enumerate(zip(legend.get_texts(), lines + [fill_pv, fill_wind, fill_chp, fill_peak])):
            legtext.set_picker(True)  # Enable picking on the legend text
            lined[legtext] = origline
        
        # 保存线条和图例映射关系
        self.lined_result = lined
        
        # 连接点击事件
        self.canvas.mpl_connect('pick_event', self.on_legend_click_result)
        
        # 连接鼠标移动事件以实现悬浮功能
        self.canvas.mpl_connect('motion_notify_event', self.on_result_hover)
        
        self.ax.grid(True, alpha=0.3)
        
        # 设置x轴范围
        self.ax.set_xlim(dates[0], dates[-1])
        
        # 调整子图参数以确保图例和标签完全显示
        self.figure.tight_layout()
        # 为图例预留额外空间
        self.figure.subplots_adjust(right=0.85)
        
        # 刷新画布
        self.canvas.draw()
        

            
    def generate_sample_data(self):
        """生成示例数据用于演示"""
        # 只在没有真实数据导入时才生成示例数据
        # 检查是否有真实数据
        has_real_data = any([
            any(self.data_model.electric_load_hourly),
            any(self.data_model.heat_load_hourly),
            any(self.data_model.solar_irradiance_hourly),
            any(self.data_model.wind_speed_hourly)
        ])
        
        if has_real_data:
            # 如果已有真实数据，则不需要生成示例数据
            return
            
        # 生成一年8760小时的示例数据
        hours = list(range(8760))
        
        # 电力负荷：基础负荷+日周期变化+季节变化
        for i in range(8760):
            # 日周期变化（假设白天负荷高）
            hour_of_day = i % 24
            daily_variation = 0.8 + 0.4 * np.sin((hour_of_day - 6) * np.pi / 12)
            
            # 季节变化（冬季和夏季负荷较高）
            day_of_year = i // 24
            seasonal_variation = 1.0 + 0.3 * np.sin((day_of_year - 80) * 2 * np.pi / 365)
            
            self.data_model.electric_load_hourly[i] = 1000 * daily_variation * seasonal_variation
            
        # 热力负荷：与室外温度相关（简化模型）
        for i in range(8760):
            day_of_year = i // 24
            # 简化的温度模型（1月和12月最冷，7月最热）
            temp_factor = 1.0 + 0.5 * np.cos((day_of_year - 15) * 2 * np.pi / 365)
            self.data_model.heat_load_hourly[i] = 500 * temp_factor
            
        # 光照强度：日变化，受天气和季节影响
        for i in range(8760):
            hour_of_day = i % 24
            day_of_year = i // 24
            
            # 日照时间随季节变化
            daylight_hours = 12 + 4 * np.cos((day_of_year - 80) * 2 * np.pi / 365)
            sunrise = 12 - daylight_hours / 2
            sunset = 12 + daylight_hours / 2
            
            if sunrise <= hour_of_day <= sunset:
                # 正午时光照最强
                solar_peak = np.sin((hour_of_day - sunrise) * np.pi / (sunset - sunrise))
                # 季节影响
                season_factor = 0.5 + 0.5 * np.sin((day_of_year - 80) * 2 * np.pi / 365)
                self.data_model.solar_irradiance_hourly[i] = 800 * solar_peak * season_factor
            else:
                self.data_model.solar_irradiance_hourly[i] = 0
                
        # 风速：随机波动
        for i in range(8760):
            # 平均风速随季节变化
            day_of_year = i // 24
            avg_wind = 5 + 2 * np.sin((day_of_year - 50) * 2 * np.pi / 365)
            # 随机波动
            self.data_model.wind_speed_hourly[i] = max(0, avg_wind + np.random.normal(0, 1.5))
            
        # 标记所有数据已导入（虽然是示例数据）
        self.data_model.data_imported['electric'] = True
        self.data_model.data_imported['heat'] = True
        self.data_model.data_imported['solar'] = True
        self.data_model.data_imported['wind'] = True
    
    def export_results(self):
        """
        导出计算结果到Excel文件
        """
        try:
            # 检查是否有计算结果
            if not self.results:
                messagebox.showwarning("警告", "请先进行计算再导出结果！")
                return
            
            # 获取项目名称用于构建默认文件名
            project_name = "未命名项目"
            if self.current_project and 'name' in self.current_project:
                project_name = self.current_project['name']
            
            # 清理项目名称，移除不适合作为文件名的字符
            import re
            clean_project_name = re.sub(r'[<>:"/\\|?*]', '_', project_name)
            
            # 询问用户保存位置
            save_path = filedialog.asksaveasfilename(
                title="保存计算结果",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=f"{clean_project_name}_计算结果.xlsx"
            )
            
            if not save_path:
                return  # 用户取消操作
            
            # 导出数据
            self._write_results_to_excel(save_path)
            
            messagebox.showinfo("成功", f"计算结果已导出至:\n{save_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出结果失败:\n{str(e)}")
    
    def _write_results_to_excel(self, file_path):
        """
        将计算结果写入Excel文件
        """
        # 检查openpyxl是否可用
        if openpyxl is None:
            messagebox.showerror("错误", "缺少openpyxl库，请先安装：pip install openpyxl")
            return
        
        try:
            from datetime import datetime, timedelta
            from openpyxl import Workbook
        except ImportError:
            messagebox.showerror("错误", "缺少openpyxl库，请先安装：pip install openpyxl")
            return
        
        # 创建工作簿
        wb = Workbook()
        
        # 获取默认工作表并重命名
        ws1 = wb.active
        ws1.title = "小时数据"
        
        # 写入表头
        headers = [
            '时间', '电力负荷(kW)', '热力负荷(kW)', '光照强度(W/m²)', '风速(m/s)', '修正后电力负荷(kW)',
            '厂用电负荷(kW)', '总负荷(kW)', '热定电机组出力(kW)', '光伏最大出力(kW)', 
            '风机最大出力(kW)', '调峰机组待定出力(kW)', '调峰机组出力(kW)', '火电出力(kW)', 
            '总出力(kW)', '风机光伏放弃出力(kW)', '风机光伏实际出力(kW)', '弃光风率', '下网负荷(kW)'
        ]
        ws1.append(headers)
        
        # 写入数据行
        base_date = datetime(2026, 1, 1)  # 从2026年开始
        for i in range(8760):
            # 构造准确的时间字符串
            current_time = base_date + timedelta(hours=i)
            time_str = current_time.strftime("%Y-%m-%d %H:%M")
            
            row = [
                time_str,
                self.data_model.electric_load_hourly[i],
                self.data_model.heat_load_hourly[i],
                self.data_model.solar_irradiance_hourly[i],
                self.data_model.wind_speed_hourly[i],
                self.results['hourly_corrected_electric_load'][i],  # 修正后电力负荷
                self.results['hourly_internal_electric_load'][i],
                self.results['hourly_total_load'][i],
                self.results['hourly_chp_output'][i],
                self.results['hourly_pv_output'][i],
                self.results['hourly_wind_output'][i],
                # 修复 KeyError: 'hourly_peak_pending_output'
                self.results.get('hourly_peak_pending_output', [0.0] * 8760)[i],
                self.results['hourly_peak_output'][i],
                self.results['hourly_thermal_output'][i],
                self.results['hourly_generation'][i],
                self.results['hourly_wind_pv_abandon'][i],
                self.results['hourly_wind_pv_actual'][i],
                self.results['hourly_abandon_rate'][i],
                self.results['hourly_grid_load'][i]
            ]
            ws1.append(row)
        
        # 创建第二个工作表 - 按月统计数据
        ws2 = wb.create_sheet("月度统计")
        
        # 月度统计数据表头
        # 按照：总用电量、总发电量、火电发电量、负荷用电量、厂用电量、光伏风电发电量、光伏风电消纳电量、弃电量、下网电量、弃风光率 排列
        monthly_headers = ['月份', '总用电量(kWh)', '总发电量(kWh)', '火电发电量(kWh)', '负荷用电量(kWh)', '厂用电量(kWh)', 
                         '光伏风电发电量(kWh)', '光伏风电消纳电量(kWh)', '弃电量(kWh)', '下网电量(kWh)', '弃风光率(%)']
        ws2.append(monthly_headers)
        
        # 计算每月统计数据
        monthly_stats = {}
        base_date = datetime(2026, 1, 1)
        
        # 初始化12个月的数据
        for month in range(1, 13):
            monthly_stats[month] = {
                'grid_load_sum': 0.0,          # 下网电量
                'abandon_sum': 0.0,            # 弃风光量
                'max_output_sum': 0.0,         # 光伏风电发电量(最大出力)
                'generation_sum': 0.0,         # 总发电量
                'total_load_sum': 0.0,         # 总用电量
                'wind_pv_abandon_sum': 0.0,    # 弃电量
                'thermal_sum': 0.0,            # 火电发电量
                'internal_electric_sum': 0.0,  # 厂用电量
                'wind_pv_actual_sum': 0.0,     # 光伏风电消纳电量
                'corrected_electric_sum': 0.0  # 负荷用电量（修正后电力负荷累加）
            }
        
        # 累计每个月的数据
        for i in range(8760):
            current_time = base_date + timedelta(hours=i)
            month = current_time.month
            
            # 累计下网电量（下网负荷相加）
            monthly_stats[month]['grid_load_sum'] += self.results['hourly_grid_load'][i]
            
            # 累计最大风光发电量
            pv_output = self.results['hourly_pv_output'][i]
            wind_output = self.results['hourly_wind_output'][i]
            monthly_stats[month]['max_output_sum'] += (pv_output + wind_output)
            
            # 累计总发电量
            monthly_stats[month]['generation_sum'] += self.results['hourly_generation'][i]
            
            # 累计总用电量
            monthly_stats[month]['total_load_sum'] += self.results['hourly_total_load'][i]
            
            # 累计弃电量（原弃电量 - 灵活负荷的消纳电量）
            hourly_original_abandon = max(self.results['hourly_wind_pv_abandon'][i], 0)  # 确保为非负数
            hourly_flexible_consumption = self.results['hourly_flexible_load_consumption'][i]
            adjusted_hourly_abandon = max(hourly_original_abandon - hourly_flexible_consumption, 0)
            monthly_stats[month]['wind_pv_abandon_sum'] += adjusted_hourly_abandon
            
            # 累计弃风光量（用于弃风光率计算，使用修正后的弃电量）
            monthly_stats[month]['abandon_sum'] = monthly_stats[month]['wind_pv_abandon_sum']
            
            # 累计火电发电量
            monthly_stats[month]['thermal_sum'] += self.results['hourly_thermal_output'][i]
            
            # 累计厂用电量
            monthly_stats[month]['internal_electric_sum'] += self.results['hourly_internal_electric_load'][i]
            
            # 累计光伏风电消纳电量（原消纳电量 + 灵活负荷的消纳电量）
            hourly_original_actual = (pv_output + wind_output) - hourly_original_abandon
            hourly_adjusted_actual = hourly_original_actual + hourly_flexible_consumption
            monthly_stats[month]['wind_pv_actual_sum'] += hourly_adjusted_actual
            
            # 累计负荷用电量（修正后电力负荷）
            monthly_stats[month]['corrected_electric_sum'] += self.results['hourly_corrected_electric_load'][i]
        
        # 计算年度汇总数据
        annual_totals = {
            'grid_load_sum': 0.0,          # 下网电量
            'abandon_sum': 0.0,            # 弃风光量
            'max_output_sum': 0.0,         # 光伏风电发电量(最大出力)
            'generation_sum': 0.0,         # 总发电量
            'total_load_sum': 0.0,         # 总用电量
            'wind_pv_abandon_sum': 0.0,    # 弃电量
            'thermal_sum': 0.0,            # 火电发电量
            'internal_electric_sum': 0.0,  # 厂用电量
            'wind_pv_actual_sum': 0.0,     # 光伏风电消纳电量
            'corrected_electric_sum': 0.0  # 负荷用电量（修正后电力负荷累加）
        }
        
        # 写入月度统计数据
        for month in range(1, 13):
            stats = monthly_stats[month]
            # 计算弃风光率
            abandon_rate = 0.0
            if stats['max_output_sum'] > 0:
                abandon_rate = (stats['abandon_sum'] / stats['max_output_sum']) * 100
            
            # 月份格式化
            month_str = f"2026-{month:02d}"
            # 按照：总用电量、总发电量、火电发电量、负荷用电量、厂用电量、光伏风电发电量、光伏风电消纳电量、弃电量、下网电量、弃风光率 排列
            row = [month_str, stats['total_load_sum'], stats['generation_sum'], stats['thermal_sum'],
                   stats['corrected_electric_sum'], stats['internal_electric_sum'], stats['max_output_sum'], 
                   stats['wind_pv_actual_sum'], stats['wind_pv_abandon_sum'], stats['grid_load_sum'], 
                   f"{abandon_rate:.2f}%"]
            ws2.append(row)
            
            # 累计年度总量
            annual_totals['grid_load_sum'] += stats['grid_load_sum']
            annual_totals['abandon_sum'] += stats['abandon_sum']
            annual_totals['max_output_sum'] += stats['max_output_sum']
            annual_totals['generation_sum'] += stats['generation_sum']
            annual_totals['total_load_sum'] += stats['total_load_sum']
            annual_totals['wind_pv_abandon_sum'] += stats['wind_pv_abandon_sum']
            annual_totals['thermal_sum'] += stats['thermal_sum']
            annual_totals['internal_electric_sum'] += stats['internal_electric_sum']
            annual_totals['wind_pv_actual_sum'] += stats['wind_pv_actual_sum']
            annual_totals['corrected_electric_sum'] += stats['corrected_electric_sum']
        
        # 计算年度弃风光率
        annual_abandon_rate = 0.0
        if annual_totals['max_output_sum'] > 0:
            annual_abandon_rate = (annual_totals['abandon_sum'] / annual_totals['max_output_sum']) * 100
        
        # 添加年度汇总行
        # 按照：总用电量、总发电量、火电发电量、负荷用电量、厂用电量、光伏风电发电量、光伏风电消纳电量、弃电量、下网电量、弃风光率 排列
        annual_row = ['年度汇总', annual_totals['total_load_sum'], annual_totals['generation_sum'], 
                      annual_totals['thermal_sum'], annual_totals['corrected_electric_sum'], 
                      annual_totals['internal_electric_sum'], annual_totals['max_output_sum'], 
                      annual_totals['wind_pv_actual_sum'], annual_totals['wind_pv_abandon_sum'], 
                      annual_totals['grid_load_sum'], f"{annual_abandon_rate:.2f}%"]
        ws2.append(annual_row)
        
        # 保存文件
        try:
            wb.save(file_path)
        except Exception as e:
            raise Exception(f"保存Excel文件时出错: {str(e)}")
    
    def _write_results_to_csv(self, file_path):
        """
        将计算结果写入CSV文件
        """
        import csv
        
        with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            
            # 写入表头
            headers = [
                '时间', '电力负荷(kW)', '热力负荷(kW)', '光照强度(W/m²)', '风速(m/s)',
                '厂用电负荷(kW)', '总负荷(kW)', '热定电机组出力(kW)', '光伏最大出力(kW)', 
                '风机最大出力(kW)', '调峰机组待定出力(kW)', '调峰机组出力(kW)', '火电出力(kW)', 
                '总出力(kW)', '风机光伏放弃出力(kW)', '风机光伏实际出力(kW)', '弃光风率', '下网负荷(kW)'
            ]
            writer.writerow(headers)
            
            # 写入数据行
            # 使用准确的2025年时间戳（根据规范要求）
            from datetime import datetime, timedelta
            base_date = datetime(2025, 1, 1)
            for i in range(8760):
                current_time = base_date + timedelta(hours=i)
                time_str = current_time.strftime("%Y-%m-%d %H:%M")
                
                row = [
                    time_str,
                    self.data_model.electric_load_hourly[i],
                    self.data_model.heat_load_hourly[i],
                    self.data_model.solar_irradiance_hourly[i],
                    self.data_model.wind_speed_hourly[i],
                    self.results['hourly_internal_electric_load'][i],
                    self.results['hourly_total_load'][i],
                    self.results['hourly_chp_output'][i],
                    self.results['hourly_pv_output'][i],
                    self.results['hourly_wind_output'][i],
                    # 修复 KeyError: 'hourly_peak_pending_output'
                    self.results.get('hourly_peak_pending_output', [0.0] * 8760)[i],
                    self.results['hourly_peak_output'][i],
                    self.results['hourly_thermal_output'][i],
                    self.results['hourly_generation'][i],
                    self.results['hourly_wind_pv_abandon'][i],
                    self.results['hourly_wind_pv_actual'][i],
                    self.results['hourly_abandon_rate'][i],
                    self.results['hourly_grid_load'][i]
                ]
                writer.writerow(row)

    def create_maintenance_schedule_tab(self, notebook):
        """
        创建检修和投产计划标签页
        """
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="🔧 检修投产")  # 添加扳手图标
        
        # 添加返回项目列表按钮
        back_btn = ttk.Button(tab, text="保存并返回项目列表", command=self.save_and_return_to_project_list)
        back_btn.grid(row=0, column=0, sticky=tk.E, padx=5, pady=5)
        
        # 创建主框架
        main_frame = ttk.Frame(tab)
        main_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 检修计划区域
        maintenance_frame = ttk.LabelFrame(main_frame, text="检修计划", padding="10")
        maintenance_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        
        # 检修计划表格
        maintenance_columns = ('计划名称', '影响负荷出力类型', '影响负荷出力大小(kW)', '开始日期', '结束日期')
        self.maintenance_tree = ttk.Treeview(maintenance_frame, columns=maintenance_columns, show='headings', height=8)
        
        # 定义列标题
        for col in maintenance_columns:
            self.maintenance_tree.heading(col, text=col)
            self.maintenance_tree.column(col, width=120)
        
        # 添加滚动条
        maintenance_scrollbar_y = ttk.Scrollbar(maintenance_frame, orient=tk.VERTICAL, command=self.maintenance_tree.yview)
        maintenance_scrollbar_x = ttk.Scrollbar(maintenance_frame, orient=tk.HORIZONTAL, command=self.maintenance_tree.xview)
        self.maintenance_tree.configure(yscrollcommand=maintenance_scrollbar_y.set, xscrollcommand=maintenance_scrollbar_x.set)
        
        # 布局
        self.maintenance_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        maintenance_scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        maintenance_scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 检修计划操作按钮
        maintenance_btn_frame = ttk.Frame(maintenance_frame)
        maintenance_btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(maintenance_btn_frame, text="添加", command=self.add_maintenance_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(maintenance_btn_frame, text="编辑", command=self.edit_maintenance_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(maintenance_btn_frame, text="删除", command=self.delete_maintenance_entry).pack(side=tk.LEFT, padx=5)
        
        # 投产计划区域
        commissioning_frame = ttk.LabelFrame(main_frame, text="投产计划", padding="10")
        commissioning_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 投产计划表格
        commissioning_columns = ('计划名称', '影响负荷出力类型', '影响负荷出力大小(kW)', '开始日期', '结束日期')
        self.commissioning_tree = ttk.Treeview(commissioning_frame, columns=commissioning_columns, show='headings', height=8)
        
        # 定义列标题
        for col in commissioning_columns:
            self.commissioning_tree.heading(col, text=col)
            self.commissioning_tree.column(col, width=120)
        
        # 添加滚动条
        commissioning_scrollbar_y = ttk.Scrollbar(commissioning_frame, orient=tk.VERTICAL, command=self.commissioning_tree.yview)
        commissioning_scrollbar_x = ttk.Scrollbar(commissioning_frame, orient=tk.HORIZONTAL, command=self.commissioning_tree.xview)
        self.commissioning_tree.configure(yscrollcommand=commissioning_scrollbar_y.set, xscrollcommand=commissioning_scrollbar_x.set)
        
        # 布局
        self.commissioning_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        commissioning_scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        commissioning_scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 投产计划操作按钮
        commissioning_btn_frame = ttk.Frame(commissioning_frame)
        commissioning_btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(commissioning_btn_frame, text="添加", command=self.add_commissioning_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(commissioning_btn_frame, text="编辑", command=self.edit_commissioning_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(commissioning_btn_frame, text="删除", command=self.delete_commissioning_entry).pack(side=tk.LEFT, padx=5)
        
        # 出力限制计划区域
        output_limit_frame = ttk.LabelFrame(main_frame, text="出力限制计划", padding="10")
        output_limit_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 0))
        
        # 出力限制计划表格
        output_limit_columns = ('计划名称', '限制类型', '限制出力大小(kW)', '开始日期', '结束日期')
        self.output_limit_tree = ttk.Treeview(output_limit_frame, columns=output_limit_columns, show='headings', height=8)
        
        # 定义列标题
        for col in output_limit_columns:
            self.output_limit_tree.heading(col, text=col)
            self.output_limit_tree.column(col, width=120)
        
        # 添加滚动条
        output_limit_scrollbar_y = ttk.Scrollbar(output_limit_frame, orient=tk.VERTICAL, command=self.output_limit_tree.yview)
        output_limit_scrollbar_x = ttk.Scrollbar(output_limit_frame, orient=tk.HORIZONTAL, command=self.output_limit_tree.xview)
        self.output_limit_tree.configure(yscrollcommand=output_limit_scrollbar_y.set, xscrollcommand=output_limit_scrollbar_x.set)
        
        # 布局
        self.output_limit_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        output_limit_scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        output_limit_scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 出力限制计划操作按钮
        output_limit_btn_frame = ttk.Frame(output_limit_frame)
        output_limit_btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(output_limit_btn_frame, text="添加", command=self.add_output_limit_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(output_limit_btn_frame, text="编辑", command=self.edit_output_limit_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(output_limit_btn_frame, text="删除", command=self.delete_output_limit_entry).pack(side=tk.LEFT, padx=5)
        
        # 配置权重
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        maintenance_frame.columnconfigure(0, weight=1)
        maintenance_frame.rowconfigure(0, weight=1)
        commissioning_frame.columnconfigure(0, weight=1)
        commissioning_frame.rowconfigure(0, weight=1)
        output_limit_frame.columnconfigure(0, weight=1)
        output_limit_frame.rowconfigure(0, weight=1)

    def create_optimization_tab(self, notebook):
        """
        创建优化标签页
        """
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="⚖️ 优化分析")  # 添加优化图标
        
        # 添加返回项目列表按钮
        back_btn = ttk.Button(tab, text="保存并返回项目列表", command=self.save_and_return_to_project_list)
        back_btn.grid(row=0, column=0, sticky=tk.E, padx=5, pady=5)
        
        # 创建主容器框架，用于放置参数设置和约束设置
        main_control_frame = ttk.Frame(tab)
        main_control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 优化参数设置区域
        params_frame = ttk.LabelFrame(main_control_frame, text="优化参数设置", padding="10")
        params_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        # 基础负荷单位收益
        ttk.Label(params_frame, text="基础负荷单位收益 (元/kWh): ").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.basic_load_revenue = tk.DoubleVar(value=self.data_model.optimization_params['basic_load_revenue'])  # 使用数据模型中的值
        ttk.Entry(params_frame, textvariable=self.basic_load_revenue, width=20).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # 灵活负荷单位收益
        ttk.Label(params_frame, text="灵活负荷单位收益 (元/kWh): ").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.flexible_load_revenue = tk.DoubleVar(value=self.data_model.optimization_params['flexible_load_revenue'])  # 使用数据模型中的值
        ttk.Entry(params_frame, textvariable=self.flexible_load_revenue, width=20).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # 火电发电单位成本
        ttk.Label(params_frame, text="火电发电单位成本 (元/kWh): ").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.thermal_cost = tk.DoubleVar(value=self.data_model.optimization_params['thermal_cost'])  # 使用数据模型中的值
        ttk.Entry(params_frame, textvariable=self.thermal_cost, width=20).grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # 光伏发电单位成本
        ttk.Label(params_frame, text="光伏发电单位成本 (元/kWh): ").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.pv_cost = tk.DoubleVar(value=self.data_model.optimization_params['pv_cost'])  # 使用数据模型中的值
        ttk.Entry(params_frame, textvariable=self.pv_cost, width=20).grid(row=3, column=1, sticky=tk.W, padx=5)
        
        # 风机发电单位成本
        ttk.Label(params_frame, text="风机发电单位成本 (元/kWh): ").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.wind_cost = tk.DoubleVar(value=self.data_model.optimization_params['wind_cost'])  # 使用数据模型中的值
        ttk.Entry(params_frame, textvariable=self.wind_cost, width=20).grid(row=4, column=1, sticky=tk.W, padx=5)
        
        # 约束设置区域
        constraint_frame = ttk.LabelFrame(main_control_frame, text="约束设置", padding="10")
        constraint_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 0))
        
        # 负荷最大变动率
        ttk.Label(constraint_frame, text="负荷最大变动率 (kW/Hr): ").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.load_change_rate_limit = tk.DoubleVar(value=self.data_model.optimization_params.get('load_change_rate_limit', 100000.0))  # 使用数据模型中的值
        ttk.Entry(constraint_frame, textvariable=self.load_change_rate_limit, width=20).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # 最小下网负荷
        ttk.Label(constraint_frame, text="最小下网负荷 (kW): ").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.min_grid_load = tk.DoubleVar(value=self.data_model.optimization_params.get('min_grid_load', 0.0))  # 使用数据模型中的值
        ttk.Entry(constraint_frame, textvariable=self.min_grid_load, width=20).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # 配置权重以让两个框架平分空间
        main_control_frame.columnconfigure(0, weight=1)
        main_control_frame.columnconfigure(1, weight=1)
        
        # 优化控制按钮
        control_frame = ttk.Frame(tab)
        control_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # 保存优化参数按钮
        ttk.Button(control_frame, text="保存优化参数", command=self.save_optimization_params).grid(row=0, column=0, padx=5, pady=10)
        ttk.Button(control_frame, text="开始优化计算", command=self.start_optimization).grid(row=0, column=1, padx=5, pady=10)
        ttk.Button(control_frame, text="导出优化结果", command=self.export_optimization_results).grid(row=0, column=2, padx=5, pady=10)
        ttk.Button(control_frame, text="更新趋势图", command=self.update_optimization_plot).grid(row=0, column=3, padx=5, pady=10)
        
        # 进度条
        self.optimization_progress = ttk.Progressbar(control_frame, mode='determinate', length=200)
        self.optimization_progress.grid(row=0, column=4, padx=10, pady=10)
        
        self.optimization_progress_label = ttk.Label(control_frame, text="")
        self.optimization_progress_label.grid(row=0, column=5, padx=5, pady=10)
        
        # 优化结果显示
        result_frame = ttk.LabelFrame(tab, text="优化结果", padding="10")
        result_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.optimization_result_text = tk.Text(result_frame, height=20, width=80)
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.optimization_result_text.yview)
        self.optimization_result_text.configure(yscrollcommand=scrollbar.set)
        
        self.optimization_result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 配置权重
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(3, weight=1)
        params_frame.columnconfigure(1, weight=1)
        constraint_frame.columnconfigure(1, weight=1)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        # 优化结果趋势图区域
        plot_optimization_frame = ttk.LabelFrame(tab, text="优化结果趋势图", padding="10")
        plot_optimization_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 时间段选择区域
        time_range_frame_opt = ttk.LabelFrame(plot_optimization_frame, text="时间段选择", padding="10")
        time_range_frame_opt.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(time_range_frame_opt, text="开始日期:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.optimization_start_date_var = tk.StringVar(value="2025-01-01")
        self.optimization_start_date_entry = ttk.Entry(time_range_frame_opt, textvariable=self.optimization_start_date_var, width=12)
        self.optimization_start_date_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(time_range_frame_opt, text="结束日期:").grid(row=0, column=2, sticky=tk.W, padx=(10, 5))
        self.optimization_end_date_var = tk.StringVar(value="2025-12-31")
        self.optimization_end_date_entry = ttk.Entry(time_range_frame_opt, textvariable=self.optimization_end_date_var, width=12)
        self.optimization_end_date_entry.grid(row=0, column=3, padx=5)
        
        ttk.Button(time_range_frame_opt, text="更新图表", command=self.update_optimization_plot).grid(row=0, column=4, padx=(10, 0))
        
        # 创建matplotlib图形
        self.optimization_figure = Figure(figsize=(10, 6), dpi=100)
        self.optimization_ax = self.optimization_figure.add_subplot(111)
        self.optimization_canvas = FigureCanvasTkAgg(self.optimization_figure, plot_optimization_frame)
        self.optimization_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 配置权重
        tab.rowconfigure(4, weight=1)

    def start_optimization(self):
        """
        开始优化计算
        """
        # 获取当前设置的参数
        basic_load_revenue = self.basic_load_revenue.get()
        flexible_load_revenue = self.flexible_load_revenue.get()
        thermal_cost = self.thermal_cost.get()
        pv_cost = self.pv_cost.get()
        wind_cost = self.wind_cost.get()
        
        # 检查是否有计算结果可供优化
        if not self.results:
            messagebox.showwarning("警告", "请先进行年度平衡计算，再进行优化！")
            return
        
        import numpy as np
        
        # 获取当前设置的参数
        basic_load_revenue = self.basic_load_revenue.get()
        flexible_load_revenue = self.flexible_load_revenue.get()
        thermal_cost = self.thermal_cost.get()
        pv_cost = self.pv_cost.get()
        wind_cost = self.wind_cost.get()
        
        # 检查是否有计算结果可供优化
        if not self.results:
            messagebox.showwarning("警告", "请先进行年度平衡计算，再进行优化！")
            return
        
        # 获取电价数据，使用导入的下网电价
        grid_price = self.data_model.grid_purchase_price_hourly
        
        # 创建优化后的结果字典
        optimized_results = {
            'hourly_basic_load': [0.0] * 8760,      # 基础负荷优化值
            'hourly_flexible_load': [0.0] * 8760,  # 灵活负荷优化值
            'hourly_revenue': [0.0] * 8760,        # 每小时收益
            'total_revenue': 0.0                     # 总收益
        }
        
        # 更新进度条
        self.optimization_progress_label.config(text="正在优化计算...")
        self.optimization_progress["value"] = 0
        self.root.update_idletasks()
        
        # 预计算可能重复使用的值
        from datetime import datetime, timedelta
        base_date = datetime(2024, 1, 1)
        
        # 获取约束参数
        load_change_rate_limit = self.load_change_rate_limit.get()
        min_grid_load = self.min_grid_load.get()
        
        # 逐小时优化
        for hour in range(8760):
            # 获取当前小时的其他参数（来自平衡计算结果）
            chp_output = self.results['hourly_chp_output'][hour]
            pv_output = self.results['hourly_pv_output'][hour]
            wind_output = self.results['hourly_wind_output'][hour]
            
            # 当前小时的约束条件
            max_flexible_load = self.data_model.flexible_load_max
            min_flexible_load = self.data_model.flexible_load_min
            
            # 获取当前小时的日期信息
            current_date = base_date + timedelta(hours=hour)
            current_month = current_date.month
            
            # 获取当前小时的活动检修计划和投产计划
            calculator = self.calculator
            active_maintenance_schedules = calculator.get_active_maintenance_schedules(hour)
            active_commissioning_schedules = calculator.get_active_commissioning_schedules(hour)
            
            # 预先计算修正后的调峰机组参数
            current_peak_power_max = self.data_model.peak_power_max
            if 5 <= current_month <= 9:  # 夏季
                current_peak_power_min = self.data_model.peak_power_min_summer
            else:  # 冬季
                current_peak_power_min = self.data_model.peak_power_min_winter
            
            # 应用检修计划对调峰机组参数的修正
            for schedule in active_maintenance_schedules:
                power_type = schedule.get('power_type', '')
                power_size = schedule.get('power_size', 0.0)
                
                if power_type == '调峰机组出力':
                    current_peak_power_max = current_peak_power_max - power_size
                    original_peak_power_max = self.data_model.peak_power_max
                    if original_peak_power_max > 0:
                        if 5 <= current_month <= 9:  # 夏季
                            current_peak_power_min = self.data_model.peak_power_min_summer * (current_peak_power_max / original_peak_power_max)
                        else:  # 冬季
                            current_peak_power_min = self.data_model.peak_power_min_winter * (current_peak_power_max / original_peak_power_max)
            
            # 应用投产计划对调峰机组参数的修正
            for schedule in active_commissioning_schedules:
                power_type = schedule.get('power_type', '')
                power_size = schedule.get('power_size', 0.0)
                start_date = schedule.get('start_date', '')
                end_date = schedule.get('end_date', '')
                
                interpolation_factor = calculator.calculate_interpolation_factor(hour, start_date, end_date)
                
                if power_type == '调峰机组最大出力':
                    adjusted_power_size = power_size * interpolation_factor
                    current_peak_power_max = current_peak_power_max - (power_size - adjusted_power_size)
                elif power_type == '调峰机组夏季最小出力':
                    adjusted_power_size = power_size * interpolation_factor
                    if 5 <= current_month <= 9:  # 夏季：5-9月
                        current_peak_power_min = self.data_model.peak_power_min_summer - adjusted_power_size
                elif power_type == '调峰机组冬季最小出力':
                    adjusted_power_size = power_size * interpolation_factor
                    if current_date.month < 5 or current_date.month > 9:  # 冬季：10-12月和1-4月
                        current_peak_power_min = self.data_model.peak_power_min_winter - adjusted_power_size
                elif power_type == '调峰机组最小出力':
                    adjusted_power_size = power_size * interpolation_factor
                    if 5 <= current_month <= 9:  # 夏季
                        current_peak_power_min = self.data_model.peak_power_min_summer - adjusted_power_size
                    else:  # 冬季
                        current_peak_power_min = self.data_model.peak_power_min_winter - adjusted_power_size
            
            # 获取平衡计算得到的电力负荷（考虑检修和投运计划修正后）作为基础负荷的最大值
            max_basic_load = self.results['hourly_corrected_electric_load'][hour]
            
            # 根据业务需求文档的优化算法思路实现
            # 1. 确定基础负荷范围：[current_peak_power_min, max_basic_load]
            min_basic_load = current_peak_power_min
            
            # 2. 初始化基础负荷和灵活负荷为当前平衡计算结果的比例
            current_total_load = self.results['hourly_corrected_electric_load'][hour]  # 当前总的电力负荷
            current_basic_load = current_total_load  # 当前基础负荷等于总电力负荷
            current_flexible_load = 0.0  # 当前灵活负荷为0
            
            # 3. 根据优化算法思路进行优化
            # 计算当前状态下的火电出力
            current_peak_pending = current_total_load - chp_output - pv_output - wind_output
            current_peak_output = max(min(current_peak_pending, current_peak_power_max), current_peak_power_min)
            current_thermal_output = chp_output + current_peak_output
            current_generation = pv_output + wind_output + current_thermal_output
            current_grid_load = current_total_load - current_generation  # 当前下网负荷
            
            # 判断是否有弃风弃电（当前下网负荷 <= 0）
            has_abandoned_energy = current_grid_load <= 0
            
            # 初始化优化负荷
            basic_load = max(min_basic_load, min(max_basic_load, current_basic_load))
            flexible_load = max(min_flexible_load, min(max_flexible_load, current_flexible_load))
            
            # 根据业务需求文档的算法思路进行优化
            if has_abandoned_energy:
                # 当有弃风弃电存在时（没有下网负荷），且负荷单位收益大于风电和光伏单位成本时，应提高基础负荷或者灵活负荷
                if basic_load_revenue > pv_cost and basic_load_revenue > wind_cost:
                    # 提高基础负荷到最大值（在约束范围内）
                    basic_load = max(min_basic_load, min(max_basic_load, max_basic_load))
                
                if flexible_load_revenue > pv_cost and flexible_load_revenue > wind_cost:
                    # 提高灵活负荷到最大值（在约束范围内）
                    flexible_load = max(min_flexible_load, min(max_flexible_load, max_flexible_load))
            else:
                # 当有下网负荷时（没有弃风弃电）
                # 当基础负荷单位收益大于下网电价时，应提高基础负荷，反之则减小
                if basic_load_revenue > grid_price[hour] if hour < len(grid_price) else 0:
                    basic_load = max(min_basic_load, min(max_basic_load, max_basic_load))
                else:
                    basic_load = max(min_basic_load, min(max_basic_load, min_basic_load))
                
                # 当灵活负荷单位收益大于下网电价时，应提高灵活负荷，反之则减小
                if flexible_load_revenue > grid_price[hour] if hour < len(grid_price) else 0:
                    flexible_load = max(min_flexible_load, min(max_flexible_load, max_flexible_load))
                else:
                    flexible_load = max(min_flexible_load, min(max_flexible_load, min_flexible_load))
            
            # 确保基础负荷不低于调峰机组最小出力
            basic_load = max(basic_load, min_basic_load)
            
            # 再次检查是否满足最小下网负荷约束，如果需要进一步调整
            total_load = basic_load + flexible_load
            peak_pending = total_load - chp_output - pv_output - wind_output
            peak_output = max(min(peak_pending, current_peak_power_max), current_peak_power_min)
            thermal_output = chp_output + peak_output
            generation = pv_output + wind_output + thermal_output
            grid_load = total_load - generation
            
            if grid_load < min_grid_load:
                # 如果仍然不满足最小下网负荷约束，需要进一步调整
                required_additional_load = min_grid_load - grid_load
                # 优先增加基础负荷，但如果超过最大值则分配给灵活负荷
                potential_basic_load = basic_load + required_additional_load
                if potential_basic_load <= max_basic_load:
                    basic_load = potential_basic_load
                else:
                    basic_load = max_basic_load
                    remaining_load = potential_basic_load - max_basic_load
                    flexible_load = min(flexible_load + remaining_load, max_flexible_load)
                
                # 重新计算所有相关值
                total_load = basic_load + flexible_load
                peak_pending = total_load - chp_output - pv_output - wind_output
                peak_output = max(min(peak_pending, current_peak_power_max), current_peak_power_min)
                thermal_output = chp_output + peak_output
                generation = pv_output + wind_output + thermal_output
                grid_load = total_load - generation
            
            # 应用到结果中
            optimized_results['hourly_basic_load'][hour] = basic_load
            optimized_results['hourly_flexible_load'][hour] = flexible_load
            
            # 计算收益
            total_load = basic_load + flexible_load
            peak_pending = total_load - chp_output - pv_output - wind_output
            peak_output = max(min(peak_pending, current_peak_power_max), current_peak_power_min)
            thermal_output = chp_output + peak_output
            generation = pv_output + wind_output + thermal_output
            grid_load = total_load - generation
            
            # 确保下网负荷不低于最小下网负荷
            if grid_load < min_grid_load:
                # 如果下网负荷低于最小值，需要调整负荷分配
                # 通过增加基础负荷来满足最小下网负荷约束
                required_additional_load = min_grid_load - grid_load
                basic_load = basic_load + required_additional_load
                total_load = basic_load + flexible_load
                
                # 重新计算调峰机组出力
                peak_pending = total_load - chp_output - pv_output - wind_output
                peak_output = max(min(peak_pending, current_peak_power_max), current_peak_power_min)
                thermal_output = chp_output + peak_output
                generation = pv_output + wind_output + thermal_output
                grid_load = total_load - generation
            
            revenue = (
                basic_load * basic_load_revenue + 
                flexible_load * flexible_load_revenue - 
                thermal_output * thermal_cost - 
                pv_output * pv_cost - 
                wind_output * wind_cost
            )
            
            if grid_load > 0:
                revenue -= grid_load * grid_price[hour] if hour < len(grid_price) else 0
            
            optimized_results['hourly_revenue'][hour] = revenue
            
            # 更新进度条
            if hour % 500 == 0:  # 每500小时更新一次进度
                progress = (hour / 8760) * 100
                self.optimization_progress["value"] = progress
                self.optimization_progress_label.config(text=f"正在优化计算... {int(progress)}%")
                self.root.update_idletasks()
        
        # 计算总收益
        total_revenue = sum(optimized_results['hourly_revenue'])
        optimized_results['total_revenue'] = total_revenue
        
        # 将优化结果存储到实例变量中
        self.optimized_results = optimized_results
        
        # 更新进度条到100%
        self.optimization_progress["value"] = 100
        self.optimization_progress_label.config(text="优化计算完成! 100%")
        self.root.update_idletasks()
        
        # 显示优化结果摘要
        avg_basic_load = np.mean(optimized_results['hourly_basic_load'])
        avg_flexible_load = np.mean(optimized_results['hourly_flexible_load'])
        
        result_text = f"""优化计算完成!

优化参数:
基础负荷单位收益: {basic_load_revenue} 元/kWh
灵活负荷单位收益: {flexible_load_revenue} 元/kWh
火电发电单位成本: {thermal_cost} 元/kWh
光伏发电单位成本: {pv_cost} 元/kWh
风机发电单位成本: {wind_cost} 元/kWh

优化结果:
总收益: {total_revenue:,.2f} 元
平均每小时收益: {total_revenue/8760:.2f} 元

基础负荷:
  平均值: {avg_basic_load:.2f} kW
  范围: {min(optimized_results['hourly_basic_load']):.2f} - {max(optimized_results['hourly_basic_load']):.2f} kW

灵活负荷:
  平均值: {avg_flexible_load:.2f} kW
  范围: {min(optimized_results['hourly_flexible_load']):.2f} - {max(optimized_results['hourly_flexible_load']):.2f} kW

说明:
- 优化目标为每小时收益最大化
- 基础负荷约束在 [当前季节调峰机组最小出力, 平衡计算得到的电力负荷（考虑了检修和投运计划修正）]
- 灵活负荷约束在 [最小灵活负荷, 最大灵活负荷]"""        
        self.optimization_result_text.delete(1.0, tk.END)
        self.optimization_result_text.insert(tk.END, result_text)
        
        # 保存优化结果到当前项目
        self.save_current_project()
        
        messagebox.showinfo("完成", "优化计算已完成！")


        
    def export_optimization_results(self):
        """
        导出优化结果
        """
        if not hasattr(self, 'optimized_results'):
            messagebox.showwarning("警告", "优化结果为空，无法导出！")
            return
        
        # 询问用户保存位置
        save_path = filedialog.asksaveasfilename(
            title="保存优化结果",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="optimized_results.xlsx"
        )
        
        if save_path:
            try:
                import pandas as pd
                from datetime import datetime, timedelta
                
                # 创建时间戳列表
                base_date = datetime(2025, 1, 1)
                time_stamps = [(base_date + timedelta(hours=i)).strftime("%Y-%m-%d %H:%M") for i in range(8760)]
                
                # 创建DataFrame
                df = pd.DataFrame({
                    '时间': time_stamps,
                    '基础负荷优化值(kW)': self.optimized_results['hourly_basic_load'],
                    '灵活负荷优化值(kW)': self.optimized_results['hourly_flexible_load'],
                    '每小时收益(元)': self.optimized_results['hourly_revenue']
                })
                
                # 导出到Excel
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='优化结果', index=False)
                    
                    # 添加汇总信息工作表
                    summary_data = [
                        ['项目', '值'],
                        ['总收益(元)', f'{self.optimized_results["total_revenue"]:.2f}'],
                        ['平均每小时收益(元)', f'{self.optimized_results["total_revenue"] / 8760:.2f}'],
                        ['基础负荷平均值(kW)', f'{sum(self.optimized_results["hourly_basic_load"]) / 8760:.2f}'],
                        ['灵活负荷平均值(kW)', f'{sum(self.optimized_results["hourly_flexible_load"]) / 8760:.2f}'],
                        ['基础负荷总计(kWh)', f'{sum(self.optimized_results["hourly_basic_load"]):.2f}'],
                        ['灵活负荷总计(kWh)', f'{sum(self.optimized_results["hourly_flexible_load"]):.2f}']
                    ]
                    
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, sheet_name='汇总信息', index=False, header=False)
                
                messagebox.showinfo("成功", f"优化结果已导出至:\n{save_path}")
            except ImportError:
                # 如果pandas不可用，使用文本格式导出
                txt_save_path = save_path.replace('.xlsx', '.txt')
                with open(txt_save_path, 'w', encoding='utf-8') as f:
                    f.write("优化结果\n\n")
                    f.write(f"总收益: {self.optimized_results['total_revenue']:.2f} 元\n")
                    f.write(f"平均每小时收益: {self.optimized_results['total_revenue']/8760:.2f} 元\n\n")
                    f.write("每小时优化结果 (前10小时示例):\n")
                    f.write("小时,基础负荷优化值(kW),灵活负荷优化值(kW),每小时收益(元)\n")
                    for i in range(min(10, 8760)):
                        f.write(f"{i},{self.optimized_results['hourly_basic_load'][i]:.2f},{self.optimized_results['hourly_flexible_load'][i]:.2f},{self.optimized_results['hourly_revenue'][i]:.2f}\n")
                messagebox.showinfo("成功", f"优化结果已导出至:\n{txt_save_path} (由于缺少pandas库，以文本格式导出)")
            except Exception as e:
                messagebox.showerror("错误", f"导出优化结果失败:\n{str(e)}")
        
    def update_optimization_plot(self):
        """
        更新优化结果趋势图
        包括优化前后的基础负荷、灵活负荷以及下网负荷对比
        """
        # 检查是否有优化结果和平衡计算结果
        if not hasattr(self, 'optimized_results') or not self.results:
            # 如果没有数据，显示提示信息
            self.optimization_ax.clear()
            self.optimization_ax.text(0.5, 0.5, '暂无优化结果\n请先进行年度平衡计算和优化计算', 
                             horizontalalignment='center', verticalalignment='center',
                             transform=self.optimization_ax.transAxes, fontsize=12)
            self.optimization_ax.set_title('优化结果趋势图')
            self.optimization_canvas.draw()
            return
        
        # 清除之前的图表
        self.optimization_ax.clear()
        
        # 解析时间段
        try:
            from datetime import datetime, timedelta
            start_date = datetime.strptime(self.optimization_start_date_var.get(), "%Y-%m-%d")
            end_date = datetime.strptime(self.optimization_end_date_var.get(), "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("错误", "日期格式不正确，请使用 YYYY-MM-DD 格式")
            return
        
        # 计算时间段对应的小时索引
        start_hour = self.date_to_hour(start_date)
        end_hour = self.date_to_hour(end_date)
        
        if start_hour > end_hour:
            messagebox.showerror("错误", "开始日期不能晚于结束日期")
            return
        
        if start_hour < 0 or end_hour >= 8760:
            messagebox.showerror("错误", "日期超出范围，应在2025-01-01至2025-12-31之间")
            return
        
        # 获取时间段内的数据
        hours = list(range(start_hour, end_hour + 1))
        
        # 获取优化结果
        optimized_basic_load = [self.optimized_results['hourly_basic_load'][i] for i in hours]
        optimized_flexible_load = [self.optimized_results['hourly_flexible_load'][i] for i in hours]
        
        # 优化后的总负荷是基础负荷和灵活负荷之和
        optimized_total_load = [optimized_basic_load[i] + optimized_flexible_load[i] for i in range(len(optimized_basic_load))]
        
        # 计算优化后的各发电设备出力
        try:
            # 计算优化后的光伏出力、风机出力和调峰机组出力
            optimized_pv_output = []
            optimized_wind_output = []
            optimized_peak_output = []
            optimized_grid_load = []
            
            for idx, i in enumerate(hours):
                # 获取原始的光伏和风机出力（这些不受负荷优化直接影响）
                pv_output = self.results['hourly_pv_output'][i]
                wind_output = self.results['hourly_wind_output'][i]
                chp_output = self.results['hourly_chp_output'][i]
                
                # 获取当前小时的修正后调峰机组参数
                from datetime import datetime, timedelta
                base_date = datetime(2024, 1, 1)
                current_date = base_date + timedelta(hours=i)
                current_month = current_date.month
                
                # 初始化当前小时的调峰机组参数
                current_peak_power_max = self.data_model.peak_power_max
                if 5 <= current_month <= 9:  # 夏季
                    current_peak_power_min = self.data_model.peak_power_min_summer
                else:  # 冬季
                    current_peak_power_min = self.data_model.peak_power_min_winter
                
                # 获取当前小时的活动检修计划和投产计划
                calculator = self.calculator
                active_maintenance_schedules = calculator.get_active_maintenance_schedules(i)
                active_commissioning_schedules = calculator.get_active_commissioning_schedules(i)
                
                # 应用检修计划对调峰机组参数的修正
                for schedule in active_maintenance_schedules:
                    power_type = schedule.get('power_type', '')
                    power_size = schedule.get('power_size', 0.0)
                    
                    if power_type == '调峰机组出力':
                        current_peak_power_max = current_peak_power_max - power_size
                        original_peak_power_max = self.data_model.peak_power_max
                        if original_peak_power_max > 0:
                            if 5 <= current_month <= 9:  # 夏季
                                current_peak_power_min = self.data_model.peak_power_min_summer * (current_peak_power_max / original_peak_power_max)
                            else:  # 冬季
                                current_peak_power_min = self.data_model.peak_power_min_winter * (current_peak_power_max / original_peak_power_max)
                
                # 应用投产计划对调峰机组参数的修正
                for schedule in active_commissioning_schedules:
                    power_type = schedule.get('power_type', '')
                    power_size = schedule.get('power_size', 0.0)
                    start_date = schedule.get('start_date', '')
                    end_date = schedule.get('end_date', '')
                    
                    interpolation_factor = calculator.calculate_interpolation_factor(i, start_date, end_date)
                    
                    if power_type == '调峰机组最大出力':
                        adjusted_power_size = power_size * interpolation_factor
                        current_peak_power_max = current_peak_power_max - (power_size - adjusted_power_size)
                    elif power_type == '调峰机组夏季最小出力':
                        adjusted_power_size = power_size * interpolation_factor
                        if 5 <= current_month <= 9:  # 夏季：5-9月
                            current_peak_power_min = self.data_model.peak_power_min_summer - adjusted_power_size
                    elif power_type == '调峰机组冬季最小出力':
                        adjusted_power_size = power_size * interpolation_factor
                        if current_date.month < 5 or current_date.month > 9:  # 冬季：10-12月和1-4月
                            current_peak_power_min = self.data_model.peak_power_min_winter - adjusted_power_size
                    elif power_type == '调峰机组最小出力':
                        adjusted_power_size = power_size * interpolation_factor
                        if 5 <= current_month <= 9:  # 夏季
                            current_peak_power_min = self.data_model.peak_power_min_summer - adjusted_power_size
                        else:  # 冬季
                            current_peak_power_min = self.data_model.peak_power_min_winter - adjusted_power_size
                
                # 计算优化后的调峰机组出力
                peak_pending = optimized_total_load[idx] - chp_output - pv_output - wind_output
                peak_output = max(min(peak_pending, current_peak_power_max), current_peak_power_min)
                
                # 保存计算结果
                optimized_pv_output.append(pv_output)
                optimized_wind_output.append(wind_output)
                optimized_peak_output.append(peak_output)
                
                # 计算优化后的下网负荷
                thermal_output = chp_output + peak_output
                generation = pv_output + wind_output + thermal_output
                optimized_grid_load_val = optimized_total_load[idx] - generation
                optimized_grid_load.append(optimized_grid_load_val)
                
        except Exception as e:
            print(f"计算优化后发电出力时出错: {e}")
            # 如果计算出错，使用原始的发电出力数据
            optimized_pv_output = [self.results['hourly_pv_output'][i] for i in hours]
            optimized_wind_output = [self.results['hourly_wind_output'][i] for i in hours]
            optimized_peak_output = [self.results['hourly_peak_output'][i] for i in hours]
            optimized_grid_load = [self.results['hourly_grid_load'][i] for i in hours]
        
        # 将小时转换为日期格式（从2025-01-01开始）
        from datetime import datetime, timedelta
        dates = [datetime(2025, 1, 1) + timedelta(hours=h) for h in hours]
        
        # 绘制优化后的发电出力图
        line_opt_basic, = self.optimization_ax.plot(dates, optimized_basic_load, label='基础负荷(优化后)', linewidth=0.8, color='blue', linestyle='-')
        line_opt_flex, = self.optimization_ax.plot(dates, optimized_flexible_load, label='灵活负荷(优化后)', linewidth=0.8, color='orange', linestyle='-')
        line_pv_output, = self.optimization_ax.plot(dates, optimized_pv_output, label='光伏出力(优化后)', linewidth=0.8, color='green', linestyle='-')
        line_wind_output, = self.optimization_ax.plot(dates, optimized_wind_output, label='风机出力(优化后)', linewidth=0.8, color='cyan', linestyle='-')
        line_peak_output, = self.optimization_ax.plot(dates, optimized_peak_output, label='调峰机组出力(优化后)', linewidth=0.8, color='purple', linestyle='-')
        line_opt_grid, = self.optimization_ax.plot(dates, optimized_grid_load, label='下网负荷(优化后)', linewidth=0.8, color='red', linestyle='-')
        
        self.optimization_ax.set_xlabel('日期 (MM-DD)')
        self.optimization_ax.set_ylabel('功率 (kW)')
        self.optimization_ax.set_title('优化后负荷与发电出力趋势图')
        
        # 设置x轴日期格式
        self.optimization_ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%m-%d'))
        
        # 根据时间跨度自动选择适当的日期定位器
        date_span = (dates[-1] - dates[0]).days
        if date_span <= 31:  # 一个月内，使用日定位器
            self.optimization_ax.xaxis.set_major_locator(plt.matplotlib.dates.DayLocator(interval=1))
        elif date_span <= 180:  # 6个月内，使用周定位器
            self.optimization_ax.xaxis.set_major_locator(plt.matplotlib.dates.DayLocator(interval=7))
        elif date_span <= 365:  # 一年内，使用双周定位器
            self.optimization_ax.xaxis.set_major_locator(plt.matplotlib.dates.DayLocator(interval=14))
        else:  # 超过一年，使用月定位器
            self.optimization_ax.xaxis.set_major_locator(plt.matplotlib.dates.MonthLocator())
        
        # 旋转x轴标签以更好地显示
        plt.setp(self.optimization_ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
        
        # 创建图例并启用点击功能
        legend = self.optimization_ax.legend(loc='upper left', bbox_to_anchor=(1, 1))  # 固定图例位置
        
        # 启用图例点击交互
        lines = [line_opt_basic, line_opt_flex, line_pv_output, line_wind_output, line_peak_output, line_opt_grid]
        
        # 为线条图例启用点击
        lined = {}
        for legline, origline in zip(legend.get_lines(), lines):
            legline.set_picker(True)  # Enable picking on the legend line
            lined[legline] = origline
        
        # 为图例中的文本也启用点击
        for legtext, origline in zip(legend.get_texts(), lines):
            legtext.set_picker(True)  # Enable picking on the legend text
            lined[legtext] = origline
        
        # 保存线条和图例映射关系
        self.lined_optimization = lined
        
        # 连接点击事件
        self.optimization_canvas.mpl_connect('pick_event', self.on_legend_click_optimization)
        
        # 连接鼠标移动事件以实现悬浮功能
        self.optimization_canvas.mpl_connect('motion_notify_event', self.on_optimization_hover)
        
        self.optimization_ax.grid(True, alpha=0.3)
        
        # 设置x轴范围
        self.optimization_ax.set_xlim(dates[0], dates[-1])
        
        # 调整子图参数以确保图例和标签完全显示
        self.optimization_figure.tight_layout()
        # 为图例预留额外空间
        self.optimization_figure.subplots_adjust(right=0.85)
        
        # 刷新画布
        self.optimization_canvas.draw()
        
    def on_legend_click_optimization(self, event):
        """
        处理优化结果图表图例点击事件，显示/隐藏对应的曲线
        """
        # 获取被点击的图例元素
        legitem = event.artist
        
        # 获取对应的原始线条
        origline = self.lined_optimization[legitem]
        
        # 切换线条的可见性
        vis = not origline.get_visible()
        origline.set_visible(vis)
        
        # 更新所有相关的图例元素的透明度
        legend = self.optimization_ax.get_legend()
        if legend:
            for legline, origline_ref in self.lined_optimization.items():
                if origline_ref == origline:
                    if vis:
                        legline.set_alpha(1.0)
                    else:
                        legline.set_alpha(0.2)
        
        # 刷新画布
        self.optimization_canvas.draw()
        
    def add_maintenance_entry(self):
        """
        添加检修计划条目
        """
        self.manage_schedule_entry("maintenance", "添加检修计划")
        
    def edit_maintenance_entry(self):
        """
        编辑检修计划条目
        """
        selection = self.maintenance_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个检修计划条目！")
            return
        self.manage_schedule_entry("maintenance", "编辑检修计划", selection[0])
        
    def delete_maintenance_entry(self):
        """
        删除检修计划条目
        """
        selection = self.maintenance_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个检修计划条目！")
            return
            
        if messagebox.askyesno("确认删除", "确定要删除选中的检修计划条目吗？"):
            # 使用更可靠的方法：首先获取UI树中的项目索引，然后从数据模型中删除相同索引的项目
            all_items = self.maintenance_tree.get_children()
            current_index = all_items.index(selection[0]) if selection[0] in all_items else -1
            
            # 如果UI和数据模型保持同步，我们可以按索引删除
            if current_index != -1 and current_index < len(self.data_model.maintenance_schedules):
                del self.data_model.maintenance_schedules[current_index]
            else:
                # 如果索引无效，使用原来的方法
                # 获取要删除的条目信息
                item = self.maintenance_tree.item(selection[0])
                values = item['values']
                
                # 从数据模型中删除对应条目
                if len(values) >= 3:  # 确保有足够的数据 (name, power_type, power_size)
                    name = values[0]  # 计划名称是第1个元素
                    power_size = float(values[2])  # 现在power_size是第3个元素（索引为2）
                    start_date = values[3]  # 开始日期是第4个元素
                    # 查找并删除匹配的条目
                    for i, sched in enumerate(self.data_model.maintenance_schedules):
                        if (sched.get('name', '') == name and
                            sched.get('power_size', 0) == power_size and 
                            sched.get('start_date', '') == start_date):
                            del self.data_model.maintenance_schedules[i]
                            break
            
            # 从UI中删除条目
            self.maintenance_tree.delete(selection[0])
            
            # 保存项目数据
            self.save_current_project()
        
    def add_commissioning_entry(self):
        """
        添加投产计划条目
        """
        self.manage_schedule_entry("commissioning", "添加投产计划")
        
    def edit_commissioning_entry(self):
        """
        编辑投产计划条目
        """
        selection = self.commissioning_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个投产计划条目！")
            return
        self.manage_schedule_entry("commissioning", "编辑投产计划", selection[0])
        
    def delete_commissioning_entry(self):
        """
        删除投产计划条目
        """
        selection = self.commissioning_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个投产计划条目！")
            return
            
        if messagebox.askyesno("确认删除", "确定要删除选中的投产计划条目吗？"):
            # 获取要删除的条目信息
            item = self.commissioning_tree.item(selection[0])
            values = item['values']
            
            # 从数据模型中删除对应条目
            # 使用更可靠的方法：首先获取UI树中的项目索引，然后从数据模型中删除相同索引的项目
            all_items = self.commissioning_tree.get_children()
            current_index = all_items.index(selection[0]) if selection[0] in all_items else -1
            
            # 如果UI和数据模型保持同步，我们可以按索引删除
            if current_index != -1 and current_index < len(self.data_model.commissioning_schedules):
                del self.data_model.commissioning_schedules[current_index]
            else:
                # 如果索引无效，使用原来的方法
                if len(values) >= 3:  # 确保有足够的数据 (name, power_type, power_size)
                    name = values[0]  # 计划名称是第1个元素
                    power_size = float(values[2])  # 现在power_size是第3个元素（索引为2）
                    start_date = values[3]  # 开始日期是第4个元素
                    # 查找并删除匹配的条目
                    for i, sched in enumerate(self.data_model.commissioning_schedules):
                        if (sched.get('name', '') == name and
                            sched.get('power_size', 0) == power_size and 
                            sched.get('start_date', '') == start_date):
                            del self.data_model.commissioning_schedules[i]
                            break
            
            # 从UI中删除条目
            self.commissioning_tree.delete(selection[0])
            
            # 保存项目数据
            self.save_current_project()
        
    def manage_schedule_entry(self, schedule_type, title, item_id=None):
        """
        管理检修或投产计划条目
        """
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x350")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 获取当前值（如果是编辑模式）
        current_values = []
        if item_id:
            tree = self.maintenance_tree if schedule_type == "maintenance" else self.commissioning_tree
            current_values = tree.item(item_id, 'values')
        
        # 计划名称
        ttk.Label(dialog, text="计划名称:").grid(row=0, column=0, sticky=tk.W, pady=5)
        name_var = tk.StringVar(value=current_values[0] if current_values and len(current_values) > 0 else "")
        ttk.Entry(dialog, textvariable=name_var, width=20).grid(row=0, column=1, pady=5, padx=5, sticky=tk.W)
        
        # 影响负荷出力类型
        ttk.Label(dialog, text="影响负荷出力类型:").grid(row=1, column=0, sticky=tk.W, pady=5)
        power_type_var = tk.StringVar(value=current_values[1] if current_values and len(current_values) > 1 else "")
        power_type_combo = ttk.Combobox(dialog, textvariable=power_type_var, state="readonly")
        if schedule_type == "maintenance":
            power_type_combo['values'] = ("光伏出力", "风机出力", "调峰机组出力", "用电负荷")
        else:  # commissioning
            power_type_combo['values'] = ("光伏出力", "风机出力", "调峰机组最大出力", "调峰机组夏季最小出力", "调峰机组冬季最小出力", "调峰机组最小出力", "用电负荷")
        power_type_combo.grid(row=1, column=1, pady=5, padx=5, sticky=tk.W)
        
        # 影响负荷出力大小
        ttk.Label(dialog, text="影响负荷出力大小(kW): ").grid(row=2, column=0, sticky=tk.W, pady=5)
        power_size_var = tk.DoubleVar(value=float(current_values[2]) if current_values and len(current_values) > 2 else 0.0)
        ttk.Entry(dialog, textvariable=power_size_var, width=20).grid(row=2, column=1, pady=5, padx=5, sticky=tk.W)
        
        # 开始日期
        ttk.Label(dialog, text="开始日期: ").grid(row=3, column=0, sticky=tk.W, pady=5)
        start_date_var = tk.StringVar(value=current_values[3] if current_values and len(current_values) > 3 else "")
        ttk.Entry(dialog, textvariable=start_date_var, width=20).grid(row=3, column=1, pady=5, padx=5, sticky=tk.W)
        ttk.Label(dialog, text="格式: YYYY-MM-DD", foreground="gray").grid(row=4, column=1, sticky=tk.W, pady=0)
        
        # 结束日期
        ttk.Label(dialog, text="结束日期: ").grid(row=5, column=0, sticky=tk.W, pady=5)
        end_date_var = tk.StringVar(value=current_values[4] if current_values and len(current_values) > 4 else "")
        ttk.Entry(dialog, textvariable=end_date_var, width=20).grid(row=5, column=1, pady=5, padx=5, sticky=tk.W)
        ttk.Label(dialog, text="格式: YYYY-MM-DD", foreground="gray").grid(row=6, column=1, sticky=tk.W, pady=0)
        
        # 设备型号选择（仅对检修计划中的光伏和风机有效）
        model_frame = ttk.Frame(dialog)
        model_frame.grid(row=7, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        
        ttk.Label(model_frame, text="设备型号: ").grid(row=0, column=0, sticky=tk.W)
        model_var = tk.StringVar()
        model_combo = ttk.Combobox(model_frame, textvariable=model_var, state="disabled", width=18)
        model_combo.grid(row=0, column=1, sticky=tk.W)
        
        def update_model_combo(event=None):
            power_type = power_type_var.get()
            # 仅在检修计划中启用设备型号选择
            if schedule_type == "maintenance" and power_type in ["光伏出力", "风机出力"]:
                model_combo.config(state="readonly")
                if power_type == "光伏出力":
                    model_combo['values'] = [model['name'] for model in self.data_model.pv_models]
                else:  # 风机出力
                    model_combo['values'] = [model['name'] for model in self.data_model.wind_turbine_models]
            else:
                model_combo.config(state="disabled")
                
        power_type_combo.bind('<<ComboboxSelected>>', update_model_combo)
        update_model_combo()  # 初始化
        
        # 如果是编辑模式且有设备型号，则设置设备型号
        if current_values and len(current_values) > 5:
            model_var.set(current_values[5])
        
        # 确定和取消按钮
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=8, column=0, columnspan=2, pady=20)
        
        def confirm():
            # 获取输入值
            name = name_var.get().strip()
            power_type = power_type_var.get()
            power_size = power_size_var.get()
            start_date = start_date_var.get()
            end_date = end_date_var.get()
            
            # 验证输入
            if not name:
                messagebox.showwarning("警告", "请输入计划名称！")
                return
                
            if not power_type:
                messagebox.showwarning("警告", "请选择影响负荷出力类型！")
                return
                
            if power_size <= 0:
                messagebox.showwarning("警告", "影响负荷出力大小必须大于0！")
                return
                
            if not start_date or not end_date:
                messagebox.showwarning("警告", "请输入开始日期和结束日期！")
                return
                
            # 验证日期格式
            try:
                datetime.strptime(start_date, "%Y-%m-%d")
                datetime.strptime(end_date, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("错误", "日期格式不正确！请使用 YYYY-MM-DD 格式。")
                return
                
            # 构造条目值
            values = [name, power_type, str(power_size), start_date, end_date]
            if power_type in ["光伏出力", "风机出力"] and model_var.get():
                values.append(model_var.get())
                
            # 添加或更新条目
            tree = self.maintenance_tree if schedule_type == "maintenance" else self.commissioning_tree
            if item_id:
                # 更新现有条目
                tree.item(item_id, values=values)
            else:
                # 添加新条目
                tree.insert('', tk.END, values=values)
                
            # 同步更新数据模型
            schedule_data = {
                'name': name,
                'power_type': power_type,
                'power_size': power_size,
                'start_date': start_date,
                'end_date': end_date
            }
            
            # 如果选择了设备型号，则添加到数据中
            if power_type in ["光伏出力", "风机出力"] and model_var.get():
                schedule_data['model'] = model_var.get()
                
            if schedule_type == "maintenance":
                if item_id:
                    # 更新现有数据（使用UI Treeview中的索引）
                    all_items = tree.get_children()
                    current_index = all_items.index(item_id) if item_id in all_items else -1
                    
                    # 如果能确定索引，就按索引更新（假定UI和数据模型顺序一致）
                    if current_index != -1 and current_index < len(self.data_model.maintenance_schedules):
                        # 直接按索引更新，这是最可靠的方法
                        self.data_model.maintenance_schedules[current_index] = schedule_data
                    else:
                        # 如果索引无效，尝试使用名称等字段匹配
                        item = tree.item(item_id)
                        old_values = item['values']
                        if len(old_values) >= 3:
                            old_name = old_values[0]  # 计划名称是第1个元素
                            old_power_size = float(old_values[2])  # power_size现在是第3个元素
                            old_start_date = old_values[3]  # start_date现在是第4个元素
                            # 查找并更新匹配的条目
                            for i, sched in enumerate(self.data_model.maintenance_schedules):
                                if (sched.get('name', '') == old_name and
                                    sched.get('power_size', 0) == old_power_size and 
                                    sched.get('start_date', '') == old_start_date):
                                    self.data_model.maintenance_schedules[i] = schedule_data
                                    break
                else:
                    # 添加新数据
                    self.data_model.maintenance_schedules.append(schedule_data)
            else:  # commissioning
                if item_id:
                    # 更新现有数据（使用UI Treeview中的索引）
                    all_items = tree.get_children()
                    current_index = all_items.index(item_id) if item_id in all_items else -1
                    
                    # 如果能确定索引，就按索引更新（假定UI和数据模型顺序一致）
                    if current_index != -1 and current_index < len(self.data_model.commissioning_schedules):
                        # 直接按索引更新，这是最可靠的方法
                        self.data_model.commissioning_schedules[current_index] = schedule_data
                    else:
                        # 如果索引无效，尝试使用名称等字段匹配
                        item = tree.item(item_id)
                        old_values = item['values']
                        if len(old_values) >= 3:
                            old_name = old_values[0]  # 计划名称是第1个元素
                            old_power_size = float(old_values[2])  # power_size现在是第3个元素
                            old_start_date = old_values[3]  # start_date现在是第4个元素
                            # 查找并更新匹配的条目
                            for i, sched in enumerate(self.data_model.commissioning_schedules):
                                if (sched.get('name', '') == old_name and
                                    sched.get('power_size', 0) == old_power_size and 
                                    sched.get('start_date', '') == old_start_date):
                                    self.data_model.commissioning_schedules[i] = schedule_data
                                    break
                else:
                    # 添加新数据
                    self.data_model.commissioning_schedules.append(schedule_data)
                
            # 关闭对话框
            dialog.destroy()
            
            # 保存项目数据
            self.save_current_project()
        
        ttk.Button(button_frame, text="确定", command=confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
    def load_maintenance_schedules(self):
        """
        加载检修和投产计划数据到表格
        """
        # 清空现有数据
        for item in self.maintenance_tree.get_children():
            self.maintenance_tree.delete(item)
            
        for item in self.commissioning_tree.get_children():
            self.commissioning_tree.delete(item)
            
        for item in self.output_limit_tree.get_children():
            self.output_limit_tree.delete(item)
        
        # 加载检修计划数据
        for schedule in self.data_model.maintenance_schedules:
            values = [
                schedule.get('name', ''),  # 添加名称字段
                schedule.get('power_type', ''),
                str(schedule.get('power_size', 0)),
                schedule.get('start_date', ''),
                schedule.get('end_date', '')
            ]
            # 如果有设备型号，则添加到values中
            if 'model' in schedule:
                values.append(schedule['model'])
            self.maintenance_tree.insert('', tk.END, values=values)
        
        # 加载投产计划数据
        for schedule in self.data_model.commissioning_schedules:
            values = [
                schedule.get('name', ''),  # 添加名称字段
                schedule.get('power_type', ''),
                str(schedule.get('power_size', 0)),
                schedule.get('start_date', ''),
                schedule.get('end_date', '')
            ]
            # 如果有设备型号，则添加到values中
            if 'model' in schedule:
                values.append(schedule['model'])
            self.commissioning_tree.insert('', tk.END, values=values)
            
        # 加载出力限制计划数据
        for schedule in self.data_model.output_limit_schedules:
            values = [
                schedule.get('name', ''),
                schedule.get('limit_type', ''),
                str(schedule.get('power_size', 0)),
                schedule.get('start_date', ''),
                schedule.get('end_date', '')
            ]
            self.output_limit_tree.insert('', tk.END, values=values)
    
    def add_output_limit_entry(self):
        """
        添加出力限制计划条目
        """
        self.manage_output_limit_entry("添加出力限制计划")
        
    def edit_output_limit_entry(self):
        """
        编辑出力限制计划条目
        """
        selection = self.output_limit_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个出力限制计划条目！")
            return
        self.manage_output_limit_entry("编辑出力限制计划", selection[0])
        
    def delete_output_limit_entry(self):
        """
        删除出力限制计划条目
        """
        selection = self.output_limit_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择一个出力限制计划条目！")
            return
            
        if messagebox.askyesno("确认删除", "确定要删除选中的出力限制计划条目吗？"):
            # 使用更可靠的方法：首先获取UI树中的项目索引，然后从数据模型中删除相同索引的项目
            all_items = self.output_limit_tree.get_children()
            current_index = all_items.index(selection[0]) if selection[0] in all_items else -1
            
            # 如果UI和数据模型保持同步，我们可以按索引删除
            if current_index != -1 and current_index < len(self.data_model.output_limit_schedules):
                del self.data_model.output_limit_schedules[current_index]
            else:
                # 如果索引无效，使用原来的方法
                # 获取要删除的条目信息
                item = self.output_limit_tree.item(selection[0])
                values = item['values']
                
                # 从数据模型中删除对应条目
                if len(values) >= 3:  # 确保有足够的数据 (name, limit_type, power_size)
                    name = values[0]  # 计划名称是第1个元素
                    power_size = float(values[2])  # power_size是第3个元素（索引为2）
                    start_date = values[3]  # 开始日期是第4个元素
                    # 查找并删除匹配的条目
                    for i, sched in enumerate(self.data_model.output_limit_schedules):
                        if (sched.get('name', '') == name and
                            sched.get('power_size', 0) == power_size and 
                            sched.get('start_date', '') == start_date):
                            del self.data_model.output_limit_schedules[i]
                            break
            
            # 从UI中删除条目
            self.output_limit_tree.delete(selection[0])
            
            # 保存项目数据
            self.save_current_project()
    
    def manage_output_limit_entry(self, title, item_id=None):
        """
        管理出力限制计划条目
        """
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 获取当前值（如果是编辑模式）
        current_values = []
        if item_id:
            current_values = self.output_limit_tree.item(item_id, 'values')
        
        # 计划名称
        ttk.Label(dialog, text="计划名称:").grid(row=0, column=0, sticky=tk.W, pady=5)
        name_var = tk.StringVar(value=current_values[0] if current_values and len(current_values) > 0 else "")
        ttk.Entry(dialog, textvariable=name_var, width=20).grid(row=0, column=1, pady=5, padx=5, sticky=tk.W)
        
        # 限制类型
        ttk.Label(dialog, text="限制类型:").grid(row=1, column=0, sticky=tk.W, pady=5)
        limit_type_var = tk.StringVar(value=current_values[1] if current_values and len(current_values) > 1 else "")
        limit_type_combo = ttk.Combobox(dialog, textvariable=limit_type_var, state="readonly")
        limit_type_combo['values'] = ("光伏最大出力限制", "风机最大出力限制")
        limit_type_combo.grid(row=1, column=1, pady=5, padx=5, sticky=tk.W)
        
        # 限制出力大小
        ttk.Label(dialog, text="限制出力大小(kW): ").grid(row=2, column=0, sticky=tk.W, pady=5)
        power_size_var = tk.DoubleVar(value=float(current_values[2]) if current_values and len(current_values) > 2 else 0.0)
        ttk.Entry(dialog, textvariable=power_size_var, width=20).grid(row=2, column=1, pady=5, padx=5, sticky=tk.W)
        
        # 开始日期
        ttk.Label(dialog, text="开始日期: ").grid(row=3, column=0, sticky=tk.W, pady=5)
        start_date_var = tk.StringVar(value=current_values[3] if current_values and len(current_values) > 3 else "")
        ttk.Entry(dialog, textvariable=start_date_var, width=20).grid(row=3, column=1, pady=5, padx=5, sticky=tk.W)
        ttk.Label(dialog, text="格式: YYYY-MM-DD", foreground="gray").grid(row=4, column=1, sticky=tk.W, pady=0)
        
        # 结束日期
        ttk.Label(dialog, text="结束日期: ").grid(row=5, column=0, sticky=tk.W, pady=5)
        end_date_var = tk.StringVar(value=current_values[4] if current_values and len(current_values) > 4 else "")
        ttk.Entry(dialog, textvariable=end_date_var, width=20).grid(row=5, column=1, pady=5, padx=5, sticky=tk.W)
        ttk.Label(dialog, text="格式: YYYY-MM-DD", foreground="gray").grid(row=6, column=1, sticky=tk.W, pady=0)
        
        # 确定和取消按钮
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=7, column=0, columnspan=2, pady=20)
        
        def confirm():
            # 获取输入值
            name = name_var.get().strip()
            limit_type = limit_type_var.get()
            power_size = power_size_var.get()
            start_date = start_date_var.get()
            end_date = end_date_var.get()
            
            # 验证输入
            if not name:
                messagebox.showwarning("警告", "请输入计划名称！")
                return
                
            if not limit_type:
                messagebox.showwarning("警告", "请选择限制类型！")
                return
                
            if power_size <= 0:
                messagebox.showwarning("警告", "限制出力大小必须大于0！")
                return
                
            if not start_date or not end_date:
                messagebox.showwarning("警告", "请输入开始日期和结束日期！")
                return
                
            # 验证日期格式
            try:
                datetime.strptime(start_date, "%Y-%m-%d")
                datetime.strptime(end_date, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("错误", "日期格式不正确！请使用 YYYY-MM-DD 格式。")
                return
                
            # 构造条目值
            values = [name, limit_type, str(power_size), start_date, end_date]
                
            # 添加或更新条目
            if item_id:
                # 更新现有条目
                self.output_limit_tree.item(item_id, values=values)
            else:
                # 添加新条目
                self.output_limit_tree.insert('', tk.END, values=values)
                
            # 同步更新数据模型
            schedule_data = {
                'name': name,
                'limit_type': limit_type,
                'power_size': power_size,
                'start_date': start_date,
                'end_date': end_date
            }
                
            if item_id:
                # 更新现有数据（需要找到对应的索引）
                item = self.output_limit_tree.item(item_id)
                old_values = item['values']
                if len(old_values) >= 3:
                    old_name = old_values[0]  # 计划名称是第1个元素
                    old_power_size = float(old_values[2])  # power_size是第3个元素
                    old_start_date = old_values[3]  # start_date是第4个元素
                    # 查找并更新匹配的条目
                    found = False
                    for i, sched in enumerate(self.data_model.output_limit_schedules):
                        if (sched.get('name', '') == old_name and
                            sched.get('power_size', 0) == old_power_size and 
                            sched.get('start_date', '') == old_start_date):
                            self.data_model.output_limit_schedules[i] = schedule_data
                            found = True
                            break
                    # 如果没找到匹配项，则添加新项
                    if not found:
                        self.data_model.output_limit_schedules.append(schedule_data)
            else:
                # 添加新数据
                self.data_model.output_limit_schedules.append(schedule_data)
                
            # 关闭对话框
            dialog.destroy()
            
            # 保存项目数据
            self.save_current_project()
        
        ttk.Button(button_frame, text="确定", command=confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def date_to_hour(self, date):
        """
        将日期转换为一年中对应的小时数
        假设数据是从2025年1月1日0时开始的
        """
        start_of_year = datetime(2025, 1, 1)
        delta = date - start_of_year
        return int(delta.total_seconds() // 3600)
    
    def on_legend_click_data(self, event):
        """
        处理数据图表图例点击事件，显示/隐藏对应的曲线
        """
        # 获取被点击的图例元素
        legitem = event.artist
        
        # 获取对应的原始线条
        origline = self.lined_data[legitem]
        
        # 切换线条的可见性
        vis = not origline.get_visible()
        origline.set_visible(vis)
        
        # 更新所有相关的图例元素的透明度
        legend = self.data_ax.get_legend()
        if legend:
            for legline, origline_ref in self.lined_data.items():
                if origline_ref == origline:
                    if vis:
                        legline.set_alpha(1.0)
                    else:
                        legline.set_alpha(0.2)
        
        # 自适应调整y轴范围
        self.auto_adjust_y_axis_data()
        
        # 刷新画布
        self.data_canvas.draw()

    def on_legend_click_result(self, event):
        """
        处理结果图表图例点击事件，显示/隐藏对应的曲线
        """
        # 获取被点击的图例元素
        legitem = event.artist
        
        # 获取对应的原始线条或填充
        origline = self.lined_result[legitem]
        
        # 切换线条或填充的可见性
        vis = not origline.get_visible()
        origline.set_visible(vis)
        
        # 更新所有相关的图例元素的透明度
        legend = self.ax.get_legend()
        if legend:
            for legline, origline_ref in self.lined_result.items():
                if origline_ref == origline:
                    if vis:
                        legline.set_alpha(1.0)
                    else:
                        legline.set_alpha(0.2)
        
        # 自适应调整y轴范围
        self.auto_adjust_y_axis_result()
        
        # 刷新画布
        self.canvas.draw()

    def on_data_hover(self, event):
        """
        处理数据图表上的鼠标悬浮事件，显示当前悬浮位置的时间和所有曲线的实际数据值
        """
        # 检查事件是否在图表区域内
        if event.inaxes != self.data_ax:
            # 如果鼠标不在图表区域内，移除注释
            if hasattr(self, 'data_annotation') and self.data_annotation:
                try:
                    self.data_annotation.remove()
                    self.data_annotation = None
                    self.data_canvas.draw_idle()
                except:
                    pass
            return
            
        # 获取当前坐标
        x, y = event.xdata, event.ydata
        
        # 将x坐标转换为日期 - 使用matplotlib的日期转换机制
        try:
            import matplotlib.dates as mdates
            from datetime import datetime, timedelta
            
            # 将matplotlib日期转换为datetime对象
            hover_datetime = mdates.num2date(x)
            # 只显示到整小时
            date_str = hover_datetime.strftime('%m-%d %H:00')  # 只显示月-日 小时:00，不显示分钟
            
            # 计算最接近的小时索引，基于相对于年初的小时数
            # 假设数据从1月1日开始，计算相对于年初的小时数
            # 获取hover_datetime是一年中的第几天和小时
            day_of_year = hover_datetime.timetuple().tm_yday  # 一年中的第几天 (1-366)
            hour_of_day = hover_datetime.hour  # 小时 (0-23)
            
            # 计算总小时数 (0-8759)
            hour_idx = (day_of_year - 1) * 24 + hour_of_day
            
            # 确保小时索引在有效范围内
            if 0 <= hour_idx < 8760:
                # 获取所有曲线在当前小时的数据（使用实际数据值，而非鼠标位置的y值）
                values_info = []
                
                # 检查是否有对应的数据
                if self.data_model.data_imported['electric'] and hour_idx < len(self.data_model.electric_load_hourly):
                    values_info.append(f"电力负荷: {self.data_model.electric_load_hourly[hour_idx]:.2f} kW")
                
                if self.data_model.data_imported['heat'] and hour_idx < len(self.data_model.heat_load_hourly):
                    values_info.append(f"热力负荷: {self.data_model.heat_load_hourly[hour_idx]:.2f} kW")
                
                if self.data_model.data_imported['solar'] and hour_idx < len(self.data_model.solar_irradiance_hourly):
                    values_info.append(f"光照强度: {self.data_model.solar_irradiance_hourly[hour_idx]:.2f} W/m²")
                
                if self.data_model.data_imported['wind'] and hour_idx < len(self.data_model.wind_speed_hourly):
                    values_info.append(f"风速: {self.data_model.wind_speed_hourly[hour_idx]:.2f} m/s")
                
                if self.data_model.data_imported['grid_price'] and hour_idx < len(self.data_model.grid_purchase_price_hourly):
                    values_info.append(f"下网电价: {self.data_model.grid_purchase_price_hourly[hour_idx]:.2f} 元/kWh")
                
                if values_info:
                    # 组合信息，只使用实际数据值，不使用鼠标位置的y值
                    tooltip_text = f"日期: {date_str}\n" + "\n".join(values_info)
                else:
                    # 如果没有可用数据，显示日期和坐标值
                    tooltip_text = f"日期: {date_str}\n数值: {y:.2f}"
            else:
                # 如果超出范围，只显示日期
                tooltip_text = f"日期: {date_str}\n超出数据范围(索引: {hour_idx})"
            
            # 清除之前的注释
            if hasattr(self, 'data_annotation') and self.data_annotation:
                try:
                    self.data_annotation.remove()
                except:
                    pass
            
            # 创建新的注释，使用x坐标位置但不依赖y坐标值
            self.data_annotation = self.data_ax.annotate(
                tooltip_text,
                xy=(x, 0),  # 只使用x坐标来定位垂直位置，y值用0作为参考
                xytext=(10, 10),
                textcoords='offset points',
                bbox=dict(boxstyle='round,pad=0.3', fc='yellow', alpha=0.7),
                fontsize=9
            )
            
            # 刷新画布
            self.data_canvas.draw_idle()
            
        except Exception as e:
            # 如果转换出错，移除可能存在的注释
            if hasattr(self, 'data_annotation') and self.data_annotation:
                try:
                    self.data_annotation.remove()
                    self.data_annotation = None
                except:
                    pass

    def on_result_hover(self, event):
        """
        处理结果图表上的鼠标悬浮事件，显示当前悬浮位置的时间和所有曲线的实际数据值
        """
        # 检查事件是否在图表区域内
        if event.inaxes != self.ax:
            # 如果鼠标不在图表区域内，移除注释
            if hasattr(self, 'result_annotation') and self.result_annotation:
                try:
                    self.result_annotation.remove()
                    self.result_annotation = None
                    self.canvas.draw_idle()
                except:
                    pass
            return
            
        # 获取当前坐标
        x, y = event.xdata, event.ydata
        
        # 将x坐标转换为日期 - 使用matplotlib的日期转换机制
        try:
            import matplotlib.dates as mdates
            from datetime import datetime, timedelta
            
            # 将matplotlib日期转换为datetime对象
            hover_datetime = mdates.num2date(x)
            # 只显示到整小时
            date_str = hover_datetime.strftime('%m-%d %H:00')  # 只显示月-日 小时:00，不显示分钟
            
            # 计算最接近的小时索引，基于相对于年初的小时数
            # 获取hover_datetime是一年中的第几天和小时
            day_of_year = hover_datetime.timetuple().tm_yday  # 一年中的第几天 (1-366)
            hour_of_day = hover_datetime.hour  # 小时 (0-23)
            
            # 计算总小时数 (0-8759)
            hour_idx = (day_of_year - 1) * 24 + hour_of_day
            
            # 确保小时索引在有效范围内
            if 0 <= hour_idx < 8760 and self.results:
                # 获取所有曲线在当前小时的数据（使用实际数据值，而非鼠标位置的y值）
                values_info = []
                
                # 检查结果数据是否可用
                if hour_idx < len(self.results.get('hourly_total_load', [])):
                    values_info.append(f"总负荷: {self.results['hourly_total_load'][hour_idx]:.2f} kW")
                
                if hour_idx < len(self.results.get('hourly_generation', [])):
                    values_info.append(f"总出力: {self.results['hourly_generation'][hour_idx]:.2f} kW")
                
                if hour_idx < len(self.results.get('hourly_grid_load', [])):
                    values_info.append(f"下网负荷: {self.results['hourly_grid_load'][hour_idx]:.2f} kW")
                
                if hour_idx < len(self.results.get('hourly_pv_output', [])):
                    values_info.append(f"光伏出力: {self.results['hourly_pv_output'][hour_idx]:.2f} kW")
                
                if hour_idx < len(self.results.get('hourly_wind_output', [])):
                    values_info.append(f"风机出力: {self.results['hourly_wind_output'][hour_idx]:.2f} kW")
                
                if hour_idx < len(self.results.get('hourly_chp_output', [])):
                    values_info.append(f"热电出力: {self.results['hourly_chp_output'][hour_idx]:.2f} kW")
                
                if values_info:
                    # 组合信息，只使用实际数据值，不使用鼠标位置的y值
                    tooltip_text = f"日期: {date_str}\n" + "\n".join(values_info)
                else:
                    # 如果没有可用数据，显示日期和坐标值
                    tooltip_text = f"日期: {date_str}\n数值: {y:.2f}"
            else:
                # 如果超出范围，只显示日期
                tooltip_text = f"日期: {date_str}\n超出数据范围(索引: {hour_idx})"
            
            # 清除之前的注释
            if hasattr(self, 'result_annotation') and self.result_annotation:
                try:
                    self.result_annotation.remove()
                except:
                    pass
            
            # 创建新的注释，使用x坐标位置但不依赖y坐标值
            self.result_annotation = self.ax.annotate(
                tooltip_text,
                xy=(x, 0),  # 只使用x坐标来定位垂直位置，y值用0作为参考
                xytext=(10, 10),
                textcoords='offset points',
                bbox=dict(boxstyle='round,pad=0.3', fc='yellow', alpha=0.7),
                fontsize=9
            )
            
            # 刷新画布
            self.canvas.draw_idle()
            
        except Exception as e:
            # 如果转换出错，移除可能存在的注释
            if hasattr(self, 'result_annotation') and self.result_annotation:
                try:
                    self.result_annotation.remove()
                    self.result_annotation = None
                except:
                    pass

    def on_optimization_hover(self, event):
        """
        处理优化结果图表上的鼠标悬浮事件，显示当前悬浮位置的时间和所有曲线的实际数据值
        """
        # 检查事件是否在图表区域内
        if event.inaxes != self.optimization_ax:
            # 如果鼠标不在图表区域内，移除注释
            if hasattr(self, 'optimization_annotation') and self.optimization_annotation:
                try:
                    self.optimization_annotation.remove()
                    self.optimization_annotation = None
                    self.optimization_canvas.draw_idle()
                except:
                    pass
            return
            
        # 获取当前坐标
        x, y = event.xdata, event.ydata
        
        # 将x坐标转换为日期 - 使用matplotlib的日期转换机制
        try:
            import matplotlib.dates as mdates
            from datetime import datetime, timedelta
            
            # 将matplotlib日期转换为datetime对象
            hover_datetime = mdates.num2date(x)
            # 只显示到整小时
            date_str = hover_datetime.strftime('%m-%d %H:00')  # 只显示月-日 小时:00，不显示分钟
            
            # 计算最接近的小时索引，基于相对于年初的小时数
            # 获取hover_datetime是一年中的第几天和小时
            day_of_year = hover_datetime.timetuple().tm_yday  # 一年中的第几天 (1-366)
            hour_of_day = hover_datetime.hour  # 小时 (0-23)
            
            # 计算总小时数 (0-8759)
            hour_idx = (day_of_year - 1) * 24 + hour_of_day
            
            # 确保小时索引在有效范围内
            if 0 <= hour_idx < 8760:
                # 获取所有曲线在当前小时的数据（使用实际数据值，而非鼠标位置的y值）
                values_info = []
                
                # 检查优化结果数据是否可用
                if hasattr(self, 'optimized_results') and self.optimized_results:
                    if hour_idx < len(self.optimized_results.get('hourly_basic_load', [])):
                        values_info.append(f"基础负荷(优化后): {self.optimized_results['hourly_basic_load'][hour_idx]:.2f} kW")
                    
                    if hour_idx < len(self.optimized_results.get('hourly_flexible_load', [])):
                        values_info.append(f"灵活负荷(优化后): {self.optimized_results['hourly_flexible_load'][hour_idx]:.2f} kW")
                    
                    if hour_idx < len(self.optimized_results.get('hourly_revenue', [])):
                        values_info.append(f"每小时收益: {self.optimized_results['hourly_revenue'][hour_idx]:.2f} 元")
                
                # 如果有平衡计算结果，也显示优化前的数据
                if self.results:
                    if hour_idx < len(self.results.get('hourly_corrected_electric_load', [])):
                        values_info.append(f"修正后电力负荷(优化前): {self.results['hourly_corrected_electric_load'][hour_idx]:.2f} kW")
                    
                    if hour_idx < len(self.results.get('hourly_grid_load', [])):
                        values_info.append(f"下网负荷: {self.results['hourly_grid_load'][hour_idx]:.2f} kW")
                
                if values_info:
                    # 组合信息，只使用实际数据值，不使用鼠标位置的y值
                    tooltip_text = f"日期: {date_str}\n" + "\n".join(values_info)
                else:
                    # 如果没有可用数据，显示日期和坐标值
                    tooltip_text = f"日期: {date_str}\n数值: {y:.2f}"
            else:
                # 如果超出范围，只显示日期
                tooltip_text = f"日期: {date_str}\n超出数据范围(索引: {hour_idx})"
            
            # 清除之前的注释
            if hasattr(self, 'optimization_annotation') and self.optimization_annotation:
                try:
                    self.optimization_annotation.remove()
                except:
                    pass
            
            # 创建新的注释，使用x坐标位置但不依赖y坐标值
            self.optimization_annotation = self.optimization_ax.annotate(
                tooltip_text,
                xy=(x, 0),  # 只使用x坐标来定位垂直位置，y值用0作为参考
                xytext=(10, 10),
                textcoords='offset points',
                bbox=dict(boxstyle='round,pad=0.3', fc='yellow', alpha=0.7),
                fontsize=9
            )
            
            # 刷新画布
            self.optimization_canvas.draw_idle()
            
        except Exception as e:
            # 如果转换出错，移除可能存在的注释
            if hasattr(self, 'optimization_annotation') and self.optimization_annotation:
                try:
                    self.optimization_annotation.remove()
                    self.optimization_annotation = None
                except:
                    pass

    def auto_adjust_y_axis_data(self):
        """
        自动调整数据图表的y轴范围，基于当前可见的线条
        """
        # 找到所有可见的线条
        visible_lines = []
        for line in self.data_ax.get_lines():
            if line.get_visible():
                visible_lines.append(line)
        
        if not visible_lines:
            # 如果没有可见线条，设置默认范围
            self.data_ax.set_ylim(0, 1)
            return
        
        # 计算所有可见线条的y值范围
        all_y_values = []
        for line in visible_lines:
            y_data = line.get_ydata()
            all_y_values.extend(y_data)
        
        if all_y_values:
            y_min = min(all_y_values)
            y_max = max(all_y_values)
            
            # 添加一些边距
            margin = (y_max - y_min) * 0.05 if y_max != y_min else 0.1
            y_min -= margin
            y_max += margin
            
            self.data_ax.set_ylim(y_min, y_max)

    def auto_adjust_y_axis_result(self):
        """
        自动调整结果图表的y轴范围，基于当前可见的线条
        """
        # 找到所有可见的线条和填充区域
        visible_lines = []
        for line in self.ax.get_lines():
            if line.get_visible():
                visible_lines.append(line)
        
        # 也检查填充区域（fill_between的结果）
        for patch in self.ax.collections:  # fill_between创建的是collections
            if hasattr(patch, 'get_visible') and patch.get_visible():
                visible_lines.append(patch)
        
        if not visible_lines:
            # 如果没有可见线条，设置默认范围
            self.ax.set_ylim(0, 1)
            return
        
        # 计算所有可见线条的y值范围
        all_y_values = []
        for line in visible_lines:
            if hasattr(line, 'get_ydata'):  # 普通线条
                y_data = line.get_ydata()
                all_y_values.extend(y_data)
            elif hasattr(line, 'get_paths'):  # 填充区域
                paths = line.get_paths()
                for path in paths:
                    vertices = path.vertices
                    y_vals = [v[1] for v in vertices]
                    all_y_values.extend(y_vals)
        
        if all_y_values:
            y_min = min(all_y_values)
            y_max = max(all_y_values)
            
            # 添加一些边距
            margin = (y_max - y_min) * 0.05 if y_max != y_min else 0.1
            y_min -= margin
            y_max += margin
            
            self.ax.set_ylim(y_min, y_max)
    
    def on_mouse_wheel_data(self, event):
        """
        处理数据区域的鼠标滚轮事件
        """
        if event.delta > 0:
            self.data_canvas.yview_scroll(-1, "units")
        else:
            self.data_canvas.yview_scroll(1, "units")

def main():
    root = tk.Tk()
    app = EnergyBalanceApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()