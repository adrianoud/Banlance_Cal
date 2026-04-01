"""
UI 组件模块
包含所有 UI 相关的组件和布局方法
将 UI 代码从 loadcalculation.py 分离出来以提高可维护性
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import json
import os


class UIManager:
    """
    UI 管理类
    负责所有 UI 组件的创建和管理
    """
    
    def __init__(self, app):
        """
        初始化 UI 管理器
        :param app: EnergyBalanceApp 实例
        """
        self.app = app
        self.root = app.root
    
    def create_all_tabs(self, notebook):
        """
        创建所有标签页
        :param notebook: ttk.Notebook 实例
        """
        # 数据导入标签页（使用 UI 管理器的完整实现）
        self.create_data_import_tab(notebook)
        
        # 机组设置标签页（调用 app 中的原始完整方法）
        self.app.create_function_settings_tab(notebook)
        
        # 检修和投产计划标签页（调用 app 中的原始完整方法）
        self.app.create_maintenance_schedule_tab(notebook)
        
        # 计算与结果标签页（调用 app 中的原始完整方法）
        self.app.create_calculation_tab(notebook)
        
        # 成本分析 2.0 标签页（使用 UI 管理器的完整实现）
        self.create_cost_analysis_v2_tab(notebook)
    
    def create_data_import_tab(self, notebook):
        """创建数据导入标签页（完整版）"""
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="📊 数据导入")
        
        # 添加返回项目列表按钮
        back_btn = ttk.Button(tab, text="保存并返回项目列表", 
                             command=self.app.save_and_return_to_project_list)
        back_btn.grid(row=0, column=3, sticky=tk.E, padx=5, pady=5)
        
        # 数据说明
        info_label = ttk.Label(
            tab, 
            text="请导入包含 8760 小时数据的 CSV 文件\n"
                 "文件应包含列：时间，电力负荷 (kW), 热力负荷 (kW), 光照强度 (W/m²), 风速 (m/s)"
        )
        info_label.grid(row=1, column=0, columnspan=4, pady=(0, 20), sticky=tk.W)
        
        # 添加下载模板按钮
        ttk.Button(tab, text="下载 CSV 模板", 
                  command=self.app.download_template).grid(row=2, column=0, pady=5, sticky=tk.W)
        
        # 单一文件导入控件
        ttk.Label(tab, text="统一数据文件:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.app.single_file_path = tk.StringVar()
        # 初始化单文件模式变量
        self.app.single_file_mode = tk.BooleanVar(value=True)  # 默认使用单文件模式
        self.app.single_file_entry = ttk.Entry(tab, textvariable=self.app.single_file_path, width=50)
        self.app.single_file_entry.grid(row=3, column=1, padx=5, pady=5)
        self.app.single_file_button = ttk.Button(
            tab, 
            text="浏览...", 
            command=lambda: self.app.browse_file(self.app.single_file_path)
        )
        self.app.single_file_button.grid(row=3, column=2, pady=5)
        
        # 厂用电率设置
        ttk.Label(tab, text="厂用电率:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.app.internal_rate_var = tk.DoubleVar(value=0.05)
        ttk.Entry(tab, textvariable=self.app.internal_rate_var, width=20).grid(row=4, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(tab, text="(小数形式，如 0.05 表示 5%)").grid(row=4, column=1, sticky=tk.E, padx=5, pady=5)
        
        # 导入按钮
        ttk.Button(tab, text="导入数据", command=self.app.import_all_data).grid(row=5, column=0, columnspan=3, pady=20)
        
        # 刷新图表按钮
        ttk.Button(tab, text="刷新趋势图", command=self.app.update_imported_data_plot).grid(row=5, column=2, pady=20, sticky=tk.E)
        
        # 数据统计信息框架
        stats_frame = ttk.LabelFrame(tab, text="数据统计", padding="10")
        stats_frame.grid(row=6, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # 创建统计文本框
        self.app.stats_text = tk.Text(stats_frame, height=15, width=80)
        stats_scrollbar = ttk.Scrollbar(stats_frame, orient=tk.VERTICAL, command=self.app.stats_text.yview)
        self.app.stats_text.configure(yscrollcommand=stats_scrollbar.set)
        
        self.app.stats_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        stats_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 时间段选择区域
        time_range_frame = ttk.LabelFrame(tab, text="时间段选择", padding="10")
        time_range_frame.grid(row=7, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(time_range_frame, text="开始日期:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.app.start_date_var = tk.StringVar(value="2025-01-01")
        self.app.start_date_entry = ttk.Entry(time_range_frame, textvariable=self.app.start_date_var, width=12)
        self.app.start_date_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(time_range_frame, text="结束日期:").grid(row=0, column=2, sticky=tk.W, padx=(10, 5))
        self.app.end_date_var = tk.StringVar(value="2025-12-31")
        self.app.end_date_entry = ttk.Entry(time_range_frame, textvariable=self.app.end_date_var, width=12)
        self.app.end_date_entry.grid(row=0, column=3, padx=5)
        
        ttk.Button(time_range_frame, text="更新图表", command=self.app.update_imported_data_plot).grid(row=0, column=4, padx=(10, 0))
        
        # 图表展示（用于显示导入数据的趋势）
        plot_frame = ttk.LabelFrame(tab, text="已导入数据趋势图", padding="10")
        plot_frame.grid(row=8, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # 创建 matplotlib 图形
        self.app.data_figure = Figure(figsize=(10, 6), dpi=100)
        self.app.data_ax = self.app.data_figure.add_subplot(111)
        self.app.data_canvas = FigureCanvasTkAgg(self.app.data_figure, plot_frame)
        self.app.data_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 配置权重
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)
        time_range_frame.columnconfigure(5, weight=1)
        plot_frame.columnconfigure(0, weight=1)
        plot_frame.rowconfigure(0, weight=1)
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(8, weight=1)
    
    def create_function_settings_tab(self, notebook):
        """创建机组设置标签页"""
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="⚙️ 机组设置")
    
    def create_maintenance_schedule_tab(self, notebook):
        """创建检修和投产计划标签页"""
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="📅 检修和投产计划")
        
        # 添加返回项目列表按钮
        back_btn = ttk.Button(tab, text="保存并返回项目列表",
                             command=self.app.save_and_return_to_project_list)
        back_btn.grid(row=0, column=0, sticky=tk.E, padx=5, pady=5)
        
        # ... (其他代码保持不变，为简洁省略)
    
    def create_calculation_tab(self, notebook):
        """创建计算与结果标签页"""
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="📊 计算与结果")
        
        # 添加返回项目列表按钮
        back_btn = ttk.Button(tab, text="保存并返回项目列表",
                             command=self.app.save_and_return_to_project_list)
        back_btn.grid(row=0, column=0, sticky=tk.E, padx=5, pady=5)
        
        # ... (其他代码保持不变，为简洁省略)
    

    
    def create_cost_analysis_v2_tab(self, notebook):
        """
        创建成本分析 2.0 标签页（完整版）
        UI 布局：
        - 上方为输入区域（三大板块：火电、新能源、下网）
        - 下方为结果展示（左右布局）
        """
        tab = ttk.Frame(notebook, padding="10")
        notebook.add(tab, text="📈 成本分析 2.0")
        
        # ===== 顶部按钮区域 =====
        top_button_frame = ttk.Frame(tab)
        top_button_frame.grid(row=0, column=0, sticky=tk.E, padx=5, pady=5)
        
        # 添加导出汇总结果按钮
        export_btn = ttk.Button(top_button_frame, text="📊 导出汇总结果", 
                               command=self.app.export_v2_cost_summary)
        export_btn.pack(side=tk.RIGHT, padx=5)
        
        refresh_btn = ttk.Button(top_button_frame, text="🔄 刷新成本数据", 
                                command=self.app.update_v2_cost_analysis)
        refresh_btn.pack(side=tk.RIGHT, padx=5)
        
        save_btn = ttk.Button(top_button_frame, text="💾 保存成本参数", 
                             command=self.app.save_v2_cost_parameters)
        save_btn.pack(side=tk.RIGHT, padx=5)
        
        back_btn = ttk.Button(top_button_frame, text="保存并返回项目列表",
                             command=self.app.save_and_return_to_project_list)
        back_btn.pack(side=tk.RIGHT, padx=5)
        
        # ===== 上方输入区域 =====
        top_input_frame = ttk.Frame(tab)
        top_input_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # --- 第 1 列：火电成本（分 2 列显示）---
        thermal_frame = ttk.LabelFrame(top_input_frame, text="1. 火电成本", padding=5)
        thermal_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        thermal_left_col = ttk.Frame(thermal_frame)
        thermal_left_col.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 2))
        
        thermal_right_col = ttk.Frame(thermal_frame)
        thermal_right_col.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(2, 0))
        
        # === 左列：直接材料、直接人工、制造费用 ===
        row_idx = 0
        
        # 1.1 直接材料成本
        material_frame = ttk.LabelFrame(thermal_left_col, text="直接材料成本", padding=5)
        material_frame.grid(row=row_idx, column=0, sticky=(tk.W, tk.E), pady=3)
        row_idx += 1
        
        ttk.Label(material_frame, text="原煤单价 (元/吨):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_coal_price_var = tk.DoubleVar(value=162.4)
        ttk.Entry(material_frame, textvariable=self.app.v2_thermal_coal_price_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(material_frame, text="原煤基准单耗 (g/kWh):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_base_coal_var = tk.DoubleVar(value=500.0)
        ttk.Entry(material_frame, textvariable=self.app.v2_thermal_base_coal_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        ttk.Label(material_frame, text="纯水及其他费用 (万元):").grid(row=2, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_pure_water_cost_var = tk.DoubleVar(value=100.0)
        ttk.Entry(material_frame, textvariable=self.app.v2_thermal_pure_water_cost_var, width=12).grid(row=2, column=1, pady=2, padx=3)
        
        ttk.Label(material_frame, text="纯水及其他单价 (元/kWh):").grid(row=3, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_pure_water_unit_var = tk.DoubleVar(value=0.05)
        ttk.Entry(material_frame, textvariable=self.app.v2_thermal_pure_water_unit_var, width=12).grid(row=3, column=1, pady=2, padx=3)
        
        # 1.2 直接人工成本
        labor_frame = ttk.LabelFrame(thermal_left_col, text="直接人工成本", padding=5)
        labor_frame.grid(row=row_idx, column=0, sticky=(tk.W, tk.E), pady=3)
        row_idx += 1
        
        ttk.Label(labor_frame, text="直接人工 (万元):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_direct_labor_var = tk.DoubleVar(value=4000.0)
        ttk.Entry(labor_frame, textvariable=self.app.v2_thermal_direct_labor_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        # 1.3 制造费用成本
        manufacturing_frame = ttk.LabelFrame(thermal_left_col, text="制造费用成本", padding=5)
        manufacturing_frame.grid(row=row_idx, column=0, sticky=(tk.W, tk.E), pady=3)
        row_idx += 1
        
        ttk.Label(manufacturing_frame, text="管理人工成本 (万元):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_mgmt_labor_var = tk.DoubleVar(value=1000.0)
        ttk.Entry(manufacturing_frame, textvariable=self.app.v2_thermal_mgmt_labor_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(manufacturing_frame, text="运维费用 (万元):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_maintenance_var = tk.DoubleVar(value=1000.0)
        ttk.Entry(manufacturing_frame, textvariable=self.app.v2_thermal_maintenance_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        ttk.Label(manufacturing_frame, text="折旧及摊销费用 (万元):").grid(row=2, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_depreciation_var = tk.DoubleVar(value=5000.0)
        ttk.Entry(manufacturing_frame, textvariable=self.app.v2_thermal_depreciation_var, width=12).grid(row=2, column=1, pady=2, padx=3)
        
        ttk.Label(manufacturing_frame, text="其他制造费用 (万元):").grid(row=3, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_other_manufacturing_var = tk.DoubleVar(value=500.0)
        ttk.Entry(manufacturing_frame, textvariable=self.app.v2_thermal_other_manufacturing_var, width=12).grid(row=3, column=1, pady=2, padx=3)
        
        # === 右列：备容与政府基金、碳排放和绿证、其他成本 ===
        right_row_idx = 0
        
        # 1.4 备容与政府基金
        reserve_fee_frame = ttk.LabelFrame(thermal_right_col, text="备容与政府基金", padding=5)
        reserve_fee_frame.grid(row=right_row_idx, column=0, sticky=(tk.W, tk.E), pady=3)
        right_row_idx += 1
        
        self.app.v2_thermal_reserve_fee_mode_var = tk.IntVar(value=0)
        
        unit_price_frame = ttk.Frame(reserve_fee_frame)
        unit_price_frame.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=2)
        ttk.Radiobutton(unit_price_frame, text="单价计费", variable=self.app.v2_thermal_reserve_fee_mode_var, value=0).pack(side=tk.LEFT)
        ttk.Label(unit_price_frame, text="备容费单价 (元/kWh):").pack(side=tk.LEFT, padx=(5, 2))
        self.app.v2_thermal_reserve_fee_unit_var = tk.DoubleVar(value=0.05)
        ttk.Entry(unit_price_frame, textvariable=self.app.v2_thermal_reserve_fee_unit_var, width=8).pack(side=tk.LEFT)
        
        total_price_frame = ttk.Frame(reserve_fee_frame)
        total_price_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)
        ttk.Radiobutton(total_price_frame, text="总价计费", variable=self.app.v2_thermal_reserve_fee_mode_var, value=1).pack(side=tk.LEFT)
        ttk.Label(total_price_frame, text="备容费总额 (万元):").pack(side=tk.LEFT, padx=(5, 2))
        self.app.v2_thermal_reserve_fee_total_var = tk.DoubleVar(value=0.0)
        ttk.Entry(total_price_frame, textvariable=self.app.v2_thermal_reserve_fee_total_var, width=8).pack(side=tk.LEFT)
        
        ttk.Label(reserve_fee_frame, text="政府性基金 (元/kWh):").grid(row=2, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_government_fund_var = tk.DoubleVar(value=0.0241)
        ttk.Entry(reserve_fee_frame, textvariable=self.app.v2_thermal_government_fund_var, width=12).grid(row=2, column=1, pady=2, padx=3)
        
        ttk.Label(reserve_fee_frame, text="政策性交叉补贴 (元/kWh):").grid(row=3, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_policy_subsidy_var = tk.DoubleVar(value=0.0129)
        ttk.Entry(reserve_fee_frame, textvariable=self.app.v2_thermal_policy_subsidy_var, width=12).grid(row=3, column=1, pady=2, padx=3)
        
        # 1.5 碳排放和绿证
        carbon_frame = ttk.LabelFrame(thermal_right_col, text="碳排放和绿证", padding=5)
        carbon_frame.grid(row=right_row_idx, column=0, sticky=(tk.W, tk.E), pady=3)
        right_row_idx += 1
        
        ttk.Label(carbon_frame, text="碳排放强度基准值:").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_carbon_intensity_var = tk.DoubleVar(value=0.8049)
        ttk.Entry(carbon_frame, textvariable=self.app.v2_thermal_carbon_intensity_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(carbon_frame, text="入炉煤发热量 (MJ/kg):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_heat_value_var = tk.DoubleVar(value=19.5)
        ttk.Entry(carbon_frame, textvariable=self.app.v2_thermal_heat_value_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        ttk.Label(carbon_frame, text="单位热值含碳量 (kg/MJ):").grid(row=2, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_carbon_content_var = tk.DoubleVar(value=0.0267)
        ttk.Entry(carbon_frame, textvariable=self.app.v2_thermal_carbon_content_var, width=12).grid(row=2, column=1, pady=2, padx=3)
        
        ttk.Label(carbon_frame, text="碳排放配额单价 (元/吨):").grid(row=3, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_carbon_price_var = tk.DoubleVar(value=80.0)
        ttk.Entry(carbon_frame, textvariable=self.app.v2_thermal_carbon_price_var, width=12).grid(row=3, column=1, pady=2, padx=3)
        
        ttk.Label(carbon_frame, text="可再生能源占比要求:").grid(row=4, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_green_ratio_var = tk.DoubleVar(value=0.3)
        ttk.Entry(carbon_frame, textvariable=self.app.v2_thermal_green_ratio_var, width=12).grid(row=4, column=1, pady=2, padx=3)
        
        # 1.6 其他成本
        other_cost_frame = ttk.LabelFrame(thermal_right_col, text="其他成本", padding=5)
        other_cost_frame.grid(row=right_row_idx, column=0, sticky=(tk.W, tk.E), pady=3)
        right_row_idx += 1
        
        ttk.Label(other_cost_frame, text="其他固定成本 (万元):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_other_fixed_var = tk.DoubleVar(value=0.0)
        ttk.Entry(other_cost_frame, textvariable=self.app.v2_thermal_other_fixed_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(other_cost_frame, text="其他可变成本 (元/kWh):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_thermal_other_variable_var = tk.DoubleVar(value=0.0)
        ttk.Entry(other_cost_frame, textvariable=self.app.v2_thermal_other_variable_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        # 1.7 火电煤耗变化曲线
        curve_btn_frame = ttk.Frame(thermal_right_col)
        curve_btn_frame.grid(row=right_row_idx, column=0, sticky=(tk.W, tk.E), pady=3)
        right_row_idx += 1
        
        ttk.Button(curve_btn_frame, text="🔧 设置火电煤耗曲线", command=self.app.open_v2_thermal_curve_dialog).pack(pady=5)
        
        thermal_left_col.columnconfigure(0, weight=1)
        thermal_right_col.columnconfigure(0, weight=1)
        
        # --- 第 2 列：新能源成本（光伏、风电）---
        new_energy_frame = ttk.LabelFrame(top_input_frame, text="2. 新能源成本", padding=5)
        new_energy_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 5))
        
        pv_col = ttk.Frame(new_energy_frame)
        pv_col.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 2))
        
        wind_col = ttk.Frame(new_energy_frame)
        wind_col.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(2, 0))
        
        # === 光伏列 ===
        pv_row = 0
        
        pv_material_frame = ttk.LabelFrame(pv_col, text="直接材料 (光伏)", padding=5)
        pv_material_frame.grid(row=pv_row, column=0, sticky=(tk.W, tk.E), pady=3)
        pv_row += 1
        
        ttk.Label(pv_material_frame, text="产品动力消耗 (元/kWh):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_power_cost_var = tk.DoubleVar(value=0.02)
        ttk.Entry(pv_material_frame, textvariable=self.app.v2_pv_power_cost_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(pv_material_frame, text="其他材料成本 (万元):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_other_material_var = tk.DoubleVar(value=200.0)
        ttk.Entry(pv_material_frame, textvariable=self.app.v2_pv_other_material_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        pv_labor_frame = ttk.LabelFrame(pv_col, text="直接人工", padding=5)
        pv_labor_frame.grid(row=pv_row, column=0, sticky=(tk.W, tk.E), pady=3)
        pv_row += 1
        
        ttk.Label(pv_labor_frame, text="直接人工 (万元):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_direct_labor_var = tk.DoubleVar(value=4000.0)
        ttk.Entry(pv_labor_frame, textvariable=self.app.v2_pv_direct_labor_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        pv_manufacturing_frame = ttk.LabelFrame(pv_col, text="制造费用", padding=5)
        pv_manufacturing_frame.grid(row=pv_row, column=0, sticky=(tk.W, tk.E), pady=3)
        pv_row += 1
        
        ttk.Label(pv_manufacturing_frame, text="管理人工 (万元):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_mgmt_labor_var = tk.DoubleVar(value=1000.0)
        ttk.Entry(pv_manufacturing_frame, textvariable=self.app.v2_pv_mgmt_labor_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(pv_manufacturing_frame, text="运维费用 (万元):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_maintenance_var = tk.DoubleVar(value=1000.0)
        ttk.Entry(pv_manufacturing_frame, textvariable=self.app.v2_pv_maintenance_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        ttk.Label(pv_manufacturing_frame, text="折旧及摊销 (万元):").grid(row=2, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_depreciation_var = tk.DoubleVar(value=5000.0)
        ttk.Entry(pv_manufacturing_frame, textvariable=self.app.v2_pv_depreciation_var, width=12).grid(row=2, column=1, pady=2, padx=3)
        
        ttk.Label(pv_manufacturing_frame, text="其他制造费 (万元):").grid(row=3, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_other_manufacturing_var = tk.DoubleVar(value=500.0)
        ttk.Entry(pv_manufacturing_frame, textvariable=self.app.v2_pv_other_manufacturing_var, width=12).grid(row=3, column=1, pady=2, padx=3)
        
        pv_reserve_frame = ttk.LabelFrame(pv_col, text="备容与政府基金", padding=5)
        pv_reserve_frame.grid(row=pv_row, column=0, sticky=(tk.W, tk.E), pady=3)
        pv_row += 1
        
        ttk.Label(pv_reserve_frame, text="备容费单价 (元/kWh):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_reserve_fee_var = tk.DoubleVar(value=0.028)
        ttk.Entry(pv_reserve_frame, textvariable=self.app.v2_pv_reserve_fee_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(pv_reserve_frame, text="政府性基金 (元/kWh):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_government_fund_var = tk.DoubleVar(value=0.03)
        ttk.Entry(pv_reserve_frame, textvariable=self.app.v2_pv_government_fund_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        ttk.Label(pv_reserve_frame, text="政策性交叉补贴 (元/kWh):").grid(row=2, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_policy_subsidy_var = tk.DoubleVar(value=0.0129)
        ttk.Entry(pv_reserve_frame, textvariable=self.app.v2_pv_policy_subsidy_var, width=12).grid(row=2, column=1, pady=2, padx=3)
        
        pv_price_frame = ttk.LabelFrame(pv_col, text="销售电价", padding=5)
        pv_price_frame.grid(row=pv_row, column=0, sticky=(tk.W, tk.E), pady=3)
        pv_row += 1
        
        ttk.Label(pv_price_frame, text="销售电价 (元/kWh):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_sale_price_var = tk.DoubleVar(value=0.3)
        ttk.Entry(pv_price_frame, textvariable=self.app.v2_pv_sale_price_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        pv_other_frame = ttk.LabelFrame(pv_col, text="其他成本", padding=5)
        pv_other_frame.grid(row=pv_row, column=0, sticky=(tk.W, tk.E), pady=3)
        pv_row += 1
        
        ttk.Label(pv_other_frame, text="其他固定成本 (万元):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_other_fixed_var = tk.DoubleVar(value=0.0)
        ttk.Entry(pv_other_frame, textvariable=self.app.v2_pv_other_fixed_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(pv_other_frame, text="其他可变成本 (元/kWh):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_pv_other_variable_var = tk.DoubleVar(value=0.0)
        ttk.Entry(pv_other_frame, textvariable=self.app.v2_pv_other_variable_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        # === 风电列 ===
        wind_row = 0
        
        wind_material_frame = ttk.LabelFrame(wind_col, text="直接材料 (风电)", padding=5)
        wind_material_frame.grid(row=wind_row, column=0, sticky=(tk.W, tk.E), pady=3)
        wind_row += 1
        
        ttk.Label(wind_material_frame, text="产品动力消耗 (元/kWh):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_power_cost_var = tk.DoubleVar(value=0.02)
        ttk.Entry(wind_material_frame, textvariable=self.app.v2_wind_power_cost_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(wind_material_frame, text="其他材料成本 (万元):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_other_material_var = tk.DoubleVar(value=200.0)
        ttk.Entry(wind_material_frame, textvariable=self.app.v2_wind_other_material_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        wind_labor_frame = ttk.LabelFrame(wind_col, text="直接人工", padding=5)
        wind_labor_frame.grid(row=wind_row, column=0, sticky=(tk.W, tk.E), pady=3)
        wind_row += 1
        
        ttk.Label(wind_labor_frame, text="直接人工 (万元):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_direct_labor_var = tk.DoubleVar(value=4000.0)
        ttk.Entry(wind_labor_frame, textvariable=self.app.v2_wind_direct_labor_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        wind_manufacturing_frame = ttk.LabelFrame(wind_col, text="制造费用", padding=5)
        wind_manufacturing_frame.grid(row=wind_row, column=0, sticky=(tk.W, tk.E), pady=3)
        wind_row += 1
        
        ttk.Label(wind_manufacturing_frame, text="管理人工 (万元):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_mgmt_labor_var = tk.DoubleVar(value=1000.0)
        ttk.Entry(wind_manufacturing_frame, textvariable=self.app.v2_wind_mgmt_labor_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(wind_manufacturing_frame, text="运维费用 (万元):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_maintenance_var = tk.DoubleVar(value=1000.0)
        ttk.Entry(wind_manufacturing_frame, textvariable=self.app.v2_wind_maintenance_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        ttk.Label(wind_manufacturing_frame, text="折旧及摊销 (万元):").grid(row=2, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_depreciation_var = tk.DoubleVar(value=5000.0)
        ttk.Entry(wind_manufacturing_frame, textvariable=self.app.v2_wind_depreciation_var, width=12).grid(row=2, column=1, pady=2, padx=3)
        
        ttk.Label(wind_manufacturing_frame, text="其他制造费 (万元):").grid(row=3, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_other_manufacturing_var = tk.DoubleVar(value=500.0)
        ttk.Entry(wind_manufacturing_frame, textvariable=self.app.v2_wind_other_manufacturing_var, width=12).grid(row=3, column=1, pady=2, padx=3)
        
        wind_reserve_frame = ttk.LabelFrame(wind_col, text="备容与政府基金", padding=5)
        wind_reserve_frame.grid(row=wind_row, column=0, sticky=(tk.W, tk.E), pady=3)
        wind_row += 1
        
        ttk.Label(wind_reserve_frame, text="备容费单价 (元/kWh):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_reserve_fee_var = tk.DoubleVar(value=0.028)
        ttk.Entry(wind_reserve_frame, textvariable=self.app.v2_wind_reserve_fee_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(wind_reserve_frame, text="政府性基金 (元/kWh):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_government_fund_var = tk.DoubleVar(value=0.03)
        ttk.Entry(wind_reserve_frame, textvariable=self.app.v2_wind_government_fund_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        ttk.Label(wind_reserve_frame, text="政策性交叉补贴 (元/kWh):").grid(row=2, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_policy_subsidy_var = tk.DoubleVar(value=0.0129)
        ttk.Entry(wind_reserve_frame, textvariable=self.app.v2_wind_policy_subsidy_var, width=12).grid(row=2, column=1, pady=2, padx=3)
        
        wind_price_frame = ttk.LabelFrame(wind_col, text="销售电价", padding=5)
        wind_price_frame.grid(row=wind_row, column=0, sticky=(tk.W, tk.E), pady=3)
        wind_row += 1
        
        ttk.Label(wind_price_frame, text="销售电价 (元/kWh):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_sale_price_var = tk.DoubleVar(value=0.3)
        ttk.Entry(wind_price_frame, textvariable=self.app.v2_wind_sale_price_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        wind_other_frame = ttk.LabelFrame(wind_col, text="其他成本", padding=5)
        wind_other_frame.grid(row=wind_row, column=0, sticky=(tk.W, tk.E), pady=3)
        wind_row += 1
        
        ttk.Label(wind_other_frame, text="其他固定成本 (万元):").grid(row=0, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_other_fixed_var = tk.DoubleVar(value=0.0)
        ttk.Entry(wind_other_frame, textvariable=self.app.v2_wind_other_fixed_var, width=12).grid(row=0, column=1, pady=2, padx=3)
        
        ttk.Label(wind_other_frame, text="其他可变成本 (元/kWh):").grid(row=1, column=0, sticky=tk.W, pady=2, padx=3)
        self.app.v2_wind_other_variable_var = tk.DoubleVar(value=0.0)
        ttk.Entry(wind_other_frame, textvariable=self.app.v2_wind_other_variable_var, width=12).grid(row=1, column=1, pady=2, padx=3)
        
        pv_col.columnconfigure(0, weight=1)
        wind_col.columnconfigure(0, weight=1)
        
        # --- 第 3 列：下网成本及绿证 ---
        grid_frame = ttk.LabelFrame(top_input_frame, text="3. 下网成本及绿证", padding=5)
        grid_frame.grid(row=0, column=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        grid_row = 0
        
        ttk.Label(grid_frame, text="基本电费 (万元):").grid(row=grid_row, column=0, sticky=tk.W, pady=3, padx=5)
        self.app.v2_grid_base_cost_var = tk.DoubleVar(value=3485.0)
        ttk.Entry(grid_frame, textvariable=self.app.v2_grid_base_cost_var, width=15).grid(row=grid_row, column=1, pady=3, padx=5)
        grid_row += 1
        
        ttk.Label(grid_frame, text="输配电费 (元/kWh):").grid(row=grid_row, column=0, sticky=tk.W, pady=3, padx=5)
        self.app.v2_grid_transmission_var = tk.DoubleVar(value=0.0486)
        ttk.Entry(grid_frame, textvariable=self.app.v2_grid_transmission_var, width=15).grid(row=grid_row, column=1, pady=3, padx=5)
        grid_row += 1
        
        ttk.Label(grid_frame, text="线损费用 (元/kWh):").grid(row=grid_row, column=0, sticky=tk.W, pady=3, padx=5)
        self.app.v2_grid_line_loss_var = tk.DoubleVar(value=0.017)
        ttk.Entry(grid_frame, textvariable=self.app.v2_grid_line_loss_var, width=15).grid(row=grid_row, column=1, pady=3, padx=5)
        grid_row += 1
        
        ttk.Label(grid_frame, text="系统运行费 (元/kWh):").grid(row=grid_row, column=0, sticky=tk.W, pady=3, padx=5)
        self.app.v2_grid_operation_var = tk.DoubleVar(value=0.065)
        ttk.Entry(grid_frame, textvariable=self.app.v2_grid_operation_var, width=15).grid(row=grid_row, column=1, pady=3, padx=5)
        grid_row += 1
        
        ttk.Label(grid_frame, text="政府性基金 (元/kWh):").grid(row=grid_row, column=0, sticky=tk.W, pady=3, padx=5)
        self.app.v2_grid_government_fund_var = tk.DoubleVar(value=0.0041)
        ttk.Entry(grid_frame, textvariable=self.app.v2_grid_government_fund_var, width=15).grid(row=grid_row, column=1, pady=3, padx=5)
        grid_row += 1
        
        green_cert_frame = ttk.LabelFrame(grid_frame, text="新能源绿证", padding=5)
        green_cert_frame.grid(row=grid_row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(green_cert_frame, text="绿证单价 (元/kWh):").grid(row=0, column=0, sticky=tk.W, pady=3, padx=5)
        self.app.v2_green_cert_price_var = tk.DoubleVar(value=0.008)
        ttk.Entry(green_cert_frame, textvariable=self.app.v2_green_cert_price_var, width=15).grid(row=0, column=1, pady=3, padx=5)
        ttk.Label(green_cert_frame, text="(光伏、风电共用)", foreground="gray").grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5)
        
        # 配置权重
        top_input_frame.columnconfigure(0, weight=2)
        top_input_frame.columnconfigure(1, weight=2)
        top_input_frame.columnconfigure(2, weight=2)
        top_input_frame.columnconfigure(3, weight=3)
        top_input_frame.rowconfigure(0, weight=1)
        
        # ===== 年度汇总表格 =====
        right_summary_frame = ttk.LabelFrame(top_input_frame, text="年度汇总", padding=5)
        right_summary_frame.grid(row=0, column=3, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        columns = ('项目', '数值', '备注')
        self.app.v2_summary_tree = ttk.Treeview(right_summary_frame, columns=columns, show='headings', height=25)
        self.app.v2_summary_tree.heading('项目', text='项目')
        self.app.v2_summary_tree.heading('数值', text='数值')
        self.app.v2_summary_tree.heading('备注', text='备注')
        self.app.v2_summary_tree.column('项目', width=160)
        self.app.v2_summary_tree.column('数值', width=140)
        self.app.v2_summary_tree.column('备注', width=100)
        
        summary_scrollbar = ttk.Scrollbar(right_summary_frame, orient=tk.VERTICAL, command=self.app.v2_summary_tree.yview)
        self.app.v2_summary_tree.configure(yscrollcommand=summary_scrollbar.set)
        
        self.app.v2_summary_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        summary_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        self.app.init_v2_cost_summary()
        
        right_summary_frame.columnconfigure(0, weight=1)
        right_summary_frame.rowconfigure(0, weight=1)
        
        # ===== 下方结果展示区域（8760 小时图表）=====
        bottom_result_frame = ttk.Frame(tab)
        bottom_result_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0))
        
        left_plot_frame = ttk.LabelFrame(bottom_result_frame, text="8760 小时发电出力及下网负荷趋势", padding=5)
        left_plot_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        self.app.v2_cost_figure = Figure(figsize=(12, 6), dpi=100)
        self.app.v2_cost_ax = self.app.v2_cost_figure.add_subplot(111)
        self.app.v2_cost_canvas = FigureCanvasTkAgg(self.app.v2_cost_figure, left_plot_frame)
        self.app.v2_cost_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.app.v2_cost_ax.text(0.5, 0.5, '暂无数据\n请先进行平衡计算', 
                             horizontalalignment='center', verticalalignment='center',
                             transform=self.app.v2_cost_ax.transAxes, fontsize=12)
        self.app.v2_cost_ax.set_title('8760 小时发电出力及下网负荷趋势')
        self.app.v2_cost_canvas.draw()
        
        bottom_result_frame.columnconfigure(0, weight=1)
        bottom_result_frame.rowconfigure(0, weight=1)
        left_plot_frame.columnconfigure(0, weight=1)
        left_plot_frame.rowconfigure(0, weight=1)
