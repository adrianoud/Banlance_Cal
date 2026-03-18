"""
成本详细分析模块
实现发电成本的详细分析和计算
"""

import numpy as np


class DetailedCostAnalyzer:
    """成本详细分析器"""
    
    def __init__(self, data_model, results):
        """
        初始化成本详细分析器
        :param data_model: 数据模型
        :param results: 平衡计算结果
        """
        self.data_model = data_model
        self.results = results
    
    def calculate_grid_cost(self):
        """
        计算下网成本相关指标
        :return: 包含下网总成本、度电成本等信息的字典
        """
        # 获取页面输入的基本电费（万元）
        grid_base_cost = self.data_model.detailed_cost_params.get('grid_base_cost', 3485.0)
        
        # 获取新增的 4 项附加费用（元/kWh）
        grid_transmission_price = self.data_model.detailed_cost_params.get('grid_transmission_price', 0.0486)
        grid_line_loss_price = self.data_model.detailed_cost_params.get('grid_line_loss_price', 0.017)
        grid_operation_price = self.data_model.detailed_cost_params.get('grid_operation_price', 0.065)
        grid_government_fund = self.data_model.detailed_cost_params.get('grid_government_fund', 0.0041)
        
        # 计算综合电价（元/kWh）= 下网电价 + 4 项附加费用
        additional_fees = grid_transmission_price + grid_line_loss_price + grid_operation_price + grid_government_fund
        
        # 获取每小时下网电价（元/kWh）
        grid_price_hourly = self.data_model.grid_purchase_price_hourly
        
        # 获取平衡计算结果中的每小时下网负荷（kW）
        if 'hourly_grid_load' in self.results:
            grid_load_hourly = self.results['hourly_grid_load']
        else:
            # 如果没有找到下网负荷数据，返回默认值
            return {
                'total_cost': grid_base_cost,  # 只有基本电费
                'unit_cost': 0.0,
                'total_energy': 0.0,
                'variable_cost': 0.0,
                'additional_fees_cost': 0.0
            }
        
        # 计算每小时下网电费（元）
        # 下网负荷（kW）× (下网电价 + 附加费用)（元/kWh）= 电费（元/小时）
        hourly_cost_yuan = np.array(grid_load_hourly) * (np.array(grid_price_hourly) + additional_fees)
        
        # 计算全年下网电费总和（元），然后转换为万元
        total_variable_cost_yuan = np.sum(hourly_cost_yuan)
        total_variable_cost_wan = total_variable_cost_yuan / 10000.0  # 转换为万元
        
        # 单独计算附加费用部分（用于分析）
        additional_fees_cost_yuan = np.sum(np.array(grid_load_hourly) * additional_fees)
        additional_fees_cost_wan = additional_fees_cost_yuan / 10000.0
        
        # 下网总成本 = 变动成本（含附加费用）+ 基本电费（万元）
        total_grid_cost_wan = total_variable_cost_wan + grid_base_cost
        
        # 计算总下网电量（kWh）
        total_grid_energy_kwh = np.sum(grid_load_hourly)
        
        # 计算下网度电成本（元/kWh）
        # 度电成本 = 总成本（万元）× 10000 / 总电量（kWh）
        if total_grid_energy_kwh > 0:
            unit_cost_yuan_kwh = total_grid_cost_wan * 10000.0 / total_grid_energy_kwh
        else:
            unit_cost_yuan_kwh = 0.0
        
        return {
            'total_cost': total_grid_cost_wan,  # 万元
            'unit_cost': unit_cost_yuan_kwh,    # 元/kWh
            'total_energy': total_grid_energy_kwh,  # kWh
            'variable_cost': total_variable_cost_wan,  # 万元（变动成本，含附加费用）
            'base_cost': grid_base_cost,  # 万元（基本电费）
            'additional_fees_cost': additional_fees_cost_wan  # 万元（附加费用部分）
        }
    
    def calculate_pv_cost(self):
        """
        计算光伏成本相关指标
        :return: 包含光伏总成本、度电成本等信息的字典
        """
        # 获取页面输入的光伏成本参数
        pv_depreciation = self.data_model.detailed_cost_params.get('pv_depreciation', 14000.0)  # 万元
        pv_backup_fee = self.data_model.detailed_cost_params.get('pv_backup_fee', 0.069)  # 元/kWh
        pv_unit_price = self.data_model.detailed_cost_params.get('pv_unit_price', 0.3)  # 元/kWh
        
        # 获取平衡计算结果中的每小时光伏出力（kW）
        if 'hourly_pv_output' in self.results:
            pv_output_hourly = self.results['hourly_pv_output']
        else:
            return {
                'total_cost': pv_depreciation,
                'unit_cost': 0.0,
                'total_energy': 0.0,
                'total_cost_by_unit_price': 0.0,
                'unit_cost_by_unit_price': 0.0
            }
        
        # 计算光伏总出力（kWh）
        total_pv_energy_kwh = np.sum(pv_output_hourly)
        
        # 方法 1：光伏总成本 = 折旧成本 + 备份费和政府基金
        backup_fee_total_wan = total_pv_energy_kwh * pv_backup_fee / 10000.0  # 转换为万元
        total_pv_cost_wan = pv_depreciation + backup_fee_total_wan
        
        # 计算光伏度电成本（元/kWh）
        if total_pv_energy_kwh > 0:
            unit_cost_yuan_kwh = total_pv_cost_wan * 10000.0 / total_pv_energy_kwh
        else:
            unit_cost_yuan_kwh = 0.0
        
        # 方法 2：基于度电价格的成本计算（不考虑折旧）
        # 总成本 = (度电价格 + 备份费和政府基金) × 总发电量
        total_pv_cost_by_unit_price_wan = total_pv_energy_kwh * (pv_unit_price + pv_backup_fee) / 10000.0  # 万元
        if total_pv_energy_kwh > 0:
            unit_cost_by_unit_price_yuan_kwh = (pv_unit_price + pv_backup_fee)
        else:
            unit_cost_by_unit_price_yuan_kwh = 0.0
        
        return {
            'total_cost': total_pv_cost_wan,  # 万元（基于投资成本）
            'unit_cost': unit_cost_yuan_kwh,  # 元/kWh（基于投资成本）
            'total_energy': total_pv_energy_kwh,  # kWh
            'total_cost_by_unit_price': total_pv_cost_by_unit_price_wan,  # 万元（基于度电价格）
            'unit_cost_by_unit_price': unit_cost_by_unit_price_yuan_kwh  # 元/kWh（基于度电价格）
        }
    
    def calculate_wind_cost(self):
        """
        计算风电成本相关指标
        :return: 包含风电总成本、度电成本等信息的字典
        """
        # 获取页面输入的风电成本参数
        wind_depreciation = self.data_model.detailed_cost_params.get('wind_depreciation', 10000.0)  # 万元
        wind_backup_fee = self.data_model.detailed_cost_params.get('wind_backup_fee', 0.05)  # 元/kWh
        wind_unit_price = self.data_model.detailed_cost_params.get('wind_unit_price', 0.3)  # 元/kWh
        
        # 获取平衡计算结果中的每小时风电出力（kW）
        if 'hourly_wind_output' in self.results:
            wind_output_hourly = self.results['hourly_wind_output']
        else:
            return {
                'total_cost': wind_depreciation,
                'unit_cost': 0.0,
                'total_energy': 0.0,
                'total_cost_by_unit_price': 0.0,
                'unit_cost_by_unit_price': 0.0
            }
        
        # 计算风电总出力（kWh）
        total_wind_energy_kwh = np.sum(wind_output_hourly)
        
        # 方法 1：风电总成本 = 折旧成本 + 备份费和政府基金
        backup_fee_total_wan = total_wind_energy_kwh * wind_backup_fee / 10000.0  # 转换为万元
        total_wind_cost_wan = wind_depreciation + backup_fee_total_wan
        
        # 计算风电度电成本（元/kWh）
        if total_wind_energy_kwh > 0:
            unit_cost_yuan_kwh = total_wind_cost_wan * 10000.0 / total_wind_energy_kwh
        else:
            unit_cost_yuan_kwh = 0.0
        
        # 方法 2：基于度电价格的成本计算（不考虑折旧）
        # 总成本 = (度电价格 + 备份费和政府基金) × 总发电量
        total_wind_cost_by_unit_price_wan = total_wind_energy_kwh * (wind_unit_price + wind_backup_fee) / 10000.0  # 万元
        if total_wind_energy_kwh > 0:
            unit_cost_by_unit_price_yuan_kwh = (wind_unit_price + wind_backup_fee)
        else:
            unit_cost_by_unit_price_yuan_kwh = 0.0
        
        return {
            'total_cost': total_wind_cost_wan,  # 万元（基于投资成本）
            'unit_cost': unit_cost_yuan_kwh,  # 元/kWh（基于投资成本）
            'total_energy': total_wind_energy_kwh,  # kWh
            'total_cost_by_unit_price': total_wind_cost_by_unit_price_wan,  # 万元（基于度电价格）
            'unit_cost_by_unit_price': unit_cost_by_unit_price_yuan_kwh  # 元/kWh（基于度电价格）
        }
    
    def calculate_thermal_cost(self):
        """
        计算火电成本相关指标
        :return: 包含火电总成本、度电成本等信息的字典
        """
        # 获取页面输入的火电成本参数
        thermal_manufacturing_cost = self.data_model.detailed_cost_params.get('thermal_manufacturing_cost', 19535.0)  # 万元
        thermal_labor_cost = self.data_model.detailed_cost_params.get('thermal_labor_cost', 4046.0)  # 万元
        thermal_government_fund = self.data_model.detailed_cost_params.get('thermal_government_fund', 0.0241)  # 元/kWh
        thermal_policy_subsidy = self.data_model.detailed_cost_params.get('thermal_policy_subsidy', 0.0128)  # 元/kWh
        thermal_reserve_fee_mode = self.data_model.detailed_cost_params.get('thermal_reserve_fee_mode', 0)  # 0=单价计费，1=总价计费
        thermal_reserve_fee_unit = self.data_model.detailed_cost_params.get('thermal_reserve_fee_unit', 0.05)  # 元/kWh
        thermal_reserve_fee_total = self.data_model.detailed_cost_params.get('thermal_reserve_fee_total', 0.0)  # 万元
        
        thermal_coal_price = self.data_model.detailed_cost_params.get('thermal_coal_price', 162.4)  # 元/吨
        thermal_base_coal = self.data_model.detailed_cost_params.get('thermal_base_coal', 500.0)  # g/kWh
        thermal_carbon_alloc = self.data_model.detailed_cost_params.get('thermal_carbon_alloc', 0.8049)
        thermal_carbon_price = self.data_model.detailed_cost_params.get('thermal_carbon_price', 80.0)  # 元/吨
        thermal_green_ratio = self.data_model.detailed_cost_params.get('thermal_green_ratio', 0.3)
        thermal_green_cert_price = self.data_model.detailed_cost_params.get('thermal_green_cert_price', 0.008)  # 元/kWh
        
        # 获取平衡计算结果中的每小时火电出力（kW）
        if 'hourly_thermal_output' in self.results:
            thermal_output_hourly = self.results['hourly_thermal_output']
        else:
            return {
                'total_cost': thermal_manufacturing_cost + thermal_labor_cost,
                'unit_cost': 0.0,
                'total_energy': 0.0
            }
        
        # 计算火电总出力（kWh）
        total_thermal_energy_kwh = np.sum(thermal_output_hourly)
        
        # 1. 燃料成本计算 - 基于 biz_req.md L373-L377
        # 计算每小时的燃料成本
        hourly_fuel_cost_yuan = []
        
        # 获取火电煤耗变化曲线参数
        thermal_coal_quadratic = self.data_model.detailed_cost_params.get('thermal_coal_quadratic', 0.0)
        thermal_coal_linear = self.data_model.detailed_cost_params.get('thermal_coal_linear', -0.4)
        thermal_coal_constant = self.data_model.detailed_cost_params.get('thermal_coal_constant', 1.4)
        
        # 计算火电最大出力（从数据模型获取）
        thermal_max_output = self.data_model.peak_power_max + self.data_model.chp_electric_params.get('base_electric', 0.0)
        
        for hour in range(8760):
            if hour < len(thermal_output_hourly):
                thermal_output_kw = thermal_output_hourly[hour]
                
                # 计算相对出力 = 火电出力 / 最大出力
                if thermal_max_output > 0:
                    relative_output = thermal_output_kw / thermal_max_output
                else:
                    relative_output = 0.0
                
                # 根据火电煤耗变化曲线计算相对煤耗
                # 相对煤耗 = 二次项系数 × 相对出力² + 一次项系数 × 相对出力 + 常数项
                relative_coal_consumption = (
                    thermal_coal_quadratic * relative_output**2 +
                    thermal_coal_linear * relative_output +
                    thermal_coal_constant
                )
                
                # 度电煤耗（g/kWh）= 基准煤耗 × 相对煤耗
                coal_consumption_per_kwh = thermal_base_coal * relative_coal_consumption
                
                # 火电出力单价（元/kWh）
                # = 度电煤耗 × 入炉煤单价 / 1000000 + 政府性基金 + 政策性交叉补贴 + 备容费
                coal_cost_part = coal_consumption_per_kwh * thermal_coal_price / 1000000.0
                
                # 备容费部分
                if thermal_reserve_fee_mode == 0:
                    # 单价计费方式
                    reserve_fee_part = thermal_reserve_fee_unit
                else:
                    # 总价计费方式，按发电量分摊到每度电
                    if total_thermal_energy_kwh > 0:
                        reserve_fee_part = thermal_reserve_fee_total * 10000.0 / total_thermal_energy_kwh
                    else:
                        reserve_fee_part = 0.0
                
                unit_price = coal_cost_part + thermal_government_fund + thermal_policy_subsidy + reserve_fee_part
                
                # 该小时的燃料成本（元）
                hourly_fuel_cost = thermal_output_kw * unit_price
                hourly_fuel_cost_yuan.append(hourly_fuel_cost)
            else:
                hourly_fuel_cost_yuan.append(0.0)
        
        # 燃料成本总和（万元）
        fuel_cost_wan = np.sum(hourly_fuel_cost_yuan) / 10000.0
        
        # 2. 火电固定成本 = 制造成本 + 人工成本 + 总价计费下容量费
        if thermal_reserve_fee_mode == 1:
            # 总价计费方式，容量费已包含在总额中
            fixed_cost_wan = thermal_manufacturing_cost + thermal_labor_cost
        else:
            # 单价计费方式，无额外固定容量费
            fixed_cost_wan = thermal_manufacturing_cost + thermal_labor_cost
        
        # 3. 碳排放成本（万元）- 根据 biz_req.md L379
        # 碳排放量 = 火电总发电量 × 基准煤耗 / 1000000 × 0.0267 × 19.5 × 0.99 × 44/12
        carbon_emissions_tons = total_thermal_energy_kwh * thermal_base_coal / 1000000.0 * 0.0267 * 19.5 * 0.99 * 44.0 / 12.0
        # 扣除碳配额：火电总发电量 × 碳排放分配系数 / 1000
        carbon_quota_tons = total_thermal_energy_kwh * thermal_carbon_alloc / 1000.0
        # 实际碳排放量
        actual_carbon_tons = carbon_emissions_tons - carbon_quota_tons
        # 碳排放成本（万元）
        carbon_cost_wan = actual_carbon_tons * thermal_carbon_price / 10000.0
        
        # 4. 绿证费用（万元）- 根据 biz_req.md L380-L381
        # 获取光伏和风电的实际发电量
        pv_output_hourly = self.results.get('hourly_pv_output', [0.0] * 8760)
        wind_output_hourly = self.results.get('hourly_wind_output', [0.0] * 8760)
        total_pv_wind_energy = np.sum(pv_output_hourly) + np.sum(wind_output_hourly)
        
        # 总用电量 = 火电 + 光伏 + 风电 + 下网
        grid_load_hourly = self.results.get('hourly_grid_load', [0.0] * 8760)
        total_energy_consumption = total_thermal_energy_kwh + total_pv_wind_energy + np.sum(grid_load_hourly)
        
        # 绿电实际占比
        actual_green_ratio = total_pv_wind_energy / total_energy_consumption if total_energy_consumption > 0 else 0.0
        
        # 绿证费用 = (绿色能源占比 - 绿电实际占比) × 总用电量 × 绿证单价 / 10000
        green_cert_gap = thermal_green_ratio - actual_green_ratio
        if green_cert_gap > 0:
            green_cert_cost_wan = green_cert_gap * total_energy_consumption * thermal_green_cert_price / 10000.0
        else:
            green_cert_cost_wan = 0.0
        
        # 火电总成本 = 燃料成本 + 固定成本 + 碳排费用 + 绿证费用
        total_thermal_cost_wan = fuel_cost_wan + fixed_cost_wan + carbon_cost_wan + green_cert_cost_wan
        
        # 计算火电度电成本（元/kWh）
        if total_thermal_energy_kwh > 0:
            unit_cost_yuan_kwh = total_thermal_cost_wan * 10000.0 / total_thermal_energy_kwh
        else:
            unit_cost_yuan_kwh = 0.0
        
        return {
            'total_cost': total_thermal_cost_wan,  # 万元
            'unit_cost': unit_cost_yuan_kwh,  # 元/kWh
            'total_energy': total_thermal_energy_kwh,  # kWh
            'fixed_cost': fixed_cost_wan,  # 万元
            'fuel_cost': fuel_cost_wan,  # 万元
            'carbon_cost': carbon_cost_wan,  # 万元
            'green_cert_cost': green_cert_cost_wan  # 万元
        }
    
    def calculate_other_cost(self):
        """
        计算其他成本相关指标
        :return: 包含其他总成本、度电成本等信息的字典
        """
        # 获取页面输入的其他成本参数
        other_fixed_cost = self.data_model.detailed_cost_params.get('other_fixed_cost', 0.0)  # 万元
        other_variable_cost = self.data_model.detailed_cost_params.get('other_variable_cost', 0.0)  # 元/kWh
        
        # 获取所有电源的总出力
        total_energy = 0.0
        if 'hourly_thermal_output' in self.results:
            total_energy += np.sum(self.results['hourly_thermal_output'])
        if 'hourly_pv_output' in self.results:
            total_energy += np.sum(self.results['hourly_pv_output'])
        if 'hourly_wind_output' in self.results:
            total_energy += np.sum(self.results['hourly_wind_output'])
        
        # 其他总成本 = 固定成本 + 可变成本
        variable_cost_wan = total_energy * other_variable_cost / 10000.0  # 转换为万元
        total_other_cost_wan = other_fixed_cost + variable_cost_wan
        
        # 计算其他度电成本（元/kWh）
        if total_energy > 0:
            unit_cost_yuan_kwh = total_other_cost_wan * 10000.0 / total_energy
        else:
            unit_cost_yuan_kwh = 0.0
        
        return {
            'total_cost': total_other_cost_wan,  # 万元
            'unit_cost': unit_cost_yuan_kwh,  # 元/kWh
            'total_energy': total_energy,  # kWh
            'fixed_cost': other_fixed_cost,  # 万元
            'variable_cost': variable_cost_wan  # 万元
        }
    
    def calculate_total_summary(self):
        """
        计算成本汇总
        :return: 包含所有成本汇总信息的字典
        """
        # 分别计算各项成本
        grid_cost = self.calculate_grid_cost()
        pv_cost = self.calculate_pv_cost()
        wind_cost = self.calculate_wind_cost()
        thermal_cost = self.calculate_thermal_cost()
        other_cost = self.calculate_other_cost()
        
        # 计算总成本（万元）- 基于投资成本
        total_cost_wan = (grid_cost['total_cost'] + pv_cost['total_cost'] + 
                         wind_cost['total_cost'] + thermal_cost['total_cost'] + 
                         other_cost['total_cost'])
        
        # 计算总成本（万元）- 基于度电价格
        # 只有光伏和风电使用度电价格算法，其他保持不变
        total_cost_by_unit_price_wan = (grid_cost['total_cost'] + pv_cost['total_cost_by_unit_price'] + 
                                       wind_cost['total_cost_by_unit_price'] + thermal_cost['total_cost'] + 
                                       other_cost['total_cost'])
        
        # 计算总发电量（kWh）
        total_generation_kwh = (pv_cost['total_energy'] + wind_cost['total_energy'] + 
                               thermal_cost['total_energy'])
        
        # 计算平均小时成本 - 基于投资成本
        avg_hourly_cost_wan = total_cost_wan / 8760.0
        
        # 计算平均小时成本 - 基于度电价格
        avg_hourly_cost_by_unit_price_wan = total_cost_by_unit_price_wan / 8760.0
        
        return {
            'total_cost': total_cost_wan,
            'total_cost_by_unit_price': total_cost_by_unit_price_wan,
            'avg_hourly_cost': avg_hourly_cost_wan,
            'avg_hourly_cost_by_unit_price': avg_hourly_cost_by_unit_price_wan,
            'total_generation': total_generation_kwh,
            'grid': grid_cost,
            'pv': pv_cost,
            'wind': wind_cost,
            'thermal': thermal_cost,
            'other': other_cost
        }
