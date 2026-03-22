"""
成本分析 2.0 计算模块
基于 biz_req.md L446-L469 的成本计算逻辑
实现火电、光伏、风电、下网的精细化成本计算
"""

import numpy as np


class CostV2Calculator:
    """成本分析 2.0 计算器"""
    
    def __init__(self, data_model, results):
        """
        初始化成本分析 2.0 计算器
        :param data_model: 数据模型（包含 cost_v2_params 参数）
        :param results: 平衡计算结果（包含 hourly_*, hourly_corrected_* 等）
        """
        self.data_model = data_model
        self.results = results
    
    def calculate_all(self):
        """
        计算所有成本指标
        :return: 包含所有成本汇总信息的字典
        """
        # 分别计算各类成本
        thermal_cost = self.calculate_thermal_cost()
        pv_cost = self.calculate_pv_cost()
        wind_cost = self.calculate_wind_cost()
        grid_cost = self.calculate_grid_cost()
        
        # 汇总计算
        total_cost = (
            thermal_cost['total_cost'] + 
            pv_cost['total_cost'] + 
            wind_cost['total_cost'] + 
            grid_cost['total_cost']
        )
        
        # 计算净绿证购买费用
        # 火电不再购买绿证，只考虑新能源的绿证抵扣收入
        green_cert_purchase = 0.0  # 火电购买绿证（万元）- 不再考虑
        green_cert_income_total = pv_cost['green_cert_income'] + wind_cost['green_cert_income']  # 新能源绿证收入（万元）
        net_green_cert_cost = green_cert_purchase - green_cert_income_total  # 净绿证费用（万元）
        
        return {
            'thermal': thermal_cost,
            'pv': pv_cost,
            'wind': wind_cost,
            'grid': grid_cost,
            'summary': {
                'total_cost': total_cost,
                'total_energy': thermal_cost['total_energy'] + pv_cost['total_energy'] + wind_cost['total_energy'],
                'thermal_ratio': thermal_cost['total_cost'] / total_cost * 100 if total_cost > 0 else 0,
                'pv_ratio': pv_cost['total_cost'] / total_cost * 100 if total_cost > 0 else 0,
                'wind_ratio': wind_cost['total_cost'] / total_cost * 100 if total_cost > 0 else 0,
                'grid_ratio': grid_cost['total_cost'] / total_cost * 100 if total_cost > 0 else 0,
                'green_cert_purchase': green_cert_purchase,  # 购买绿证费用（未抵扣）
                'green_cert_income_total': green_cert_income_total,  # 新能源绿证抵扣总额
                'net_green_cert_cost': net_green_cert_cost  # 净绿证费用（已抵扣）
            }
        }
    
    def calculate_thermal_cost(self):
        """
        计算火电成本
        火电总成本 = 火电可变成本 + 火电固定成本 + 碳排费用 + 购买绿证费用
        
        :return: 包含火电成本详细分项的字典
        """
        # ===== 1. 获取参数 =====
        params = self.data_model.cost_v2_params if hasattr(self.data_model, 'cost_v2_params') else {}
        
        # 直接材料
        coal_price = params.get('thermal_coal_price', 162.4)  # 元/吨
        base_coal = params.get('thermal_base_coal', 500.0)  # g/kWh
        pure_water_cost = params.get('thermal_pure_water_cost', 100.0)  # 万元
        pure_water_unit = params.get('thermal_pure_water_unit', 0.05)  # 元/kWh
        
        # 直接人工
        direct_labor = params.get('thermal_direct_labor', 4000.0)  # 万元
        
        # 制造费用
        mgmt_labor = params.get('thermal_mgmt_labor', 1000.0)  # 万元
        maintenance = params.get('thermal_maintenance', 1000.0)  # 万元
        depreciation = params.get('thermal_depreciation', 8000.0)  # 万元
        other_manufacturing = params.get('thermal_other_manufacturing', 500.0)  # 万元
        
        # 备容与政府基金
        reserve_fee_mode = params.get('thermal_reserve_fee_mode', 0)  # 0=单价，1=总价
        reserve_fee_unit = params.get('thermal_reserve_fee_unit', 0.05)  # 元/kWh
        reserve_fee_total = params.get('thermal_reserve_fee_total', 0.0)  # 万元
        government_fund = params.get('thermal_government_fund', 0.0241)  # 元/kWh
        policy_subsidy = params.get('thermal_policy_subsidy', 0.0128)  # 元/kWh
        
        # 碳排放和绿证
        carbon_intensity = params.get('thermal_carbon_intensity', 0.8049)  # kg/kWh
        heat_value = params.get('thermal_heat_value', 19.5)  # MJ/kg
        carbon_content = params.get('thermal_carbon_content', 0.0267)  # kg/MJ
        carbon_price = params.get('thermal_carbon_price', 80.0)  # 元/吨
        green_ratio = params.get('thermal_green_ratio', 0.3)  # 绿证抵扣比例
        
        # 其他成本
        other_fixed = params.get('thermal_other_fixed', 0.0)  # 万元
        other_variable = params.get('thermal_other_variable', 0.0)  # 元/kWh
        
        # 煤耗曲线参数
        coal_quadratic = params.get('thermal_coal_quadratic', 0.0)
        coal_linear = params.get('thermal_coal_linear', -0.4)
        coal_constant = params.get('thermal_coal_constant', 1.4)
        
        # ===== 2. 获取平衡计算结果 =====
        if 'hourly_thermal_output' not in self.results:
            # 没有平衡计算结果，返回默认值
            return {
                'total_cost': 0.0,
                'variable_cost': 0.0,
                'fixed_cost': 0.0,
                'carbon_cost': 0.0,
                'green_cert_cost': 0.0,
                'total_energy': 0.0,
                'unit_cost': 0.0
            }
        
        thermal_output_hourly = self.results['hourly_thermal_output']
        
        # ===== 3. 基础计算 =====
        # 总发电量 (kWh)
        total_thermal_energy_kwh = np.sum(thermal_output_hourly)
        
        # 最大出力 (kW) - 用于计算相对出力
        thermal_max_output = self.data_model.peak_power_max + self.data_model.chp_electric_params.get('base_electric', 0.0)
        
        # ===== 4. 火电可变成本计算 =====
        # 4.1 计算每小时的度电煤耗和燃料成本
        hourly_variable_cost_yuan = []
        
        for hour in range(8760):
            if hour < len(thermal_output_hourly):
                output_kw = thermal_output_hourly[hour]
                
                # 计算相对出力
                if thermal_max_output > 0:
                    relative_output = output_kw / thermal_max_output
                else:
                    relative_output = 0.0
                
                # 计算相对煤耗（二次曲线）
                relative_coal = (
                    coal_quadratic * relative_output**2 +
                    coal_linear * relative_output +
                    coal_constant
                )
                
                # 度电煤耗 (g/kWh)
                coal_consumption = base_coal * relative_coal
                
                # 燃料成本 (元/kWh) = 度电煤耗 × 原煤单价 / 10^6
                fuel_cost_per_kwh = coal_consumption * coal_price / 1000000.0
                
                # 可变成本单价 (元/kWh)
                variable_unit_cost = (
                    fuel_cost_per_kwh +  # 燃料成本
                    pure_water_unit +  # 纯水及其他
                    government_fund +  # 政府性基金
                    policy_subsidy +  # 政策性交叉补贴
                    (reserve_fee_unit if reserve_fee_mode == 0 else 0.0) +  # 容量费（单价模式）
                    other_variable  # 其他可变成本
                )
                
                # 每小时可变成本 (元)
                hourly_cost = output_kw * variable_unit_cost
                hourly_variable_cost_yuan.append(hourly_cost)
            else:
                hourly_variable_cost_yuan.append(0.0)
        
        # 可变成本总和 (万元)
        total_variable_cost_yuan = np.sum(hourly_variable_cost_yuan)
        total_variable_cost_wan = total_variable_cost_yuan / 10000.0
        
        # 计算备容与政府基金（从可变成本中分离）
        # 备容费（单价模式）
        reserve_fee_unit_total = 0.0
        government_fund_total = 0.0
        policy_subsidy_total = 0.0
        
        for hour in range(8760):
            if hour < len(thermal_output_hourly):
                output_kw = thermal_output_hourly[hour]
                # 备容费（单价模式）
                if reserve_fee_mode == 0:
                    reserve_fee_unit_total += output_kw * reserve_fee_unit
                # 政府性基金
                government_fund_total += output_kw * government_fund
                # 政策性交叉补贴
                policy_subsidy_total += output_kw * policy_subsidy
        
        # 转换为万元
        reserve_fee_unit_total_wan = reserve_fee_unit_total / 10000.0
        government_fund_total_wan = government_fund_total / 10000.0
        policy_subsidy_total_wan = policy_subsidy_total / 10000.0
        
        # 容量费（总价模式）
        reserve_fee_total_wan = (reserve_fee_total if reserve_fee_mode == 1 else 0.0)
        
        # 备容与政府基金总计（万元）
        reserve_and_government_fund_total = (
            reserve_fee_unit_total_wan +  # 容量费（单价模式）
            reserve_fee_total_wan +  # 容量费（总价模式）
            government_fund_total_wan +  # 政府性基金
            policy_subsidy_total_wan  # 政策性交叉补贴
        )
        
        # ===== 5. 火电固定成本计算 =====
        # 固定成本 (万元)
        total_fixed_cost_wan = (
            pure_water_cost +  # 纯水及其他费用
            direct_labor +  # 直接人工成本
            mgmt_labor +  # 管理人工成本
            maintenance +  # 运维成本
            depreciation +  # 折旧及摊销
            other_manufacturing +  # 其他制造费用
            other_fixed +  # 其他固定成本
            (reserve_fee_total if reserve_fee_mode == 1 else 0.0)  # 容量费（总价模式）
        )
        
        # ===== 6. 碳排费用计算 =====
        # 碳排费用 = (总发电量 × 基准煤耗 × 单位热值含碳量 × 0.99 × 44/12 - 总发电量 × 碳排放强度) × 碳价
        # 注意单位转换：基准煤耗是 g/kWh，需要转换为 kg/kWh
        carbon_emission_factor = (
            base_coal / 1000.0 *  # g/kWh → kg/kWh
            carbon_content *  # kg/MJ
            heat_value *  # MJ/kg
            0.99 *  # 氧化率
            44.0 / 12.0  # C → CO2
        )  # kg CO2/kWh
        
        # 总碳排放量 (吨)
        total_carbon_emission_ton = (
            total_thermal_energy_kwh * carbon_emission_factor / 1000.0 -  # kg → 吨
            total_thermal_energy_kwh * carbon_intensity / 1000.0  # 扣除配额
        )
        
        # 碳排费用 (万元)
        carbon_cost_wan = total_carbon_emission_ton * carbon_price / 10000.0
        
        # ===== 8. 汇总 =====
        # 火电总成本 (万元) - 不考虑购买绿证费用
        total_thermal_cost_wan = (
            total_variable_cost_wan +
            total_fixed_cost_wan +
            carbon_cost_wan
        )
        
        # 度电成本 (元/kWh)
        if total_thermal_energy_kwh > 0:
            unit_cost_yuan_kwh = total_thermal_cost_wan * 10000.0 / total_thermal_energy_kwh
        else:
            unit_cost_yuan_kwh = 0.0
        
        return {
            'total_cost': total_thermal_cost_wan,  # 万元
            'variable_cost': total_variable_cost_wan,  # 万元
            'fixed_cost': total_fixed_cost_wan,  # 万元
            'carbon_cost': carbon_cost_wan,  # 万元
            'green_cert_cost': 0.0,  # 万元（不再考虑）
            'reserve_and_government_fund': reserve_and_government_fund_total,  # 万元（备容与政府基金）
            'total_energy': total_thermal_energy_kwh,  # kWh
            'unit_cost': unit_cost_yuan_kwh,  # 元/kWh
            'hourly_variable_cost': hourly_variable_cost_yuan  # 每小时可变成本 (元)，用于 8760 分析
        }
    
    def calculate_pv_cost(self):
        """
        计算光伏成本
        光伏总成本 = 可变成本 + 固定成本 - 抵扣绿证费用
        
        :return: 包含光伏成本详细分项的字典
        """
        # ===== 1. 获取参数 =====
        params = self.data_model.cost_v2_params if hasattr(self.data_model, 'cost_v2_params') else {}
        
        # 直接材料
        power_cost = params.get('pv_power_cost', 0.02)  # 元/kWh
        other_material = params.get('pv_other_material', 200.0)  # 万元
        
        # 直接人工
        direct_labor = params.get('pv_direct_labor', 4000.0)  # 万元
        
        # 制造费用
        mgmt_labor = params.get('pv_mgmt_labor', 1000.0)  # 万元
        maintenance = params.get('pv_maintenance', 1000.0)  # 万元
        depreciation = params.get('pv_depreciation', 14000.0)  # 万元
        other_manufacturing = params.get('pv_other_manufacturing', 500.0)  # 万元
        
        # 备容与政府基金
        reserve_fee = params.get('pv_reserve_fee', 0.028)  # 元/kWh
        government_fund = params.get('pv_government_fund', 0.03)  # 元/kWh
        policy_subsidy = params.get('pv_policy_subsidy', 0.0129)  # 元/kWh
        
        # 电价（备用，目前成本计算不使用）
        # sale_price = params.get('pv_sale_price', 0.3)  # 元/kWh
        
        # 其他成本
        other_fixed = params.get('pv_other_fixed', 0.0)  # 万元
        other_variable = params.get('pv_other_variable', 0.0)  # 元/kWh
        
        # **共享绿证单价** (关键参数)
        green_cert_price = params.get('green_cert_price', 0.008)  # 元/kWh
        
        # ===== 2. 获取平衡计算结果 =====
        if 'hourly_pv_output' not in self.results:
            return {
                'total_cost': 0.0,
                'variable_cost': 0.0,
                'fixed_cost': 0.0,
                'green_cert_income': 0.0,
                'total_energy': 0.0,
                'unit_cost': 0.0
            }
        
        pv_output_hourly = self.results['hourly_pv_output']
        
        # ===== 3. 基础计算 =====
        # 总发电量 (kWh)
        total_pv_energy_kwh = np.sum(pv_output_hourly)
        
        # ===== 4. 可变成本计算 =====
        # 可变成本单价 (元/kWh)
        variable_unit_cost = (
            power_cost +  # 产品动力消耗
            reserve_fee +  # 备容费
            government_fund +  # 政府性基金
            policy_subsidy +  # 政策性交叉补贴
            other_variable  # 其他可变成本
        )
        
        # 可变成本总和 (万元)
        total_variable_cost_wan = total_pv_energy_kwh * variable_unit_cost / 10000.0
        
        # ===== 5. 固定成本计算 =====
        # 固定成本 (万元)
        total_fixed_cost_wan = (
            other_material +  # 其他直接材料
            direct_labor +  # 直接人工
            mgmt_labor +  # 管理人工
            maintenance +  # 运维费
            depreciation +  # 折旧及摊销
            other_manufacturing +  # 其他制造费用
            other_fixed  # 其他固定成本
        )
        
        # ===== 6. 抵扣绿证费用计算 =====
        # 抵扣绿证费用 = 实际发电量 × 绿证单价
        green_cert_income_yuan = total_pv_energy_kwh * green_cert_price
        green_cert_income_wan = green_cert_income_yuan / 10000.0
        
        # ===== 7. 汇总 =====
        # 光伏总成本 (万元)
        total_pv_cost_wan = (
            total_variable_cost_wan +
            total_fixed_cost_wan -
            green_cert_income_wan  # 绿证收入作为抵扣项
        )
        
        # 度电成本 (元/kWh)
        if total_pv_energy_kwh > 0:
            unit_cost_yuan_kwh = total_pv_cost_wan * 10000.0 / total_pv_energy_kwh
        else:
            unit_cost_yuan_kwh = 0.0
        
        return {
            'total_cost': total_pv_cost_wan,  # 万元
            'variable_cost': total_variable_cost_wan,  # 万元
            'fixed_cost': total_fixed_cost_wan,  # 万元
            'green_cert_income': green_cert_income_wan,  # 万元（抵扣项）
            'total_energy': total_pv_energy_kwh,  # kWh
            'unit_cost': unit_cost_yuan_kwh,  # 元/kWh
            'hourly_variable_cost': [output * variable_unit_cost for output in pv_output_hourly]  # 每小时可变成本 (元)
        }
    
    def calculate_wind_cost(self):
        """
        计算风电成本
        风电总成本 = 可变成本 + 固定成本 - 抵扣绿证费用
        
        :return: 包含风电成本详细分项的字典
        """
        # ===== 1. 获取参数 =====
        params = self.data_model.cost_v2_params if hasattr(self.data_model, 'cost_v2_params') else {}
        
        # 直接材料
        power_cost = params.get('wind_power_cost', 0.02)  # 元/kWh
        other_material = params.get('wind_other_material', 200.0)  # 万元
        
        # 直接人工
        direct_labor = params.get('wind_direct_labor', 4000.0)  # 万元
        
        # 制造费用
        mgmt_labor = params.get('wind_mgmt_labor', 1000.0)  # 万元
        maintenance = params.get('wind_maintenance', 1000.0)  # 万元
        depreciation = params.get('wind_depreciation', 10000.0)  # 万元
        other_manufacturing = params.get('wind_other_manufacturing', 500.0)  # 万元
        
        # 备容与政府基金
        reserve_fee = params.get('wind_reserve_fee', 0.028)  # 元/kWh
        government_fund = params.get('wind_government_fund', 0.03)  # 元/kWh
        policy_subsidy = params.get('wind_policy_subsidy', 0.0129)  # 元/kWh
        
        # 电价（备用）
        # sale_price = params.get('wind_sale_price', 0.3)  # 元/kWh
        
        # 其他成本
        other_fixed = params.get('wind_other_fixed', 0.0)  # 万元
        other_variable = params.get('wind_other_variable', 0.0)  # 元/kWh
        
        # **共享绿证单价** (关键参数，与光伏使用同一个值)
        green_cert_price = params.get('green_cert_price', 0.008)  # 元/kWh
        
        # ===== 2. 获取平衡计算结果 =====
        if 'hourly_wind_output' not in self.results:
            return {
                'total_cost': 0.0,
                'variable_cost': 0.0,
                'fixed_cost': 0.0,
                'green_cert_income': 0.0,
                'total_energy': 0.0,
                'unit_cost': 0.0
            }
        
        wind_output_hourly = self.results['hourly_wind_output']
        
        # ===== 3. 基础计算 =====
        # 总发电量 (kWh)
        total_wind_energy_kwh = np.sum(wind_output_hourly)
        
        # ===== 4. 可变成本计算 =====
        # 可变成本单价 (元/kWh)
        variable_unit_cost = (
            power_cost +  # 产品动力消耗
            reserve_fee +  # 备容费
            government_fund +  # 政府性基金
            policy_subsidy +  # 政策性交叉补贴
            other_variable  # 其他可变成本
        )
        
        # 可变成本总和 (万元)
        total_variable_cost_wan = total_wind_energy_kwh * variable_unit_cost / 10000.0
        
        # ===== 5. 固定成本计算 =====
        # 固定成本 (万元)
        total_fixed_cost_wan = (
            other_material +  # 其他直接材料
            direct_labor +  # 直接人工
            mgmt_labor +  # 管理人工
            maintenance +  # 运维费
            depreciation +  # 折旧及摊销
            other_manufacturing +  # 其他制造费用
            other_fixed  # 其他固定成本
        )
        
        # ===== 6. 抵扣绿证费用计算 =====
        # 抵扣绿证费用 = 实际发电量 × 绿证单价
        # 注意：使用与光伏相同的共享绿证单价
        green_cert_income_yuan = total_wind_energy_kwh * green_cert_price
        green_cert_income_wan = green_cert_income_yuan / 10000.0
        
        # ===== 7. 汇总 =====
        # 风电总成本 (万元)
        total_wind_cost_wan = (
            total_variable_cost_wan +
            total_fixed_cost_wan -
            green_cert_income_wan  # 绿证收入作为抵扣项
        )
        
        # 度电成本 (元/kWh)
        if total_wind_energy_kwh > 0:
            unit_cost_yuan_kwh = total_wind_cost_wan * 10000.0 / total_wind_energy_kwh
        else:
            unit_cost_yuan_kwh = 0.0
        
        return {
            'total_cost': total_wind_cost_wan,  # 万元
            'variable_cost': total_variable_cost_wan,  # 万元
            'fixed_cost': total_fixed_cost_wan,  # 万元
            'green_cert_income': green_cert_income_wan,  # 万元（抵扣项）
            'total_energy': total_wind_energy_kwh,  # kWh
            'unit_cost': unit_cost_yuan_kwh,  # 元/kWh
            'hourly_variable_cost': [output * variable_unit_cost for output in wind_output_hourly]  # 每小时可变成本 (元)
        }
    
    def calculate_grid_cost(self):
        """
        计算下网成本
        下网总成本 = 求和（每小时下网负荷 × 下网电价）/ 10000 + 基本电费
        
        :return: 包含下网成本详细分项的字典
        """
        # ===== 1. 获取参数 =====
        params = self.data_model.cost_v2_params if hasattr(self.data_model, 'cost_v2_params') else {}
        
        # 基本电费 (万元)
        grid_base_cost = params.get('grid_base_cost', 3485.0)
        
        # 输配电费 (元/kWh)
        transmission_price = params.get('grid_transmission', 0.0486)
        
        # 线损费用 (元/kWh)
        line_loss_price = params.get('grid_line_loss', 0.017)
        
        # 系统运行费 (元/kWh)
        operation_price = params.get('grid_operation', 0.065)
        
        # 政府性基金 (元/kWh)
        government_fund = params.get('grid_government_fund', 0.0041)
        
        # 综合附加费 (元/kWh)
        additional_fees = transmission_price + line_loss_price + operation_price + government_fund
        
        # ===== 2. 获取数据 =====
        # 每小时下网电价 (元/kWh)
        grid_price_hourly = self.data_model.grid_purchase_price_hourly
        
        # 每小时下网负荷 (kW)
        if 'hourly_grid_load' not in self.results:
            return {
                'total_cost': grid_base_cost,
                'variable_cost': 0.0,
                'base_cost': grid_base_cost,
                'total_energy': 0.0,
                'unit_cost': 0.0
            }
        
        grid_load_hourly = self.results['hourly_grid_load']
        
        # ===== 3. 计算每小时下网电费 =====
        # 下网电价 = 上传数据集电价 + 附加费用
        hourly_grid_price = np.array(grid_price_hourly) + additional_fees
        
        # 每小时下网电费 (元)
        hourly_cost_yuan = np.array(grid_load_hourly) * hourly_grid_price
        
        # ===== 4. 汇总 =====
        # 变动成本总和 (万元)
        total_variable_cost_yuan = np.sum(hourly_cost_yuan)
        total_variable_cost_wan = total_variable_cost_yuan / 10000.0
        
        # 下网总成本 (万元)
        total_grid_cost_wan = total_variable_cost_wan + grid_base_cost
        
        # 总下网电量 (kWh)
        total_grid_energy_kwh = np.sum(grid_load_hourly)
        
        # 度电成本 (元/kWh)
        if total_grid_energy_kwh > 0:
            unit_cost_yuan_kwh = total_grid_cost_wan * 10000.0 / total_grid_energy_kwh
        else:
            unit_cost_yuan_kwh = 0.0
        
        return {
            'total_cost': total_grid_cost_wan,  # 万元
            'variable_cost': total_variable_cost_wan,  # 万元
            'base_cost': grid_base_cost,  # 万元
            'additional_fees_cost': np.sum(np.array(grid_load_hourly) * additional_fees) / 10000.0,  # 万元
            'total_energy': total_grid_energy_kwh,  # kWh
            'unit_cost': unit_cost_yuan_kwh,  # 元/kWh
            'hourly_cost': hourly_cost_yuan.tolist()  # 每小时下网电费 (元)
        }
