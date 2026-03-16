"""
成本优化模块
实现基于最小化发电成本的优化算法
"""

import tkinter as tk
import numpy as np
from datetime import datetime, timedelta


def start_cost_optimization(app_instance):
    """
    开始成本优化计算
    
    优化目标：每一个负荷时刻的成本最小
    计算公式：火力发电出力 × 火电发电单位成本 + 光伏发电出力 × 光伏发电单位成本 + 
             风机发电出力 × 风机发电单位成本 + 下网负荷 × 下网电价
    
    优化变量：每一个负荷时刻的 火力发电出力，光伏发电出力，风机发电出力和下网负荷
    
    优化约束：
    1. 总负荷减去厂用电负荷（电力负荷）不变
    2. 火力发电出力在调峰机组最大出力和最小出力之间
    3. 光伏出力小于光伏最大出力并大于等于 0
    4. 风机出力小于风机最大出力并大于等于 0
    5. 下网负荷不小于最小下网负荷
    """
    # 检查是否有计算结果可供优化
    if not app_instance.results:
        from tkinter import messagebox
        messagebox.showwarning("警告", "请先进行年度平衡计算，再进行优化！")
        return
    
    # 更新进度条
    app_instance.optimization_progress_label.config(text="正在成本优化计算...")
    app_instance.optimization_progress["value"] = 0
    app_instance.root.update_idletasks()
    
    data_model = app_instance.data_model
    results = app_instance.results
    calculator = app_instance.calculator
    
    # 获取约束参数
    min_grid_load = app_instance.min_grid_load.get()
    
    # 获取电价数据
    grid_price = data_model.grid_purchase_price_hourly
    
    # 创建优化后的结果字典 - 成本优化模式
    optimized_results = {
        'hourly_thermal_output': [0.0] * 8760,      # 火电出力优化值
        'hourly_pv_output': [0.0] * 8760,           # 光伏出力优化值
        'hourly_wind_output': [0.0] * 8760,         # 风电出力优化值
        'hourly_grid_load': [0.0] * 8760,           # 下网负荷优化值
        'hourly_total_generation': [0.0] * 8760,    # 总出力优化值
        'hourly_cost': [0.0] * 8760,                # 每小时成本
        'total_cost': 0.0                            # 总成本
    }
    
    # 预计算装机容量用于成本计算
    total_wind_capacity = data_model.calculate_wind_total_capacity()
    total_pv_capacity = data_model.calculate_pv_total_capacity()
    thermal_max_output = data_model.peak_power_max + data_model.chp_electric_params['base_electric']
    
    base_date = datetime(2024, 1, 1)
    
    # 逐小时优化
    for hour in range(8760):
        # 获取当前小时的原始数据
        chp_output = results['hourly_chp_output'][hour]
        pv_output_orig = results['hourly_pv_output'][hour]
        wind_output_orig = results['hourly_wind_output'][hour]
        thermal_output_orig = results['hourly_thermal_output'][hour]
        grid_load_orig = results['hourly_grid_load'][hour]
        
        # 计算当前小时的真实发电成本（使用发电成本页的成本曲线）
        # 火电成本
        thermal_relative = thermal_output_orig / thermal_max_output if thermal_max_output > 0 else 0
        thermal_relative_cost = (data_model.thermal_cost_curve['quadratic_coefficient'] * thermal_relative**2 +
                               data_model.thermal_cost_curve['linear_coefficient'] * thermal_relative +
                               data_model.thermal_cost_curve['constant_term'])
        current_thermal_cost = thermal_relative_cost * data_model.thermal_cost_curve['base_cost']
        
        # 风电成本
        wind_relative = wind_output_orig / total_wind_capacity if total_wind_capacity > 0 else 0
        wind_relative_cost = (data_model.wind_cost_curve['linear_coefficient'] * wind_relative +
                            data_model.wind_cost_curve['constant_term'])
        current_wind_cost = wind_relative_cost * data_model.wind_cost_curve['base_cost']
        
        # 光电成本
        pv_relative = pv_output_orig / total_pv_capacity if total_pv_capacity > 0 else 0
        pv_relative_cost = (data_model.pv_cost_curve['linear_coefficient'] * pv_relative +
                          data_model.pv_cost_curve['constant_term'])
        current_pv_cost = pv_relative_cost * data_model.pv_cost_curve['base_cost']
        
        # 获取当前小时的日期信息
        current_date = base_date + timedelta(hours=hour)
        current_month = current_date.month
        
        # 获取当前小时的活动检修计划和投产计划
        active_maintenance_schedules = calculator.get_active_maintenance_schedules(hour)
        active_commissioning_schedules = calculator.get_active_commissioning_schedules(hour)
        
        # 预先计算修正后的调峰机组参数
        current_peak_power_max = data_model.peak_power_max
        if 5 <= current_month <= 9:  # 夏季
            current_peak_power_min = data_model.peak_power_min_summer
        else:  # 冬季
            current_peak_power_min = data_model.peak_power_min_winter
        
        # 应用检修计划对调峰机组参数的修正
        for schedule in active_maintenance_schedules:
            power_type = schedule.get('power_type', '')
            power_size = schedule.get('power_size', 0.0)
            
            if power_type == '调峰机组出力':
                current_peak_power_max = current_peak_power_max - power_size
                original_peak_power_max = data_model.peak_power_max
                if original_peak_power_max > 0:
                    if 5 <= current_month <= 9:  # 夏季
                        current_peak_power_min = data_model.peak_power_min_summer * (current_peak_power_max / original_peak_power_max)
                    else:  # 冬季
                        current_peak_power_min = data_model.peak_power_min_winter * (current_peak_power_max / original_peak_power_max)
        
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
                if 5 <= current_month <= 9:  # 夏季：5-9 月
                    current_peak_power_min = data_model.peak_power_min_summer - adjusted_power_size
            elif power_type == '调峰机组冬季最小出力':
                adjusted_power_size = power_size * interpolation_factor
                if current_date.month < 5 or current_date.month > 9:  # 冬季：10-12 月和 1-4 月
                    current_peak_power_min = data_model.peak_power_min_winter - adjusted_power_size
            elif power_type == '调峰机组最小出力':
                adjusted_power_size = power_size * interpolation_factor
                if 5 <= current_month <= 9:  # 夏季
                    current_peak_power_min = data_model.peak_power_min_summer - adjusted_power_size
                else:  # 冬季
                    current_peak_power_min = data_model.peak_power_min_winter - adjusted_power_size
        
        # 获取平衡计算得到的电力负荷（考虑检修和投运计划修正后）
        electric_load = results['hourly_corrected_electric_load'][hour]
        
        # 成本优化策略
        # 1. 优先使用成本最低的电源
        # 2. 满足电力负荷不变的约束
        # 3. 满足各电源的出力约束
        
        # 排序各种电源的单位成本（从低到高）
        energy_sources = []
        
        # 光伏（边际成本最低，通常接近 0）
        if pv_output_orig > 0:
            energy_sources.append(('pv', current_pv_cost, pv_output_orig))
        
        # 风电（边际成本低）
        if wind_output_orig > 0:
            energy_sources.append(('wind', current_wind_cost, wind_output_orig))
        
        # 火电（成本相对较高）
        energy_sources.append(('thermal', current_thermal_cost, current_peak_power_max))
        
        # 下网负荷（当成本更低时考虑）
        current_grid_price = grid_price[hour] if hour < len(grid_price) else 0
        energy_sources.append(('grid', current_grid_price, float('inf')))  # 下网购电理论上无限制
        
        # 按成本排序
        energy_sources.sort(key=lambda x: x[1])
        
        # 初始化优化后的出力
        optimized_thermal = 0.0
        optimized_pv = 0.0
        optimized_wind = 0.0
        optimized_grid = 0.0
        
        # 需要满足的总负荷
        required_load = electric_load
        
        # 按照成本从低到高分配出力
        for source_type, cost, max_output in energy_sources:
            if required_load <= 0:
                break
            
            if source_type == 'pv':
                # 光伏出力
                actual_output = min(pv_output_orig, required_load)
                optimized_pv = actual_output
                required_load -= actual_output
                
            elif source_type == 'wind':
                # 风电出力
                actual_output = min(wind_output_orig, required_load)
                optimized_wind = actual_output
                required_load -= actual_output
                
            elif source_type == 'thermal':
                # 火电出力（在约束范围内）
                # 首先满足 CHP 出力
                thermal_needed = max(0, required_load - chp_output)
                
                # 调峰机组出力在最小和最大之间
                peak_output = max(current_peak_power_min, min(thermal_needed, current_peak_power_max))
                optimized_thermal = chp_output + peak_output
                required_load -= optimized_thermal
                
            elif source_type == 'grid':
                # 下网负荷（当其他电源不足时）
                # 但要满足最小下网负荷约束
                optimized_grid = max(required_load, min_grid_load)
                required_load -= optimized_grid
        
        # 如果还有剩余负荷未满足，增加火电出力
        if required_load > 0:
            additional_thermal = min(required_load, current_peak_power_max - optimized_thermal + chp_output)
            optimized_thermal += additional_thermal
            required_load -= additional_thermal
        
        # 确保下网负荷满足最小约束
        if optimized_grid < min_grid_load:
            # 需要从电网购买更多电力
            deficit = min_grid_load - optimized_grid
            optimized_grid = min_grid_load
            
            # 为了平衡，减少其他电源出力
            # 优先减少成本最高的电源
            if optimized_thermal > chp_output + current_peak_power_min:
                reduce_amount = min(deficit, optimized_thermal - chp_output - current_peak_power_min)
                optimized_thermal -= reduce_amount
                deficit -= reduce_amount
        
        # 计算总出力
        total_generation = optimized_thermal + optimized_pv + optimized_wind + optimized_grid
        
        # 计算每小时成本
        # 注意：下网负荷小于 0 时（上网），下网电价设为 0
        grid_cost = optimized_grid * (current_grid_price if optimized_grid >= 0 else 0)
        hourly_cost = (optimized_thermal * current_thermal_cost + 
                      optimized_pv * current_pv_cost + 
                      optimized_wind * current_wind_cost + 
                      grid_cost)
        
        # 存储优化结果
        optimized_results['hourly_thermal_output'][hour] = optimized_thermal
        optimized_results['hourly_pv_output'][hour] = optimized_pv
        optimized_results['hourly_wind_output'][hour] = optimized_wind
        optimized_results['hourly_grid_load'][hour] = optimized_grid
        optimized_results['hourly_total_generation'][hour] = total_generation
        optimized_results['hourly_cost'][hour] = hourly_cost
        
        # 更新进度
        if hour % 500 == 0:
            progress = (hour / 8760) * 100
            app_instance.optimization_progress["value"] = progress
            app_instance.optimization_progress_label.config(text=f"正在成本优化计算... {int(progress)}%")
            app_instance.root.update_idletasks()
    
    # 计算总成本
    total_cost = sum(optimized_results['hourly_cost'])
    optimized_results['total_cost'] = total_cost
    
    # 将优化结果存储到实例变量中
    app_instance.optimized_results = optimized_results
    app_instance.optimization_mode = 'cost_optimization'  # 标记为成本优化模式
    
    # 更新进度条到 100%
    app_instance.optimization_progress["value"] = 100
    app_instance.optimization_progress_label.config(text="成本优化计算完成！100%")
    app_instance.root.update_idletasks()
    
    # 计算优化前的各项数据（来自平衡计算结果）
    original_thermal = app_instance.results['hourly_thermal_output']
    original_pv = app_instance.results['hourly_pv_output']
    original_wind = app_instance.results['hourly_wind_output']
    original_grid = app_instance.results['hourly_grid_load']
    
    # 计算优化前的总出力
    original_total_generation = [original_thermal[i] + original_pv[i] + original_wind[i] for i in range(8760)]
    
    # 计算优化后的总出力
    optimized_total_generation = [optimized_results['hourly_thermal_output'][i] + 
                                 optimized_results['hourly_pv_output'][i] + 
                                 optimized_results['hourly_wind_output'][i] for i in range(8760)]
    
    # 计算优化前的成本（使用动态成本曲线）
    original_total_cost = 0.0
    for hour in range(8760):
        thermal_relative = original_thermal[hour] / thermal_max_output if thermal_max_output > 0 else 0
        thermal_cost = (data_model.thermal_cost_curve['quadratic_coefficient'] * thermal_relative**2 +
                       data_model.thermal_cost_curve['linear_coefficient'] * thermal_relative +
                       data_model.thermal_cost_curve['constant_term']) * data_model.thermal_cost_curve['base_cost']
        
        wind_relative = original_wind[hour] / total_wind_capacity if total_wind_capacity > 0 else 0
        wind_cost = (data_model.wind_cost_curve['linear_coefficient'] * wind_relative +
                    data_model.wind_cost_curve['constant_term']) * data_model.wind_cost_curve['base_cost']
        
        pv_relative = original_pv[hour] / total_pv_capacity if total_pv_capacity > 0 else 0
        pv_cost = (data_model.pv_cost_curve['linear_coefficient'] * pv_relative +
                  data_model.pv_cost_curve['constant_term']) * data_model.pv_cost_curve['base_cost']
        
        # 下网负荷小于 0 时，下网电价设为 0
        grid_cost = original_grid[hour] * (grid_price[hour] if original_grid[hour] >= 0 else 0)
        
        hourly_cost = (original_thermal[hour] * thermal_cost + 
                      original_pv[hour] * pv_cost + 
                      original_wind[hour] * wind_cost + 
                      grid_cost)
        original_total_cost += hourly_cost
    
    # 显示优化结果摘要 - 包含优化前后对比
    result_text = f"""成本优化计算完成!

优化目标：最小化发电成本

========== 优化前 ========== 
火电总出力：{sum(original_thermal):,.2f} kWh
风电总出力：{sum(original_wind):,.2f} kWh
光伏总出力：{sum(original_pv):,.2f} kWh
下网总出力：{sum(original_grid):,.2f} kWh
总出力：{sum(original_total_generation):,.2f} kWh
总成本：{original_total_cost:,.2f} 元
平均每小时成本：{original_total_cost/8760:.2f} 元

========== 优化后 ==========
火电总出力：{sum(optimized_results['hourly_thermal_output']):,.2f} kWh
风电总出力：{sum(optimized_results['hourly_wind_output']):,.2f} kWh
光伏总出力：{sum(optimized_results['hourly_pv_output']):,.2f} kWh
下网总出力：{sum(optimized_results['hourly_grid_load']):,.2f} kWh
总出力：{sum(optimized_results['hourly_total_generation']):,.2f} kWh
总成本：{total_cost:,.2f} 元
平均每小时成本：{total_cost/8760:.2f} 元

========== 优化效果 ==========
火电出力变化：{sum(optimized_results['hourly_thermal_output']) - sum(original_thermal):,.2f} kWh ({((sum(optimized_results['hourly_thermal_output']) - sum(original_thermal)) / sum(original_thermal) * 100) if sum(original_thermal) != 0 else 0:+.2f}%)
风电出力变化：{sum(optimized_results['hourly_wind_output']) - sum(original_wind):,.2f} kWh ({((sum(optimized_results['hourly_wind_output']) - sum(original_wind)) / sum(original_wind) * 100) if sum(original_wind) != 0 else 0:+.2f}%)
光伏出力变化：{sum(optimized_results['hourly_pv_output']) - sum(original_pv):,.2f} kWh ({((sum(optimized_results['hourly_pv_output']) - sum(original_pv)) / sum(original_pv) * 100) if sum(original_pv) != 0 else 0:+.2f}%)
下网负荷变化：{sum(optimized_results['hourly_grid_load']) - sum(original_grid):,.2f} kWh ({((sum(optimized_results['hourly_grid_load']) - sum(original_grid)) / sum(original_grid) * 100) if sum(original_grid) != 0 else 0:+.2f}%)
总出力变化：{sum(optimized_results['hourly_total_generation']) - sum(original_total_generation):,.2f} kWh
成本节约：{original_total_cost - total_cost:,.2f} 元 ({((original_total_cost - total_cost) / original_total_cost * 100) if original_total_cost != 0 else 0:+.2f}%)

说明:
- 优化目标为每小时发电成本最小化
- 优先使用低成本电源（光伏、风电）
- 满足电力负荷不变和各电源出力约束
- 下网负荷不小于最小下网负荷约束
- 下网负荷小于 0 时（上网），下网电价设为 0"""
    
    app_instance.optimization_result_text.delete(1.0, tk.END)
    app_instance.optimization_result_text.insert(tk.END, result_text)
    
    # 保存优化结果到当前项目
    app_instance.save_current_project()
    
    from tkinter import messagebox
    messagebox.showinfo("完成", "成本优化计算已完成！")


def export_cost_optimization_results(app_instance):
    """
    导出成本优化结果，包含优化前后的详细对比数据和汇总信息
    包括每小时优化结果和结果汇总（按月汇总和按年汇总）
    """
    if not hasattr(app_instance, 'optimized_results') or not app_instance.results:
        from tkinter import messagebox
        messagebox.showwarning("警告", "优化结果或平衡计算结果为空，无法导出！")
        return
    
    # 询问用户保存位置
    save_path = tk.filedialog.asksaveasfilename(
        title="保存成本优化结果",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile="cost_optimization_results.xlsx"
    )
    
    if save_path:
        try:
            import pandas as pd
            
            # 创建时间戳列表
            base_date = datetime(2025, 1, 1)
            time_stamps = [(base_date + timedelta(hours=i)).strftime("%Y-%m-%d %H:%M") for i in range(8760)]
            
            data_model = app_instance.data_model
            results = app_instance.results
            optimized_results = app_instance.optimized_results
            
            # 预计算参数
            total_wind_capacity = data_model.calculate_wind_total_capacity()
            total_pv_capacity = data_model.calculate_pv_total_capacity()
            thermal_max_output = data_model.peak_power_max + data_model.chp_electric_params['base_electric']
            grid_price = data_model.grid_purchase_price_hourly
            
            # 获取优化前的数据
            original_thermal = results['hourly_thermal_output']
            original_pv = results['hourly_pv_output']
            original_wind = results['hourly_wind_output']
            original_grid = results['hourly_grid_load']
            original_total_generation = [original_thermal[i] + original_pv[i] + original_wind[i] for i in range(8760)]
            
            # 获取优化后的数据
            optimized_thermal = optimized_results['hourly_thermal_output']
            optimized_pv = optimized_results['hourly_pv_output']
            optimized_wind = optimized_results['hourly_wind_output']
            optimized_grid = optimized_results['hourly_grid_load']
            optimized_total_generation = optimized_results['hourly_total_generation']
            optimized_cost = optimized_results['hourly_cost']
            
            # 计算优化前的每小时成本
            original_hourly_cost = []
            for i in range(8760):
                thermal_relative = original_thermal[i] / thermal_max_output if thermal_max_output > 0 else 0
                thermal_cost = (data_model.thermal_cost_curve['quadratic_coefficient'] * thermal_relative**2 +
                               data_model.thermal_cost_curve['linear_coefficient'] * thermal_relative +
                               data_model.thermal_cost_curve['constant_term']) * data_model.thermal_cost_curve['base_cost']
                
                wind_relative = original_wind[i] / total_wind_capacity if total_wind_capacity > 0 else 0
                wind_cost = (data_model.wind_cost_curve['linear_coefficient'] * wind_relative +
                            data_model.wind_cost_curve['constant_term']) * data_model.wind_cost_curve['base_cost']
                
                pv_relative = original_pv[i] / total_pv_capacity if total_pv_capacity > 0 else 0
                pv_cost = (data_model.pv_cost_curve['linear_coefficient'] * pv_relative +
                          data_model.pv_cost_curve['constant_term']) * data_model.pv_cost_curve['base_cost']
                
                grid_cost = original_grid[i] * (grid_price[i] if original_grid[i] >= 0 else 0)
                
                hourly_cost = (original_thermal[i] * thermal_cost + 
                              original_pv[i] * pv_cost + 
                              original_wind[i] * wind_cost + 
                              grid_cost)
                original_hourly_cost.append(hourly_cost)
            
            # 计算变化值
            thermal_diff = [optimized_thermal[i] - original_thermal[i] for i in range(8760)]
            wind_diff = [optimized_wind[i] - original_wind[i] for i in range(8760)]
            pv_diff = [optimized_pv[i] - original_pv[i] for i in range(8760)]
            grid_diff = [optimized_grid[i] - original_grid[i] for i in range(8760)]
            total_gen_diff = [optimized_total_generation[i] - original_total_generation[i] for i in range(8760)]
            cost_diff = [optimized_cost[i] - original_hourly_cost[i] for i in range(8760)]
            
            # 创建每小时详细对比 DataFrame
            df_detailed = pd.DataFrame({
                '时间': time_stamps,
                # 优化前数据
                '优化前火电出力 (kW)': original_thermal,
                '优化前风电出力 (kW)': original_wind,
                '优化前光伏出力 (kW)': original_pv,
                '优化前下网负荷 (kW)': original_grid,
                '优化前总出力 (kW)': original_total_generation,
                '优化前成本 (元)': original_hourly_cost,
                # 优化后数据
                '优化后火电出力 (kW)': optimized_thermal,
                '优化后风电出力 (kW)': optimized_wind,
                '优化后光伏出力 (kW)': optimized_pv,
                '优化后下网负荷 (kW)': optimized_grid,
                '优化后总出力 (kW)': optimized_total_generation,
                '优化后成本 (元)': optimized_cost,
                # 差异数据
                '火电出力差 (kW)': thermal_diff,
                '风电出力差 (kW)': wind_diff,
                '光伏出力差 (kW)': pv_diff,
                '下网负荷差 (kW)': grid_diff,
                '总出力差 (kW)': total_gen_diff,
                '成本差 (元)': cost_diff
            })
            
            # 按月度汇总
            monthly_summary = []
            for month in range(1, 13):
                # 计算该月的小时范围
                if month == 1:
                    start_hour = 0
                    end_hour = 744  # 31 天
                elif month in [3, 5, 7, 8, 10, 12]:
                    start_hour = sum([31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30][:month-1]) * 24
                    end_hour = start_hour + 31 * 24
                elif month in [4, 6, 9, 11]:
                    start_hour = sum([31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30][:month-1]) * 24
                    end_hour = start_hour + 30 * 24
                else:  # 2 月
                    start_hour = 31 * 24
                    end_hour = start_hour + 28 * 24
                
                # 确保不超出范围
                start_hour = max(0, min(start_hour, 8759))
                end_hour = max(0, min(end_hour, 8760))
                
                if start_hour < end_hour:
                    month_hours = list(range(start_hour, end_hour))
                    
                    # 计算月度汇总数据
                    orig_thermal = sum(original_thermal[i] for i in month_hours)
                    orig_wind = sum(original_wind[i] for i in month_hours)
                    orig_pv = sum(original_pv[i] for i in month_hours)
                    orig_grid = sum(original_grid[i] for i in month_hours)
                    orig_total = sum(original_total_generation[i] for i in month_hours)
                    orig_cost = sum(original_hourly_cost[i] for i in month_hours)
                    
                    opt_thermal = sum(optimized_thermal[i] for i in month_hours)
                    opt_wind = sum(optimized_wind[i] for i in month_hours)
                    opt_pv = sum(optimized_pv[i] for i in month_hours)
                    opt_grid = sum(optimized_grid[i] for i in month_hours)
                    opt_total = sum(optimized_total_generation[i] for i in month_hours)
                    opt_cost = sum(optimized_cost[i] for i in month_hours)
                    
                    cost_saving = orig_cost - opt_cost
                    
                    monthly_summary.append([
                        f'{month}月',
                        f'{orig_thermal:.2f}',
                        f'{orig_wind:.2f}',
                        f'{orig_pv:.2f}',
                        f'{orig_grid:.2f}',
                        f'{orig_total:.2f}',
                        f'{orig_cost:.2f}',
                        f'{opt_thermal:.2f}',
                        f'{opt_wind:.2f}',
                        f'{opt_pv:.2f}',
                        f'{opt_grid:.2f}',
                        f'{opt_total:.2f}',
                        f'{opt_cost:.2f}',
                        f'{opt_thermal - orig_thermal:.2f}',
                        f'{opt_wind - orig_wind:.2f}',
                        f'{opt_pv - orig_pv:.2f}',
                        f'{opt_grid - orig_grid:.2f}',
                        f'{opt_total - orig_total:.2f}',
                        f'{cost_saving:.2f}'
                    ])
            
            # 年度汇总
            annual_orig_thermal = sum(original_thermal)
            annual_orig_wind = sum(original_wind)
            annual_orig_pv = sum(original_pv)
            annual_orig_grid = sum(original_grid)
            annual_orig_total = sum(original_total_generation)
            annual_orig_cost = sum(original_hourly_cost)
            
            annual_opt_thermal = sum(optimized_thermal)
            annual_opt_wind = sum(optimized_wind)
            annual_opt_pv = sum(optimized_pv)
            annual_opt_grid = sum(optimized_grid)
            annual_opt_total = sum(optimized_total_generation)
            annual_opt_cost = sum(optimized_cost)
            
            annual_cost_saving = annual_orig_cost - annual_opt_cost
            
            monthly_summary.append([
                '年度总计',
                f'{annual_orig_thermal:.2f}',
                f'{annual_orig_wind:.2f}',
                f'{annual_orig_pv:.2f}',
                f'{annual_orig_grid:.2f}',
                f'{annual_orig_total:.2f}',
                f'{annual_orig_cost:.2f}',
                f'{annual_opt_thermal:.2f}',
                f'{annual_opt_wind:.2f}',
                f'{annual_opt_pv:.2f}',
                f'{annual_opt_grid:.2f}',
                f'{annual_opt_total:.2f}',
                f'{annual_opt_cost:.2f}',
                f'{annual_opt_thermal - annual_orig_thermal:.2f}',
                f'{annual_opt_wind - annual_orig_wind:.2f}',
                f'{annual_opt_pv - annual_orig_pv:.2f}',
                f'{annual_opt_grid - annual_orig_grid:.2f}',
                f'{annual_opt_total - annual_orig_total:.2f}',
                f'{annual_cost_saving:.2f}'
            ])
            
            # 创建汇总 DataFrame
            summary_df = pd.DataFrame(monthly_summary, columns=[
                '月份',
                '优化前火电总出力 (kWh)', '优化前风电总出力 (kWh)', '优化前光伏总出力 (kWh)',
                '优化前下网总出力 (kWh)', '优化前总出力 (kWh)', '优化前总成本 (元)',
                '优化后火电总出力 (kWh)', '优化后风电总出力 (kWh)', '优化后光伏总出力 (kWh)',
                '优化后下网总出力 (kWh)', '优化后总出力 (kWh)', '优化后总成本 (元)',
                '火电出力变化 (kWh)', '风电出力变化 (kWh)', '光伏出力变化 (kWh)',
                '下网负荷变化 (kWh)', '总出力变化 (kWh)', '成本节约 (元)'
            ])
            
            # 导出到 Excel
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                # 每小时详细对比数据
                df_detailed.to_excel(writer, sheet_name='每小时优化结果', index=False)
                
                # 汇总信息
                summary_df.to_excel(writer, sheet_name='结果汇总', index=False)
            
            from tkinter import messagebox
            messagebox.showinfo("成功", f"成本优化结果已导出至:\n{save_path}\n包含每小时详细对比和月度/年度汇总数据")
            
        except ImportError:
            # 如果 pandas 不可用，使用文本格式导出
            txt_save_path = save_path.replace('.xlsx', '.txt')
            with open(txt_save_path, 'w', encoding='utf-8') as f:
                f.write("成本优化结果对比\n\n")
                f.write(f"总成本 (优化前): {sum(original_hourly_cost):.2f} 元\n")
                f.write(f"总成本 (优化后): {sum(optimized_cost):.2f} 元\n")
                f.write(f"成本节约：{sum(original_hourly_cost) - sum(optimized_cost):.2f} 元\n\n")
                f.write("每小时优化结果对比 (前 10 小时示例):\n")
                f.write("小时，优化前火电 (kW),优化前风电 (kW),优化前光伏 (kW),优化前下网 (kW),优化后火电 (kW),优化后风电 (kW),优化后光伏 (kW),优化后下网 (kW),优化前成本 (元),优化后成本 (元),成本差 (元)\n")
                for i in range(min(10, 8760)):
                    f.write(f"{i},{original_thermal[i]:.2f},{original_wind[i]:.2f},{original_pv[i]:.2f},{original_grid[i]:.2f},"
                           f"{optimized_thermal[i]:.2f},{optimized_wind[i]:.2f},{optimized_pv[i]:.2f},{optimized_grid[i]:.2f},"
                           f"{original_hourly_cost[i]:.2f},{optimized_cost[i]:.2f},{optimized_cost[i]-original_hourly_cost[i]:.2f}\n")
            messagebox.showinfo("成功", f"成本优化结果已导出至:\n{txt_save_path} (由于缺少 pandas 库，以文本格式导出)")
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("错误", f"导出优化结果失败:\n{str(e)}")


# 如果需要作为独立函数调用
if __name__ == "__main__":
    print("此模块需要与主程序配合使用")
