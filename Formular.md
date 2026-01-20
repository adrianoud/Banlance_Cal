# 能源平衡计算系统 - 完整公式文档

## 项目概述

XJ 能源平衡计算系统是一个园区用电用热负荷与出力平衡计算系统，主要功能包括：
- 项目管理（创建、打开、删除项目）
- 数据导入（电力负荷、热力负荷、光照强度、风速）
- 风机和光伏型号管理
- 检修和投产计划设置
- 年度平衡计算
- 结果可视化和导出

## 主要计算公式

### 1. 风机出力计算

#### 单台风机出力计算
```
风机出力计算函数根据风速分段计算：

1) 当风速 < 切入风速 或 风速 > 切出风速时：
   出力 = 0.0

2) 当 切入风速 ≤ 风速 < 额定风速时（二次曲线拟合）：
   a = 额定功率 / ((额定风速 - 切入风速)²)
   出力 = a × (风速 - 切入风速)² × 修正系数

3) 当 额定风速 ≤ 风速 < 最大额定风速时：
   出力 = 额定功率 × 修正系数

4) 当 最大额定风速 ≤ 风速 < 切出风速时（线性递减）：
   斜率 = 额定功率 / (切出风速 - 最大额定风速)
   出力 = (额定功率 - 斜率 × (风速 - 最大额定风速)) × 修正系数
```

#### 风机总出力计算
```
风机总装机容量 = Σ(各型号风机额定功率 × 该型号风机数量)
总出力 = Σ(单台风机出力 × 该型号风机数量)
```

### 2. 光伏出力计算

#### 光伏出力计算（支持两种方法）

**方法一：面积效率法**
```
光伏出力 = 光伏板面积 × 光照强度 × 光伏板效率 / 1000.0 × 修正系数
```

**方法二：装机容量法**
```
光伏出力 = 光照强度 / 1000.0 × 装机容量 × 系统效率 × 修正系数
```

#### 光伏总出力计算
```
光伏总装机容量 = Σ(各型号光伏装机容量 × 该型号光伏数量)
总出力 = Σ(单个光伏型号出力 × 该型号光伏数量)
```

### 3. 热电联产出力计算

```
热电联产电出力 = 基础电出力 + 热力负荷 × 电热比
```

## 年度平衡计算公式

在年度8760小时的能源平衡计算中，按照以下公式序列计算：

### 1) 修正后电力负荷
```
修正后电力负荷 = 原始电力负荷 / 最大电力负荷 × (最大电力负荷 - 检修影响 - 投运影响)
```

### 2) 厂用电负荷（迭代计算）
```
厂用电负荷 = 火电出力 × 厂用电率
```

### 3) 总负荷
```
总负荷 = 修正后电力负荷 + 厂用电负荷
```

### 4) 热定电机组出力
```
热定电机组出力 = 热力负荷 × 电热比
```

### 5) 光伏最大出力
```
光伏最大出力 = 根据所有光伏型号计算的总出力
```

### 6) 风机最大出力
```
风机最大出力 = 根据所有风机型号计算的总出力
```

### 7) 调峰机组待定出力
```
调峰机组待定出力 = 总负荷 - 热定电机组出力 - 光伏最大出力 - 风机最大出力
```

### 8) 调峰机组出力
```
调峰机组出力 = max(min(调峰机组待定出力, 调峰机组最大出力), 调峰机组最小出力)
```
注：调峰机组最小出力根据月份不同而变化，夏季（5-9月）和冬季（10-12月和1-4月）使用不同的最小出力值。

### 9) 火电出力
```
火电出力 = 热定电机组出力 + 调峰机组出力
```

### 10) 风机光伏放弃出力
```
风机光伏放弃出力 = max(当前月份对应的调峰机组最小出力 - 调峰机组待定出力, 0)
```

### 11) 灵活负荷消纳
```
if 风机光伏放弃出力 < 最小灵活负荷:
    灵活负荷消纳量 = 0
elif 最小灵活负荷 ≤ 风机光伏放弃出力 ≤ 最大灵活负荷:
    灵活负荷消纳量 = 风机光伏放弃出力
else:  # 风机光伏放弃出力 > 最大灵活负荷
    灵活负荷消纳量 = 最大灵活负荷
```

### 12) 修正后风机光伏放弃出力
```
修正后风机光伏放弃出力 = max(原风机光伏放弃出力 - 新增的灵活负荷消纳出力, 0)
```

### 13) 风机光伏实际出力
```
风机光伏实际出力 = (光伏最大出力 + 风机最大出力) - 修正后风机光伏放弃出力
```

### 14) 总出力
```
总出力 = 光伏出力 + 风机出力 + 火电出力 - 修正后风机光伏放弃出力
```

### 15) 弃光风率
```
弃光风率 = 修正后风机光伏放弃出力 / (光伏最大出力 + 风机最大出力)
```
注：当分母为0时，弃光风率为0

### 16) 下网负荷
```
下网负荷 = 总负荷 + 新增的灵活负荷消纳出力 - 总出力
```

## 投产计划影响计算方法

### 概述

投产计划是指在特定时间段内，某些设备或系统逐步投入运行的过程。系统通过线性插值的方式计算投产计划在不同时间点对系统参数的影响。

### 核心计算原理

#### 1. 线性插值因子计算

线性插值因子是投产计划计算的基础，决定了在不同时间点的影响程度：

```
def calculate_interpolation_factor(hour, start_date_str, end_date_str):
    """
    计算线性插值因子
    :param hour: 小时索引 (0-8759)
    :param start_date_str: 起始日期字符串 (YYYY-MM-DD)
    :param end_date_str: 结束日期字符串 (YYYY-MM-DD)
    :return: 插值因子 (0.0-1.0)
    """
    
    # 计算当前日期
    base_date = datetime(2024, 1, 1)
    current_date = base_date + timedelta(hours=hour)
    
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
```

**插值因子含义：**
- 投产前（开始日期之前）：插值因子 = 0.0
- 投产后（结束日期之后）：插值因子 = 1.0
- 投产期间（开始日期至结束日期之间）：插值因子 = 已过天数 / 总天数，呈线性增长

### 投产计划类型及其计算方法

#### 1. 光伏出力影响

```
# 在投产开始日期前，其最大出力修正为原最大出力减去设置的影响出力负荷大小
# 在投产结束日后，最大出力即为原最大出力
# 在投产期间，采用线性变化

if power_type == '光伏出力':
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
```

**计算逻辑：**
- 投产前：光伏出力 = (原总装机容量 - 影响负荷) / 原总装机容量 × 原出力
- 投产中：光伏出力 = (原总装机容量 - 影响负荷 × (1 - 插值因子)) / 原总装机容量 × 原出力
- 投产后：光伏出力 = 原出力（不受影响）

#### 2. 风机出力影响

```
# 在投产开始日期前，其最大出力修正为原最大出力减去设置的影响出力负荷大小
# 在投产结束日后，最大出力即为原最大出力
# 在投产期间，采用线性变化

if power_type == '风机出力':
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
```

**计算逻辑：**
- 与光伏出力影响类似，采用相同的线性插值方法

#### 3. 调峰机组最大出力影响

```
if power_type == '调峰机组最大出力':
    # 在投产开始日期前，修正的调峰机组最大负荷为原调峰机组最大负荷减去设置的影响出力负荷大小
    # 在投产结束日后，修正的调峰机组最大负荷即为原调峰机组最大负荷
    # 在投产期间，采用线性变化。即最大负荷随着投产进行逐步增大
    adjusted_power_size = power_size * interpolation_factor
    current_peak_power_max = current_peak_power_max - (power_size - adjusted_power_size)
```

**计算逻辑：**
- 投产前：最大出力 = 原最大出力 - 影响负荷
- 投产中：最大出力 = 原最大出力 - 影响负荷 × (1 - 插值因子)
- 投产后：最大出力 = 原最大出力

#### 4. 调峰机组最小出力影响

```
# 支持多种最小出力类型：夏季最小出力、冬季最小出力、通用最小出力

# 夏季最小出力影响（5-9月）
elif power_type == '调峰机组夏季最小出力':
    adjusted_power_size = power_size * interpolation_factor
    if 5 <= current_date.month <= 9:  # 夏季：5-9月
        current_peak_power_min = original_peak_power_min_summer - adjusted_power_size

# 冬季最小出力影响（10-12月和1-4月）
elif power_type == '调峰机组冬季最小出力':
    adjusted_power_size = power_size * interpolation_factor
    if current_date.month < 5 or current_date.month > 9:  # 冬季：10-12月和1-4月
        current_peak_power_min = original_peak_power_min_winter - adjusted_power_size

# 通用最小出力影响（兼容旧版本）
elif power_type == '调峰机组最小出力':
    adjusted_power_size = power_size * interpolation_factor
    if 5 <= current_date.month <= 9:  # 夏季
        current_peak_power_min = original_peak_power_min_summer - adjusted_power_size
    else:  # 冬季
        current_peak_power_min = original_peak_power_min_winter - adjusted_power_size
```

**计算逻辑：**
- 投产前：最小出力 = 原最小出力
- 投产中：最小出力 = 原最小出力 - 影响负荷 × 插值因子
- 投产后：最小出力 = 原最小出力 - 影响负荷

#### 5. 用电负荷影响

```
if power_type == '用电负荷':
    # 投运计划对用电负荷的影响是渐变的：开始前是全部影响，结束后是无影响，中间线性变化
    adjusted_power_size = power_size * (1 - interpolation_factor)  # 1-interpolation_factor 表示剩余影响
    commissioning_load_reduction += adjusted_power_size
```

**计算逻辑：**
- 投产前：用电负荷影响 = 全部影响（100%）
- 投产中：用电负荷影响 = 原影响 × (1 - 插值因子)
- 投产后：用电负荷影响 = 无影响（0%）

### 投产计划生效范围

```
def get_active_commissioning_schedules(self, hour):
    """
    获取在指定小时处于活动状态的投产计划
    对于投产计划，不仅在计划日期范围内需要激活，在计划开始日期之前也需要激活（用于投产前修正）
    """
    # 检查当前日期是否在投产计划的起止日期范围内，或者在起始日期之前
    # 如果在开始日期之前，也要激活计划（用于投产前修正）
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    if current_date <= start_date or self.is_date_in_range(current_date_str, start_date_str, end_date_str):
        active_schedules.append(schedule)
```

**生效范围：**
- 投产前：从开始日期之前到开始日期当天
- 投产期间：从开始日期到结束日期（包含首尾）
- 适用于需要提前影响系统参数的场景

### 多计划叠加处理

当存在多个投产计划影响同一参数时，系统会进行叠加处理：

```
# 收集所有光伏出力相关的投产计划
pv_schedules = []
for schedule in active_commissioning_schedules:
    if schedule.get('power_type', '') == '光伏出力':
        pv_schedules.append(schedule)

# 计算累计影响
cumulative_impact = 0.0
for schedule in pv_schedules:
    # ... 计算每个计划的影响并累加
    cumulative_impact += impact
```

**叠加原则：**
- 同类影响按线性叠加
- 不同类影响独立计算后合并
- 确保不会出现负值等不合理结果

### 计算精度保证

- 使用10次迭代确保计算收敛
- 设置收敛阈值为1e-6
- 对除零情况进行特殊处理
- 对日期解析失败进行异常捕获

### 实际应用示例

假设某光伏设备投产计划如下：
- 开始日期：2024-06-01
- 结束日期：2024-08-31
- 影响负荷：1000 kW

在不同时间点的计算结果：
- 2024-05-31（投产前）：光伏出力 = (总容量 - 1000) / 总容量 × 原出力
- 2024-07-15（投产中，插值因子=0.5）：光伏出力 = (总容量 - 1000×0.5) / 总容量 × 原出力
- 2024-09-01（投产后）：光伏出力 = 原出力

## 系统特性

### 检修和投产计划影响
- 检修计划会降低设备的最大出力能力
- 投产计划会在特定时间段内逐步增加设备出力
- 投产计划采用线性插值方法计算时间相关的出力影响

### 出力限制计划
- 可设置光伏和风机的最大出力限制
- 多个限制计划取最小限制值

### 灵活负荷机制
- 灵活负荷在风机光伏放弃出力超过最小灵活负荷时启动
- 最大消纳能力不超过最大灵活负荷
- 有助于减少新能源弃电现象

## 数据格式

数据导入使用[data_template.csv](file:///C:/R&D(Local)/EAM/SD/XJ/data_template.csv)格式，包含以下字段：
- 时间：时间戳（格式：YYYY-MM-DD HH:MM）
- 电力负荷(kW)：每小时电力负荷
- 热力负荷(kW)：每小时热力负荷
- 光照强度(W/m²)：每小时光照强度
- 风速(m/s)：每小时风速

## 技术栈

- Python 3.x
- Tkinter（图形界面）
- Matplotlib（绘图库）
- NumPy（数值计算）
- OpenPyXL（Excel文件处理）
- CSV、JSON、OS（标准库）

## 输出结果

系统计算结果包含8760小时（一年）的以下数据：
- 每小时厂用电负荷
- 每小时总负荷
- 每小时热定电机组出力
- 每小时光伏出力
- 每小时风机出力
- 每小时调峰机组待定出力
- 每小时调峰机组出力
- 每小时火电出力
- 每小时总出力
- 每小时风机光伏放弃出力
- 每小时风机光伏实际出力
- 每小时下网负荷
- 每小时弃光风率
- 每小时修正后电力负荷
- 每小时灵活负荷消纳量