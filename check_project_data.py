import json

with open('C:/R&D(Local)/EAM/SD/XJ/projects/20251211_232815/project_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)
    
    print('风机型号数据:')
    for i, model in enumerate(data.get('wind_turbine_models', [])):
        print(f'  风机型号 {i+1}: name={model["name"]}, output_correction_factor={model.get("output_correction_factor", "NOT FOUND")}')
        
    print('\n光伏型号数据:')
    for i, model in enumerate(data.get('pv_models', [])):
        print(f'  光伏型号 {i+1}: name={model["name"]}, output_correction_factor={model.get("output_correction_factor", "NOT FOUND")}')