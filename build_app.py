import subprocess
import sys
import os

# 切换到项目目录
os.chdir(r'C:\R&D(Local)\EAM\SD\XJ')

# 运行PyInstaller命令
cmd = [
    sys.executable, '-m', 'PyInstaller',
    '--onefile',
    '--windowed', 
    '--name', 'EnergyBalanceSystem',
    'loadcalculation.py'
]

print("正在执行打包命令...")
print(' '.join(cmd))

try:
    result = subprocess.run(cmd, check=True, capture_output=True, text=True)
    print("打包成功完成!")
    print("标准输出:", result.stdout)
except subprocess.CalledProcessError as e:
    print(f"打包失败，错误代码: {e.returncode}")
    print("错误输出:", e.stderr)
    print("标准输出:", e.stdout)
except Exception as e:
    print(f"执行过程中发生错误: {e}")