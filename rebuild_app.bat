@echo off
cd /d "C:\R&D(Local)\EAM\SD\XJ"
python -m PyInstaller --onefile --windowed --name EnergyBalanceSystem loadcalculation.py
pause