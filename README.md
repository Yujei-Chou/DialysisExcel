# DialysisExcel
Generate Peritoneal Dialysis Log
## Setting:
Create Google Form with the following sections:
   - 體重 (kg) / 簡答
   - 收縮壓 (mmHg) / 簡答
   - 舒張壓 (mmHg) / 簡答
   - 透析液濃度 (%) / 選擇題 (1.5 or 2.5)
   - 脫水量 (cc) / 簡答
## Get exec file:
```
pyinstaller app.py --onefile --noconsole
```
## Execute app:
open dist/app.exe
