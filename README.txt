SPT Manufacturing Shortage Dashboard

Quick start:
1. Extract this ZIP completely.
2. Open the extracted folder.
3. Double-click start_dashboard.bat.
4. Your browser will open the Streamlit dashboard.

Included files:
- app.py
- requirements.txt
- start_dashboard.bat
- data/sample_data.xlsx

Main logic:
- Extract project code from remarks, e.g. 22M0026-01 -> 22M0026.
- Group the same material number under the same project.
- Sum shortage quantity and combine work-order / remarks lists.
- Allow project-name mapping and Excel export.
