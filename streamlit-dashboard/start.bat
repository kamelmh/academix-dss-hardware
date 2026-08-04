@echo off
echo ========================================
echo   Academix DSS - Dashboard Demo
echo ========================================
echo.
echo Starting server...
echo URL: http://localhost:8501
echo.
echo Press Ctrl+C to stop
echo ========================================
echo.

"C:\Users\Admin\AppData\Local\Python\pythoncore-3.14-64\Scripts\streamlit.exe" run "C:\Users\Admin\projects\active\academix-dss-hardware\streamlit-dashboard\app.py" --server.headless true --server.port 8501

pause
