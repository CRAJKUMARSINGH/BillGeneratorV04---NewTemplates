@echo off
set PORT=8501

echo Checking if port %PORT% is in use...
for /f "tokens=5" %%i in ('netstat -ano ^| findstr :%PORT%') do (
    echo Process ID using port %PORT% is %%i
    taskkill /pid %%i /f
    echo Process %%i killed.
)

echo Running Streamlit on port %PORT%...
streamlit run app.py --server.port=%PORT%