@echo off
setlocal
REM (optionnel) activer un venv :
REM call %~dp0venv\Scripts\activate

python -m pip install -r "%~dp0requirements.txt"
start "" http://localhost:8501
python -m streamlit run "%~dp0app_streamlit.py" --server.port 8501 --server.headless true --global.developmentMode=false
