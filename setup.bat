@echo off
echo Setting up MujahidPatient Project...
echo.

echo Creating virtual environment...
python -m venv venv

echo Activating virtual environment...
call venv\Scripts\activate.bat

echo Installing dependencies...
pip install --upgrade pip
pip install fastapi uvicorn pydantic python-multipart jinja2 aiofiles

echo.
echo Setup complete! To run the project:
echo 1. venv\Scripts\activate
echo 2. python run.py
echo.
echo Or just double-click run.bat

pause