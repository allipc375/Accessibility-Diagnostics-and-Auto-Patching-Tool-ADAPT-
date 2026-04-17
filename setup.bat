@echo off

echo ======================================
echo Setting up Accessibility Checker
echo ======================================

REM 1. Create virtual environment
echo Creating virtual environment...
python -m venv venv

REM 2. Activate it
echo Activating virtual environment...
call venv\Scripts\activate

REM 3. Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip

REM 4. Install PyTorch (CPU version)
echo Installing PyTorch (CPU)...
pip install torch torchvision --index-url https://download.pytorch.org/whl/cpu

REM 5. Install dependencies
echo Installing requirements...
pip install -r requirements.txt

echo ======================================
echo Setup complete!
echo To activate later: venv\Scripts\activate
echo ======================================

pause
