#!/usr/bin/env bash

echo "======================================"
echo "Setting up Accessibility Checker البيئة"
echo "======================================"

# Exit on error
set -e

# 1. Create virtual environment
echo "Creating virtual environment..."
python3 -m venv venv

# 2. Activate it
echo "Activating virtual environment..."
source venv/bin/activate

# 3. Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip

# 4. Install PyTorch (CPU version for compatibility)
echo "Installing PyTorch (CPU)..."
pip install torch torchvision --index-url https://download.pytorch.org/whl/cpu

# 5. Install remaining dependencies
echo "Installing requirements..."
pip install -r requirements.txt

echo "======================================"
echo "Setup complete!"
echo "Activate with: source venv/bin/activate"
echo "======================================"
