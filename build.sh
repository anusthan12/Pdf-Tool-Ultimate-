#!/usr/bin/env bash
# exit on error
set -o errexit

# Install system dependencies for Tesseract and Poppler
apt-get update
apt-get install -y tesseract-ocr poppler-utils

# Install Python dependencies
pip install -r requirements.txt
