Offline install folder
======================

Contents:
  wheels/              - Python packages (openpyxl + et_xmlfile)
  install_offline.bat  - install without internet

Requirements on PC:
  - Python 3.10+ from https://www.python.org/downloads/
  - During install: enable "Add python.exe to PATH"
  - tkinter is part of standard Python (needed for GUI window)

Steps:
  1. Double-click install_offline.bat
  2. Go to project folder and run run_svod.bat --check

CLI without GUI:
  run_svod.bat --stage all

Refresh wheels (developer, with internet):
  bash установка/download_wheels.sh
