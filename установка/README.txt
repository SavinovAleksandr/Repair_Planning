Папка офлайн-установки
======================

Содержимое:
  wheels/              — openpyxl 3.1.5 + et_xmlfile 2.0.0 (все нужные pip-пакеты)
  install_offline.bat  — установка БЕЗ интернета
  download_wheels.sh   — обновление wheels (на Mac/Linux, с интернетом)

На компьютере должны быть:
  • Python 3.10+  https://www.python.org/downloads/
  • галочки при установке: Add python.exe to PATH, tcl/tk (для окна GUI)

Дополнительных библиотек кроме openpyxl/et_xmlfile программа не требует.
tkinter идёт вместе с Python — отдельно не ставится.

Шаги (Windows, без интернета):
  1. Дважды кликнуть install_offline.bat
  2. Перейти в корень программы (папка с run_svod.bat)
  3. run_svod.bat --check
  4. run_svod.bat  — открыть окно программы

Альтернатива из корня программы:
  run_svod.bat --setup
  (сначала wheels из этой папки, потом интернет при неудаче)

CLI без окна:
  run_svod.bat --stage all
