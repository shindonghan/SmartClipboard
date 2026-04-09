@echo off
python -m PyInstaller --onefile --windowed --name SmartClipboard --hidden-import pystray._win32 --hidden-import PIL._imaging --hidden-import PIL.Image --hidden-import PIL.ImageDraw --hidden-import pyperclip --hidden-import keyboard --hidden-import pyautogui --collect-all pystray --collect-all PIL main.py
if exist "dist\SmartClipboard.exe" (
    echo BUILD OK: dist\SmartClipboard.exe
) else (
    echo BUILD FAILED
)
pause
