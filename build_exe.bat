@echo off
chcp 65001 >nul
echo ========================================================
echo   DONG GOI IN HANG LOAT EXCEL PRO THANH FILE .EXE DOC LAP
echo ========================================================
echo.
echo [1/3] Dang cai dat cac thu vien can thiet...
pip install pyinstaller pywin32 tkinterdnd2

echo.
echo [2/3] Dang bien dich thanh file InHangLoatExcelPro.exe...
pyinstaller --noconfirm --onefile --windowed --name "InHangLoatExcelPro" --collect-all tkinterdnd2 main.py

echo.
echo [3/3] HOAN TAT!
echo File chay truc tiep nam trong thu muc: dist\InHangLoatExcelPro.exe
echo Ban co the copy file .exe nay sang bat ky may Windows nao (khong can Python).
pause
