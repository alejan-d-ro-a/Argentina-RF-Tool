@echo off
echo Compilando Argentina RF Tool...
echo.

pyinstaller --noconfirm --onefile --windowed --icon="argentina.ico" --name="Argentina RF Tool" --add-data="argentina.ico;." --add-data="Cluster Final Report_PRE_vs_POST TEMPLATE.xlsx;." main.py

echo.
echo =========================================
echo Compilacion terminada exitosamente.
echo Revisa la carpeta "dist".
echo =========================================
pause