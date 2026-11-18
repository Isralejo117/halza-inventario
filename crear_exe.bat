@echo off
REM ==================================================================
REM  Crea el programa para instalar en otras PCs (no necesitan Python).
REM  Resultado: la carpeta dist\HalzaInventario
REM
REM  IMPORTANTE: no uses el programa desde dist\HalzaInventario.
REM  Cada vez que corres este archivo, esa carpeta se borra y se crea
REM  de nuevo (con todo y la base de datos que tuviera adentro).
REM  Copia la carpeta a otro lugar (ej. C:\HalzaInventario) para usarla.
REM ==================================================================

python -m PyInstaller --noconfirm --onedir --windowed --name HalzaInventario --collect-data customtkinter interfaz.py

echo.
echo Listo. El programa quedo en: dist\HalzaInventario
pause
