@echo off
echo ================================================
echo   BUILD IngresoMasivo.exe - Xpress El Salvador
echo ================================================
echo.

:: Verificar que Python esta instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no encontrado. Instala Python 3.x primero.
    pause
    exit /b 1
)

:: Instalar PyInstaller si no esta
echo Verificando PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Instalando PyInstaller...
    pip install pyinstaller
)

:: Instalar dependencias del proyecto
echo Instalando dependencias...
pip install openpyxl pillow

echo.
echo Generando exe...
echo.

:: Generar el exe
:: --onedir: carpeta con exe + deps (mas estable que onefile para este proyecto)
:: --windowed: sin ventana de consola
:: --icon: icono del exe
:: --name: nombre del ejecutable
pyinstaller ^
    --onedir ^
    --windowed ^
    --icon=xpress_icon.png ^
    --name=IngresoMasivo ^
    --add-data "xpress_logo.png;." ^
    --add-data "xpress_icon.png;." ^
    --hidden-import=openpyxl ^
    --hidden-import=PIL ^
    --hidden-import=PIL.Image ^
    --hidden-import=PIL.ImageTk ^
    --noconfirm ^
    Ingreso_Masivo_XPES.pyw

if errorlevel 1 (
    echo.
    echo ERROR: Fallo la generacion del exe.
    pause
    exit /b 1
)

echo.
echo ================================================
echo   COPIANDO archivos necesarios al dist...
echo ================================================

:: Copiar archivos .py sueltos que el exe necesita llamar
set DIST=dist\IngresoMasivo

copy main_local.py        "%DIST%\" >nul
copy logica_local.py      "%DIST%\" >nul
copy indexar.py           "%DIST%\" >nul
copy servicios_variantes.py "%DIST%\" >nul
copy config_local.py      "%DIST%\" >nul
copy test.py              "%DIST%\" >nul
copy version.json         "%DIST%\" >nul

echo.
echo ================================================
echo   LISTO
echo ================================================
echo.
echo El ejecutable esta en:  dist\IngresoMasivo\IngresoMasivo.exe
echo.
echo Distribuye TODA la carpeta dist\IngresoMasivo\
echo (no solo el .exe, necesita los .py y recursos)
echo.
pause
