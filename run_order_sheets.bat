@echo off
chcp 65001 > nul
title Управление генерацией заявочников АЗС

:: Путь к интерпретатору Python (если он в PATH, можно оставить просто python)
set PYTHON_BIN=C:\Users\Пользователь\Desktop\Python Scripts\zayavochnik_generator\.venv\Scripts\python.exe

:: Пути к твоим скриптам (если батник лежит в той же папке, можно оставить только имена файлов)
set SCRIPT_BP="run_azs_bp.py"
set SCRIPT_PN="run_azs_pn.py"
set SCRIPT_MP_FOOD="run_azs_mp_food.py"
set SCRIPT_MP_NONFOOD="run_azs_mp_nonfood.py"

:menu
cls
echo =======================================================
echo       ПАНЕЛЬ УПРАВЛЕНИЯ ГЕНЕРАЦИЕЙ ЗАЯВОЧНИКОВ АЗС
echo =======================================================
echo.
echo [1] Запустить ВСЕ скрипты по очереди
echo [2] Запустить Bishkek Petroleum (%SCRIPT_BP%)
echo [3] Запустить PARTNER NEFT (%SCRIPT_PN%)
echo [4] Запустить Мунай Пром Food (%SCRIPT_MP_FOOD%)
echo [5] Запустить Мунай Пром Non-Food (%SCRIPT_MP_NONFOOD%)
echo.
echo [0] Выход из программы
echo.
echo =======================================================
set /p choice="Выберите вариант (0-5) и нажмите Enter: "

if "%choice%"=="1" goto run_all
if "%choice%"=="2" goto run_bp
if "%choice%"=="3" goto run_pn
if "%choice%"=="4" goto run_mp_food
if "%choice%"=="5" goto run_mp_nonfood
if "%choice%"=="0" goto exit
goto menu

:run_all
cls
echo [СТАРТ] Запуск всех конфигураций...
echo -------------------------------------------------------
echo 1/4 Переработка: Bishkek Petroleum...
%PYTHON_BIN% %SCRIPT_BP%
echo.
echo 2/4 Переработка: PARTNER NEFT...
%PYTHON_BIN% %SCRIPT_PN%
echo.
echo 3/4 Переработка: Мунай Пром Food...
%PYTHON_BIN% %SCRIPT_MP_FOOD%
echo.
echo 4/4 Переработка: Мунай Пром Non-Food...
%PYTHON_BIN% %SCRIPT_MP_NONFOOD%
echo -------------------------------------------------------
echo [ГОТОВО] Все заявочники успешно сгенерированы!
pause
goto menu

:run_bp
cls
echo [ЗАПУСК] Bishkek Petroleum...
%PYTHON_BIN% %SCRIPT_BP%
echo.
pause
goto menu

:run_pn
cls
echo [ЗАПУСК] PARTNER NEFT...
%PYTHON_BIN% %SCRIPT_PN%
echo.
pause
goto menu

:run_mp_food
cls
echo [ЗАПУСК] Мунай Пром Food...
%PYTHON_BIN% %SCRIPT_MP_FOOD%
echo.
pause
goto menu

:run_mp_nonfood
cls
echo [ЗАПУСК] Мунай Пром Non-Food...
%PYTHON_BIN% %SCRIPT_MP_NONFOOD%
echo.
pause
goto menu

:exit
echo Выход...
timeout /t 1 > nul
exit