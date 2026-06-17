@echo off
chcp 65001 > nul
title Управление генерацией и загрузкой заявочников АЗС

:: ==========================================
:: ПУТИ К ИНТЕРПРЕТАТОРАМ PYTHON (без кавычек при set)
:: ==========================================
:: 1. Окружение для генерации заявочников
set PYTHON_BIN=C:\Users\Пользователь\Desktop\Python Scripts\zayavochnik_generator\.venv\Scripts\python.exe

:: 2. Окружение для загрузки файлов
set PYTHON_BIN_UPLOAD=C:\Users\Пользователь\Desktop\Python Scripts\update_zayavochnik_for_azs\.venv\Scripts\python.exe

:: ==========================================
:: ПУТИ К СКРИПТАМ ГЕНЕРАЦИИ (В текущей папке батника)
:: ==========================================
set SCRIPT_BP=run_azs_bp.py
set SCRIPT_PN=run_azs_pn.py
set SCRIPT_MP_FOOD=run_azs_mp_food.py
set SCRIPT_MP_NONFOOD=run_azs_mp_nonfood.py

:: ==========================================
:: ПУТИ К СКРИПТАМ ЗАГРУЗКИ
:: ==========================================
set DIR_UPLOAD=C:\Users\Пользователь\Desktop\Python Scripts\update_zayavochnik_for_azs
set UPLOAD_BP=%DIR_UPLOAD%\update_order_form_bp.py
set UPLOAD_PN=%DIR_UPLOAD%\update_order_form_pn.py
set UPLOAD_MP_FOOD=%DIR_UPLOAD%\update_order_form_mp_food.py
set UPLOAD_MP_NONFOOD=%DIR_UPLOAD%\update_order_form_mp_non-food.py

:menu
cls
echo =======================================================
echo       ПАНЕЛЬ УПРАВЛЕНИЯ ГЕНЕРАЦИЕЙ ЗАЯВОЧНИКОВ АЗС
echo =======================================================
echo.
echo [1] Запустить ВСЕ скрипты (Генерация + Загрузка) по очереди
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
echo [СТАРТ] Запуск всех конфигураций генерации и загрузки...
echo -------------------------------------------------------
echo 1/4 Переработка и загрузка: Bishkek Petroleum...
"%PYTHON_BIN%" "%SCRIPT_BP%"
echo.
"%PYTHON_BIN_UPLOAD%" "%UPLOAD_BP%"
echo.
echo.
echo 2/4 Переработка и загрузка: PARTNER NEFT...
"%PYTHON_BIN%" "%SCRIPT_PN%"
echo.
"%PYTHON_BIN_UPLOAD%" "%UPLOAD_PN%"
echo.
echo.
echo 3/4 Переработка и загрузка: Мунай Пром Food...
"%PYTHON_BIN%" "%SCRIPT_MP_FOOD%"
echo.
"%PYTHON_BIN_UPLOAD%" "%UPLOAD_MP_FOOD%"
echo.
echo.
echo 4/4 Переработка и загрузка: Мунай Пром Non-Food...
"%PYTHON_BIN%" "%SCRIPT_MP_NONFOOD%"
echo.
"%PYTHON_BIN_UPLOAD%" "%UPLOAD_MP_NONFOOD%"
echo -------------------------------------------------------
echo [ГОТОВО] Все заявочники успешно сгенерированы и загружены!
pause
goto menu

:run_bp
cls
echo [ЗАПУСК] Генерация Bishkek Petroleum...
"%PYTHON_BIN%" "%SCRIPT_BP%"
echo.
echo.
echo [ЗАПУСК] Загрузка на Google Drive...
"%PYTHON_BIN_UPLOAD%" "%UPLOAD_BP%"
pause
goto menu

:run_pn
cls
echo [ЗАПУСК] Генерация PARTNER NEFT...
"%PYTHON_BIN%" "%SCRIPT_PN%"
echo.
echo.
echo [ЗАПУСК] Загрузка на Google Drive...
"%PYTHON_BIN_UPLOAD%" "%UPLOAD_PN%"
pause
goto menu

:run_mp_food
cls
echo [ЗАПУСК] Генерация Мунай Пром Food...
"%PYTHON_BIN%" "%SCRIPT_MP_FOOD%"
echo.
echo.
echo [ЗАПУСК] Загрузка на Google Drive...
"%PYTHON_BIN_UPLOAD%" "%UPLOAD_MP_FOOD%"
pause
goto menu

:run_mp_nonfood
cls
echo [ЗАПУСК] Генерация Мунай Пром Non-Food...
"%PYTHON_BIN%" "%SCRIPT_MP_NONFOOD%"
echo.
echo.
echo [ЗАПУСК] Загрузка на Google Drive...
"%PYTHON_BIN_UPLOAD%" "%UPLOAD_MP_NONFOOD%"
pause
goto menu

:exit
echo Выход...
timeout /t 1 > nul
exit