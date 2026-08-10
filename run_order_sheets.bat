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
set SCRIPT_BP_MINI=run_azs_bp_mini.py
set SCRIPT_PN=run_azs_pn.py
set SCRIPT_MP_FOOD=run_azs_mp_food.py
set SCRIPT_MP_NONFOOD=run_azs_mp_nonfood.py
set SCRIPT_AGNKS=run_agnks.py
set SCRIPT_SKY_MARKET=run_sky_market.py

:: ==========================================
:: ПУТИ К СКРИПТАМ ЗАГРУЗКИ
:: ==========================================
set DIR_UPLOAD=C:\Users\Пользователь\Desktop\Python Scripts\update_zayavochnik_for_azs
set UPLOAD_BP=%DIR_UPLOAD%\update_order_form_bp.py
set UPLOAD_BP_MINI=%DIR_UPLOAD%\update_order_form_bp_mini.py
set UPLOAD_PN=%DIR_UPLOAD%\update_order_form_pn.py
set UPLOAD_MP_FOOD=%DIR_UPLOAD%\update_order_form_mp_food.py
set UPLOAD_MP_NONFOOD=%DIR_UPLOAD%\update_order_form_mp_non-food.py
set UPLOAD_AGNKS=%DIR_UPLOAD%\update_order_form_agnks.py
set UPLOAD_SKY_MARKET=%DIR_UPLOAD%\update_order_form_sky_market.py

:menu
cls
echo =======================================================
echo       ПАНЕЛЬ УПРАВЛЕНИЯ ГЕНЕРАЦИЕЙ ЗАЯВОЧНИКОВ АЗС
echo =======================================================
echo.
echo [1] Bishkek Petroleum
echo [2] Bishkek Petroleum Mini
echo [3] PARTNER NEFT
echo [4] Мунай Пром Food
echo [5] Мунай Пром Non-Food
echo [6] Газинтерсервис
echo [7] Sky Market
echo [8] ВСЕ сети сразу
echo.
echo [0] Выход из программы
echo.
echo =======================================================
set /p choice="Выберите сеть (0-8) и нажмите Enter: "

if "%choice%"=="1" goto submenu_bp
if "%choice%"=="2" goto submenu_bp_mini
if "%choice%"=="3" goto submenu_pn
if "%choice%"=="4" goto submenu_mp_food
if "%choice%"=="5" goto submenu_mp_nonfood
if "%choice%"=="6" goto submenu_agnks
if "%choice%"=="7" goto submenu_sky_market
if "%choice%"=="8" goto submenu_all
if "%choice%"=="0" goto exit
goto menu

:: ==========================================
:: ПОДМЕНЮ ВЫБОРА ДЕЙСТВИЯ ДЛЯ ОДНОЙ СЕТИ
:: ==========================================

:submenu_bp
call :show_action_menu "Bishkek Petroleum"
set /p act="Выберите действие (0-3) и нажмите Enter: "
if "%act%"=="0" goto menu
if "%act%"=="1" call :run_action 1 "%SCRIPT_BP%" "%UPLOAD_BP%" "Bishkek Petroleum"
if "%act%"=="2" call :run_action 2 "%SCRIPT_BP%" "%UPLOAD_BP%" "Bishkek Petroleum"
if "%act%"=="3" call :run_action 3 "%SCRIPT_BP%" "%UPLOAD_BP%" "Bishkek Petroleum"
if not "%act%"=="0" if not "%act%"=="1" if not "%act%"=="2" if not "%act%"=="3" goto submenu_bp
pause
goto menu

:submenu_bp_mini
call :show_action_menu "Bishkek Petroleum Mini"
set /p act="Выберите действие (0-3) и нажмите Enter: "
if "%act%"=="0" goto menu
if "%act%"=="1" call :run_action 1 "%SCRIPT_BP_MINI%" "%UPLOAD_BP_MINI%" "Bishkek Petroleum Mini"
if "%act%"=="2" call :run_action 2 "%SCRIPT_BP_MINI%" "%UPLOAD_BP_MINI%" "Bishkek Petroleum Mini"
if "%act%"=="3" call :run_action 3 "%SCRIPT_BP_MINI%" "%UPLOAD_BP_MINI%" "Bishkek Petroleum Mini"
if not "%act%"=="0" if not "%act%"=="1" if not "%act%"=="2" if not "%act%"=="3" goto submenu_bp_mini
pause
goto menu

:submenu_pn
call :show_action_menu "PARTNER NEFT"
set /p act="Выберите действие (0-3) и нажмите Enter: "
if "%act%"=="0" goto menu
if "%act%"=="1" call :run_action 1 "%SCRIPT_PN%" "%UPLOAD_PN%" "PARTNER NEFT"
if "%act%"=="2" call :run_action 2 "%SCRIPT_PN%" "%UPLOAD_PN%" "PARTNER NEFT"
if "%act%"=="3" call :run_action 3 "%SCRIPT_PN%" "%UPLOAD_PN%" "PARTNER NEFT"
if not "%act%"=="0" if not "%act%"=="1" if not "%act%"=="2" if not "%act%"=="3" goto submenu_pn
pause
goto menu

:submenu_mp_food
call :show_action_menu "Мунай Пром Food"
set /p act="Выберите действие (0-3) и нажмите Enter: "
if "%act%"=="0" goto menu
if "%act%"=="1" call :run_action 1 "%SCRIPT_MP_FOOD%" "%UPLOAD_MP_FOOD%" "Мунай Пром Food"
if "%act%"=="2" call :run_action 2 "%SCRIPT_MP_FOOD%" "%UPLOAD_MP_FOOD%" "Мунай Пром Food"
if "%act%"=="3" call :run_action 3 "%SCRIPT_MP_FOOD%" "%UPLOAD_MP_FOOD%" "Мунай Пром Food"
if not "%act%"=="0" if not "%act%"=="1" if not "%act%"=="2" if not "%act%"=="3" goto submenu_mp_food
pause
goto menu

:submenu_mp_nonfood
call :show_action_menu "Мунай Пром Non-Food"
set /p act="Выберите действие (0-3) и нажмите Enter: "
if "%act%"=="0" goto menu
if "%act%"=="1" call :run_action 1 "%SCRIPT_MP_NONFOOD%" "%UPLOAD_MP_NONFOOD%" "Мунай Пром Non-Food"
if "%act%"=="2" call :run_action 2 "%SCRIPT_MP_NONFOOD%" "%UPLOAD_MP_NONFOOD%" "Мунай Пром Non-Food"
if "%act%"=="3" call :run_action 3 "%SCRIPT_MP_NONFOOD%" "%UPLOAD_MP_NONFOOD%" "Мунай Пром Non-Food"
if not "%act%"=="0" if not "%act%"=="1" if not "%act%"=="2" if not "%act%"=="3" goto submenu_mp_nonfood
pause
goto menu

:submenu_agnks
call :show_action_menu "Газинтерсервис"
set /p act="Выберите действие (0-3) и нажмите Enter: "
if "%act%"=="0" goto menu
if "%act%"=="1" call :run_action 1 "%SCRIPT_AGNKS%" "%UPLOAD_AGNKS%" "Газинтерсервис"
if "%act%"=="2" call :run_action 2 "%SCRIPT_AGNKS%" "%UPLOAD_AGNKS%" "Газинтерсервис"
if "%act%"=="3" call :run_action 3 "%SCRIPT_AGNKS%" "%UPLOAD_AGNKS%" "Газинтерсервис"
if not "%act%"=="0" if not "%act%"=="1" if not "%act%"=="2" if not "%act%"=="3" goto submenu_agnks
pause
goto menu

:submenu_sky_market
call :show_action_menu "Sky Market"
set /p act="Выберите действие (0-3) и нажмите Enter: "
if "%act%"=="0" goto menu
if "%act%"=="1" call :run_action 1 "%SCRIPT_SKY_MARKET%" "%UPLOAD_SKY_MARKET%" "Sky Market"
if "%act%"=="2" call :run_action 2 "%SCRIPT_SKY_MARKET%" "%UPLOAD_SKY_MARKET%" "Sky Market"
if "%act%"=="3" call :run_action 3 "%SCRIPT_SKY_MARKET%" "%UPLOAD_SKY_MARKET%" "Sky Market"
if not "%act%"=="0" if not "%act%"=="1" if not "%act%"=="2" if not "%act%"=="3" goto submenu_sky_market
pause
goto menu

:: ==========================================
:: ПОДМЕНЮ ВЫБОРА ДЕЙСТВИЯ ДЛЯ ВСЕХ СЕТЕЙ СРАЗУ
:: ==========================================

:submenu_all
call :show_action_menu "ВСЕ сети"
set /p act="Выберите действие (0-3) и нажмите Enter: "
if "%act%"=="0" goto menu
if not "%act%"=="1" if not "%act%"=="2" if not "%act%"=="3" goto submenu_all

cls
echo [СТАРТ] Применяем выбранное действие по очереди ко всем сетям...
echo -------------------------------------------------------
echo 1/7: Bishkek Petroleum...
call :run_action %act% "%SCRIPT_BP%" "%UPLOAD_BP%" "Bishkek Petroleum"
echo.
echo 2/7: Bishkek Petroleum Mini...
call :run_action %act% "%SCRIPT_BP_MINI%" "%UPLOAD_BP_MINI%" "Bishkek Petroleum Mini"
echo.
echo 3/7: PARTNER NEFT...
call :run_action %act% "%SCRIPT_PN%" "%UPLOAD_PN%" "PARTNER NEFT"
echo.
echo 4/7: Мунай Пром Food...
call :run_action %act% "%SCRIPT_MP_FOOD%" "%UPLOAD_MP_FOOD%" "Мунай Пром Food"
echo.
echo 5/7: Мунай Пром Non-Food...
call :run_action %act% "%SCRIPT_MP_NONFOOD%" "%UPLOAD_MP_NONFOOD%" "Мунай Пром Non-Food"
echo.
echo 6/7: Газинтерсервис...
call :run_action %act% "%SCRIPT_AGNKS%" "%UPLOAD_AGNKS%" "Газинтерсервис"
echo.
echo 7/7: Sky Market...
call :run_action %act% "%SCRIPT_SKY_MARKET%" "%UPLOAD_SKY_MARKET%" "Sky Market"
echo -------------------------------------------------------
echo [ГОТОВО] Действие выполнено для всех сетей!
pause
goto menu

:: ==========================================
:: ПОДПРОГРАММЫ (общие, без дублирования)
:: ==========================================

:show_action_menu
cls
echo =======================================================
echo       %~1
echo =======================================================
echo.
echo [1] Только генерация заявочника
echo [2] Только загрузка на Google Drive
echo [3] Генерация + загрузка
echo.
echo [0] Назад в главное меню
echo.
echo =======================================================
exit /b

:: %1 = действие (1=генерация, 2=загрузка, 3=оба)
:: %2 = путь к скрипту генерации
:: %3 = путь к скрипту загрузки
:: %4 = название сети для вывода в консоль
:run_action
if "%~1"=="1" (
    echo [ЗАПУСК] Генерация: %~4...
    "%PYTHON_BIN%" "%~2"
) else if "%~1"=="2" (
    echo [ЗАПУСК] Загрузка на Google Drive: %~4...
    "%PYTHON_BIN_UPLOAD%" "%~3"
) else if "%~1"=="3" (
    echo [ЗАПУСК] Генерация: %~4...
    "%PYTHON_BIN%" "%~2"
    echo.
    echo [ЗАПУСК] Загрузка на Google Drive: %~4...
    "%PYTHON_BIN_UPLOAD%" "%~3"
)
exit /b

:exit
echo Выход...
timeout /t 1 > nul
exit
