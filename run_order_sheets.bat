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
echo [1] Bishkek Petroleum
echo [2] PARTNER NEFT
echo [3] Мунай Пром Food
echo [4] Мунай Пром Non-Food
echo [5] ВСЕ сети АЗС сразу
echo.
echo [0] Выход из программы
echo.
echo =======================================================
set /p choice="Выберите сеть (0-5) и нажмите Enter: "

if "%choice%"=="1" goto submenu_bp
if "%choice%"=="2" goto submenu_pn
if "%choice%"=="3" goto submenu_mp_food
if "%choice%"=="4" goto submenu_mp_nonfood
if "%choice%"=="5" goto submenu_all
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

:: ==========================================
:: ПОДМЕНЮ ВЫБОРА ДЕЙСТВИЯ ДЛЯ ВСЕХ СЕТЕЙ СРАЗУ
:: ==========================================

:submenu_all
call :show_action_menu "ВСЕ сети АЗС"
set /p act="Выберите действие (0-3) и нажмите Enter: "
if "%act%"=="0" goto menu
if not "%act%"=="1" if not "%act%"=="2" if not "%act%"=="3" goto submenu_all

cls
echo [СТАРТ] Применяем выбранное действие по очереди ко всем сетям...
echo -------------------------------------------------------
echo 1/4: Bishkek Petroleum...
call :run_action %act% "%SCRIPT_BP%" "%UPLOAD_BP%" "Bishkek Petroleum"
echo.
echo 2/4: PARTNER NEFT...
call :run_action %act% "%SCRIPT_PN%" "%UPLOAD_PN%" "PARTNER NEFT"
echo.
echo 3/4: Мунай Пром Food...
call :run_action %act% "%SCRIPT_MP_FOOD%" "%UPLOAD_MP_FOOD%" "Мунай Пром Food"
echo.
echo 4/4: Мунай Пром Non-Food...
call :run_action %act% "%SCRIPT_MP_NONFOOD%" "%UPLOAD_MP_NONFOOD%" "Мунай Пром Non-Food"
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
