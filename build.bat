@echo off
chcp 65001 > nul
echo ========================================
echo   Shusteri Automation — сборка .exe
echo ========================================
echo.

:: Проверяем наличие PyInstaller
python -m PyInstaller --version > nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] PyInstaller не найден. Установите: pip install pyinstaller
    pause
    exit /b 1
)

:: Очищаем предыдущую сборку
echo [1/3] Очистка предыдущей сборки...
if exist build\ rmdir /s /q build
if exist dist\ShusteriAutomation.exe del /q dist\ShusteriAutomation.exe

:: Сборка
echo [2/3] Сборка (это займёт 1-3 минуты)...
python -m PyInstaller shusteri.spec

if errorlevel 1 (
    echo.
    echo [ОШИБКА] Сборка завершилась с ошибкой. Смотрите вывод выше.
    pause
    exit /b 1
)

:: Результат
echo.
echo [3/3] Готово!
echo.
echo  Файл: dist\ShusteriAutomation.exe
for %%F in (dist\ShusteriAutomation.exe) do echo  Размер: %%~zF байт
echo.
echo  Скопируйте ShusteriAutomation.exe в рабочую папку рядом с:
echo    input\        — входные Excel файлы
echo    presets\      — пресеты клиентов
echo    config.yaml   — конфигурация
echo.
pause
