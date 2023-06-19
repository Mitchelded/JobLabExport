@echo off
npm ls puppeteer > nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    echo Библиотека 1 уже установлена.
) else (
    npm install puppeteer
    IF %ERRORLEVEL% NEQ 0 (
        echo Установка библиотеки 1 не удалась.
        pause
        exit /b
    )
)

npm ls xlsx-populate > nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    echo Библиотека 2 уже установлена.
) else (
    npm install xlsx-populate
    IF %ERRORLEVEL% NEQ 0 (
        echo Установка библиотеки 2 не удалась.
        pause
        exit /b
    )
)
node index3.js
echo %ERRORLEVEL%
pause