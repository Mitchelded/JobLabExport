@echo off
npm ls puppeteer > nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    echo ���������� 1 ��� �����������.
) else (
    npm install puppeteer
    IF %ERRORLEVEL% NEQ 0 (
        echo ��������� ���������� 1 �� �������.
        pause
        exit /b
    )
)

npm ls xlsx-populate > nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    echo ���������� 2 ��� �����������.
) else (
    npm install xlsx-populate
    IF %ERRORLEVEL% NEQ 0 (
        echo ��������� ���������� 2 �� �������.
        pause
        exit /b
    )
)
node index3.js
echo %ERRORLEVEL%
pause