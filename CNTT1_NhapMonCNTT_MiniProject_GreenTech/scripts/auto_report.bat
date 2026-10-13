@echo off
setlocal
set "ROOT=%~dp0.."
set "LOG=%ROOT%\data\log-he-thong.log"
set "REPORT=%ROOT%\backup\bao-cao-su-co.txt"

if not exist "%LOG%" (
    echo [ERROR] Khong tim thay file log: %LOG%
    exit /b 1
)

>>"%REPORT%" echo ===== BAO CAO SU CO HE THONG =====
>>"%REPORT%" echo Thoi gian chay: %date% %time%
type "%LOG%" | findstr /I /C:"CRITICAL" /C:"ERROR" /C:"TIMEOUT" >> "%REPORT%"
>>"%REPORT%" echo ===== KET THUC =====
>>"%REPORT%" echo.

echo [OK] Hoan tat loc log. Bao cao: %REPORT%
endlocal
