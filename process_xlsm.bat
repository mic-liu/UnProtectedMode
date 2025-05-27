@echo off
if "%~1"=="" (
    echo 请将XLSM文件拖拽到本批处理文件上！
    pause
    exit /b
)

java -jar xlsm-processor-1.0-SNAPSHOT-jar-with-dependencies.jar "%~1"