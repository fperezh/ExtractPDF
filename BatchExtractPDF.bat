@echo off
:loop
C:\Users\USEROA\AppData\Local\Programs\Python\Python312\python.exe C:\ExtractPDF\ExtractPDF.py
echo Ejecutando tarea...
timeout /t 2 /nobreak > nul
goto loop