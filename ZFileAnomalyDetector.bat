@echo off

powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show(\"Welcome to ZFile Anomaly Detector v1.0`nBy Abbas Sobhi - June 2025\", \"ZFile Anomaly Detector\", 'OK', 'Information')"


Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass

powershell -ExecutionPolicy Bypass -File "C:\Users\abbas\Desktop\PDF\ZFileAnomalyDetector\ZFileAnomalyDetector.ps1"

pause
