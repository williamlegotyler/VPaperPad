@echo off

IF EXIST "C:\Windows\SysWOW64\regsvr32.exe" (
"C:\Windows\SysWOW64\regsvr32" "Kixforms.dll"
) ELSE (
regsvr32 "Kixforms.dll"
)
exit