@echo off

IF EXIST "C:\Windows\SysWOW64\regsvr32.exe" (
"C:\Windows\SysWOW64\regsvr32" /u "Kixforms.dll"
) ELSE (
regsvr32 /u "Kixforms.dll"
)
exit