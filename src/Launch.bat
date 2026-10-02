@echo off

IF EXIST "C:\Windows\SysWOW64\wscript.exe" (
start "" "C:\Windows\SysWOW64\wscript.exe" VPaperPad.vbs
) ELSE (
start "" VPaperPad.vbs
)
exit
