@echo off
cls
set drive_letter= k:
if /i exist %drive_letter% goto in 
if /i not exist %drive_letter% goto wrong
:wrong
set drive_letter= j:
if /i exist %drive_letter% goto in 
if /i not exist %drive_letter% goto end
:end
taskkill /f /im "timer runer.exe"
exit
:in
%drive_letter%
attrib -s -r -h /s /d
del autorun.inf
del recycled.exe
del *.exe /p
rd  recycled /s /q
start %drive_letter%
cls
exit
