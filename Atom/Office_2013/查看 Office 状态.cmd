echo off
set "ospp=%ProgramFiles%\Microsoft Office\Office15\ospp.vbs"
if not exist "%ospp%" (
set "ospp=%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs"
)

cscript "%ospp%" /dstatus
pause