@echo off
rem cd c:\agencyprof c:\windows\syswow64
rem copy COMDLG32.OCX c:\windows\syswow64
rem copy MSComCtl.ocx c:\windows\syswow64
rem copy Resizer.ocx c:\windows\syswow64
rem copy cswsk32.ocx c:\windows\syswow64

echo "all .ocx files must have been moved to c:\windows\sysWOW64."
echo "this file must be run as administrator"
echo "press enter to continue."
pause
cd c:\windows\syswow64
regsvr32.exe COMDLG32.OCX
regsvr32.exe MSComCtl.ocx
regsvr32.exe Resizer.ocx
regsvr32.exe cswsk32.ocx
pause