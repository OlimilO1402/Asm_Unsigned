@echo off
if exist C:\Masm32Projects\Int32UOps.obj del C:\Masm32Projects\Int32UOps.obj
if exist C:\Masm32Projects\Int32UOps.dll del C:\Masm32Projects\Int32UOps.dll
copy "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\Int32UOpsDll\Int32UOps.asm" "C:\Masm32Projects\Int32UOps.asm"
copy "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\Int32UOpsDll\Int32UOps.def" "C:\Masm32Projects\Int32UOps.def"
cd C:\Masm32Projects\
C:\Masm32Projects\makeit.bat
rem C:\masm32\bin\ml /c /coff C:\Masm32Projects\Int32UOps.asm
rem C:\masm32\bin\Link /SUBSYSTEM:WINDOWS /DLL /DEF:C:\Masm32Projects\Int32UOps.def C:\Masm32Projects\Int32UOps.obj 
copy "C:\Masm32Projects\Int32UOps.dll" "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\Int32UOpsDll\Int32UOps.dll"
copy "C:\Masm32Projects\Int32UOps.lib" "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\Int32UOpsDll\Int32UOps.lib"
del C:\Masm32Projects\Int32UOps.obj
del C:\Masm32Projects\Int32UOps.exp
dir C:\Masm32Projects\Int32UOps.*
pause
