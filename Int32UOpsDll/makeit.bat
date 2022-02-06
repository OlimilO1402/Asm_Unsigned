@echo off
if exist C:\Masm32Projects\UnsignedOps.obj del C:\Masm32Projects\UnsignedOps.obj
if exist C:\Masm32Projects\UnsignedOps.dll del C:\Masm32Projects\UnsignedOps.dll
copy "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\Int32UOpsDll\UnsignedOps.asm" "C:\Masm32Projects\UnsignedOps.asm"
copy "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\Int32UOpsDll\UnsignedOps.def" "C:\Masm32Projects\UnsignedOps.def"
cd C:\Masm32Projects\
C:\Masm32Projects\makeit.bat
rem C:\masm32\bin\ml /c /coff C:\Masm32Projects\UnsignedOps.asm
rem C:\masm32\bin\Link /SUBSYSTEM:WINDOWS /DLL /DEF:C:\Masm32Projects\UnsignedOps.def C:\Masm32Projects\UnsignedOps.obj 
copy "C:\Masm32Projects\UnsignedOps.dll" "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\Int32UOpsDll\UnsignedOps.dll"
copy "C:\Masm32Projects\UnsignedOps.lib" "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\Int32UOpsDll\UnsignedOps.lib"
del C:\Masm32Projects\UnsignedOps.obj
del C:\Masm32Projects\UnsignedOps.exp
dir C:\Masm32Projects\UnsignedOps.*
pause
