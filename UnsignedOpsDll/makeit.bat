@echo off
if exist C:\Masm32Projects\UnsignedOps\UnsignedOps.obj del C:\Masm32Projects\UnsignedOps\UnsignedOps.obj
if exist C:\Masm32Projects\UnsignedOps\UnsignedOps.dll del C:\Masm32Projects\UnsignedOps\UnsignedOps.dll
copy "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\UnsignedOpsDll\UnsignedOps.asm" "C:\Masm32Projects\UnsignedOps\UnsignedOps.asm"
copy "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\UnsignedOpsDll\UnsignedOps.def" "C:\Masm32Projects\UnsignedOps\UnsignedOps.def"
cd C:\Masm32Projects\UnsignedOps\
C:\Masm32Projects\UnsignedOps\makeit.bat
rem C:\masm32\bin\ml /c /coff C:\Masm32Projects\UnsignedOps\UnsignedOps.asm
rem C:\masm32\bin\Link /SUBSYSTEM:WINDOWS /DLL /DEF:C:\Masm32Projects\UnsignedOps\UnsignedOps.def C:\Masm32Projects\UnsignedOps\UnsignedOps.obj 
copy "C:\Masm32Projects\UnsignedOps\UnsignedOps.dll" "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\UnsignedOpsDll\UnsignedOps.dll"
copy "C:\Masm32Projects\UnsignedOps\UnsignedOps.lib" "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\UnsignedOpsDll\UnsignedOps.lib"
del C:\Masm32Projects\UnsignedOps\UnsignedOps.obj
del C:\Masm32Projects\UnsignedOps\UnsignedOps.exp
dir C:\Masm32Projects\UnsignedOps\UnsignedOps.*
pause