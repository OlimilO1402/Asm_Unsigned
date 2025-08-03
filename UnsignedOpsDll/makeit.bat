@echo off
if exist C:\MASM\Projects\UnsignedOps\UnsignedOps.obj del C:\MASM\Projects\UnsignedOps\UnsignedOps.obj
if exist C:\MASM\Projects\UnsignedOps\UnsignedOps.dll del C:\MASM\Projects\UnsignedOps\UnsignedOps.dll
copy "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\UnsignedOpsDll\UnsignedOps.asm" "C:\MASM\Projects\UnsignedOps\UnsignedOps.asm"
copy "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\UnsignedOpsDll\UnsignedOps.def" "C:\MASM\Projects\UnsignedOps\UnsignedOps.def"
cd C:\MASM\Projects\UnsignedOps\
C:\MASM\Projects\UnsignedOps\makeit.bat
rem C:\MASM\x86\ml /c /coff C:\Masm32Projects\UnsignedOps\UnsignedOps.asm
rem C:\MASM\x86\link /SUBSYSTEM:WINDOWS /DLL /DEF:C:\MASM\Projects\UnsignedOps\UnsignedOps.def C:\MASM\Projects\UnsignedOps\UnsignedOps.obj 
copy "C:\MASM\Projects\UnsignedOps\UnsignedOps.dll" "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\UnsignedOpsDll\UnsignedOps.dll"
copy "C:\MASM\Projects\UnsignedOps\UnsignedOps.lib" "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_Unsigned\UnsignedOpsDll\UnsignedOps.lib"
del C:\MASM\Projects\UnsignedOps\UnsignedOps.obj
del C:\MASM\Projects\UnsignedOps\UnsignedOps.exp
dir C:\MASM\Projects\UnsignedOps\UnsignedOps.*
pause