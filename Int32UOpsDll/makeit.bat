@echo off
if exist C:\Masm32Projects\Int32UOps.obj del C:\Masm32Projects\Int32UOps.obj
if exist C:\Masm32Projects\Int32UOps.dll del C:\Masm32Projects\Int32UOps.dll
C:\masm32\bin\ml /c /coff C:\Masm32Projects\Int32UOps.asm
C:\masm32\bin\Link /SUBSYSTEM:WINDOWS /DLL /DEF:C:\Masm32Projects\Int32UOps.def C:\Masm32Projects\Int32UOps.obj 
del C:\Masm32Projects\Int32UOps.obj
del C:\Masm32Projects\Int32UOps.exp
dir C:\Masm32Projects\Int32UOps.*
pause
