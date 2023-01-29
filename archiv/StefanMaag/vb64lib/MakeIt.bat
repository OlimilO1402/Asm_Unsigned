@echo off

if exist vbExt.obj del vbExt.obj
if exist vbExt.dll del vbExt.dll


\masm32\bin\ml /c /coff /Cp vbExt.asm
\masm32\bin\link /DLL /DEF:vbExt.def /SUBSYSTEM:WINDOWS /LIBPATH:c:\masm32\lib vbExt.obj

dir vbExt.*

pause