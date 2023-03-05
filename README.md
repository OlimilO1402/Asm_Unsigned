# Asm_Unsigned  
## Unsigned arithmetic, boolean and shifting operations for VB using VB.Integer, VB.Long, VB.Currency and VB.Decimal  
  
[![GitHub](https://img.shields.io/github/license/OlimilO1402/Asm_Unsigned?style=plastic)](https://github.com/OlimilO1402/Asm_Unsigned/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Asm_Unsigned?style=plastic)](https://github.com/OlimilO1402/Asm_Unsigned/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Asm_Unsigned/total.svg)](https://github.com/OlimilO1402/Asm_Unsigned/releases/download/v2023.3.5/UnsignedOps_v2023.3.5.zip)
![GitHub followers](https://img.shields.io/github/followers/OlimilO1402?style=social)

Project started in january 2022.  
Except for the datatype Byte, VB does not have any other intrinsic unsigned datatypes.  
This is a dll in asm with some functions to do arithmetic, boolean and shifting operations on UInt16, UInt32 and UInt64 using VB.Integer, VB.Long, VB.Currency and VB.Decimal just like they were unsigned.  

the function UInt64_Mul will take 2 Currency-Variables (As UInt64) and the result will be returned in a Decimal. 
Decimal is a Variant, and as such it consumes 128-Bits of memory, but in total  it has a precision of 96-Bit in Visual Basic 6.
This is more than what you have in other languages.  

If you want to compile the project you need the following repos:  
* [Sys_Strings](https://github.com/OlimilO1402/Sys_Strings)
* [Ptr_Pointers](https://github.com/OlimilO1402/Ptr_Pointers)
  
![<AppName> Image](Resources/XL-UInt64_Mul.png "XL-UInt64_Mul Image")
