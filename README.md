# Asm_Unsigned  
## Unsigned arithmetic, boolean and shifting operations for VB using VB.Long, VB.Currency and VB.Decimal  
  
[![GitHub](https://img.shields.io/github/license/OlimilO1402/Asm_Unsigned?style=plastic)](https://github.com/OlimilO1402/Asm_Unsigned/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Asm_Unsigned?style=plastic)](https://github.com/OlimilO1402/Asm_Unsigned/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Asm_Unsigned/total.svg)](https://github.com/OlimilO1402/Asm_Unsigned/releases/download/v2.3.4/UnsignedOps_v2.3.4.zip)
[![Follow](https://img.shields.io/github/followers/OlimilO1402.svg?style=social&label=Follow&maxAge=2592000)](https://github.com/OlimilO1402/Asm_Unsigned/watchers)

Project started in january 2022.  
Except for the datatype Byte, VB does not have any other intrinsic unsigned datatypes.  
This is a dll in asm with some functions to do arithmetic, boolean and shifting operations on UInt32 and UInt64 using VB.Long, VB.Currency and VB.Decimal just like they were unsigned.  

the function UInt64_Mul will take 2 Currency-Variables (As UInt64) and the result will be returned in an Decimal. 
Decimal is a Variant, and as such it consumes 128-Bits of memory, but in total  it has a precision of 96-Bit in Visual Basic 6.
This is more than what you have in other languages.

![<AppName> Image](Resources/XL-UInt64_Mul.png "XL-UInt64_Mul Image")
