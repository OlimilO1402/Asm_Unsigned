Attribute VB_Name = "MUnsigned"
Option Explicit
Private Type TLng
    Value As Long
End Type
Private Type T2Int
    Value0 As Integer
    value1 As Integer
End Type
Private Type T2Lng
    Value0 As Long
    value1 As Long
End Type
Private Type TCur
    Value As Currency
End Type

' ARITHMETIC Operations
' --------~~~~~~~~========++++++++######## '  Unsigned Int16 arithmetic operations  ' ########++++++++========~~~~~~~~-------- '
Public Declare Function UInt16_Add Lib "UnsignedOps" (ByVal v1 As Integer, ByVal v2 As Integer) As Integer
Public Declare Function UInt16_Sub Lib "UnsignedOps" (ByVal v1 As Integer, ByVal v2 As Integer) As Integer
Public Declare Function UInt16_Mul Lib "UnsignedOps" (ByVal v1 As Integer, ByVal v2 As Integer) As Long
Public Declare Function UInt16_Div Lib "UnsignedOps" (ByVal v1 As Integer, ByVal v2 As Integer) As Integer

' --------~~~~~~~~========++++++++######## '  Unsigned Int32 arithmetic operations  ' ########++++++++========~~~~~~~~-------- '
Public Declare Function UInt32_Add Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function UInt32_Sub Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function UInt32_Mul Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Currency
'alternative declaration:
Public Declare Function UInt32_MulL Lib "UnsignedOps" Alias "UInt32_Mul" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function UInt32_Div Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long

' --------~~~~~~~~========++++++++######## '  Unsigned Int64 arithmetic operations  ' ########++++++++========~~~~~~~~-------- '
Public Declare Function UInt64_Add Lib "UnsignedOps" (ByVal v1 As Currency, ByVal v2 As Currency) As Currency
Public Declare Function UInt64_Sub Lib "UnsignedOps" (ByVal v1 As Currency, ByVal v2 As Currency) As Currency
Private Declare Sub UInt64Mul Lib "UnsignedOps" Alias "UInt64_Mul" (ByVal v1 As Currency, ByVal v2 As Currency, ByVal pVar_out As LongPtr) 'As Decimal
'alternative declaration:
'Public Declare Function UInt64Mul Lib "UnsignedOps" Alias "UInt64_Mul" (ByVal V1 As Currency, ByVal V2 As Currency, ByRef Var_out As Variant) 'As Decimal
Public Declare Function UInt64_Div Lib "UnsignedOps" (ByVal v1 As Currency, ByVal v2 As Currency) As Currency
'really?? why??:
Private Declare Sub UInt64Div Lib "UnsignedOps" Alias "UInt64_Div" (ByVal v1 As Currency, ByVal v2 As Currency, ByVal pVar_out As LongPtr) 'As Decimal

' ADDITIONAL Operations
' --------~~~~~~~~========++++++++######## '  Unsigned Int32 shifting operations  ' ########++++++++========~~~~~~~~-------- '
Public Declare Function UInt32_Shl Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Public Declare Function UInt32_Shr Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Public Declare Function UInt32_Sar Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Public Declare Function UInt32_Rol Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Public Declare Function UInt32_Rcl Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Public Declare Function UInt32_Ror Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Public Declare Function UInt32_Rcr Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Public Declare Function UInt32_Shld Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Currency
Public Declare Function UInt32_Shrd Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Currency

' --------~~~~~~~~========++++++++######## '  Unsigned Int32 boolean operations  ' ########++++++++========~~~~~~~~-------- '
Public Declare Function UInt32_And Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function UInt32_Or Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function UInt32_Not Lib "UnsignedOps" (ByVal Value As Long) As Long
Public Declare Function UInt32_XOr Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function UInt32_XNOr Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function UInt32_NOr Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function UInt32_NAnd Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long

' INPUT / OUTPUT functions
' --------~~~~~~~~========++++++++######## '  Unsigned Int16 input/output functions  ' ########++++++++========~~~~~~~~-------- '
Private Declare Sub UInt16ToStr Lib "UnsignedOps" Alias "UInt16_ToStr" (ByVal Value As Long, ByVal pStr_out As LongPtr)
Private Declare Sub UInt16ToHex Lib "UnsignedOps" Alias "UInt16_ToHex" (ByVal Value As Long, ByVal pStr_out As LongPtr)
Private Declare Sub UInt16ToBin Lib "UnsignedOps" Alias "UInt16_ToBin" (ByVal Value As Long, ByVal pStr_out As LongPtr) 'Bin as in Binary, don't think of a trash bin ;-)
Private Declare Sub UInt16ToDec Lib "UnsignedOps" Alias "UInt16_ToDec" (ByVal Value As Long, ByVal pStr_out As LongPtr)
Private Declare Function UInt16Parse Lib "UnsignedOps" Alias "UInt16_Parse" (ByVal pStrVal As LongPtr, ByVal radix As Long) As Long 'Integer
Private Declare Function UInt16ParseI Lib "UnsignedOps" Alias "UInt16_Parse" (ByVal pStrVal As LongPtr, ByVal radix As Long) As Integer

' --------~~~~~~~~========++++++++######## '  Unsigned Int32 input/output functions  ' ########++++++++========~~~~~~~~-------- '
Public Declare Sub UInt32ToStr Lib "UnsignedOps" Alias "UInt32_ToStr" (ByVal Value As Long, ByVal pStr_out As LongPtr)
Public Declare Sub UInt32ToHex Lib "UnsignedOps" Alias "UInt32_ToHex" (ByVal Value As Long, ByVal pStr_out As LongPtr)
Public Declare Sub UInt32ToBin Lib "UnsignedOps" Alias "UInt32_ToBin" (ByVal Value As Long, ByVal pStr_out As LongPtr)
Public Declare Sub UInt32ToDec Lib "UnsignedOps" Alias "UInt32_ToDec" (ByVal Value As Long, ByRef Dec_out As Variant)
Public Declare Function UInt32_Parse Lib "UnsignedOps" (ByVal pStr As LongPtr, ByVal radix As Long) As Long

' --------~~~~~~~~========++++++++######## '  Unsigned Int64 input/output functions  ' ########++++++++========~~~~~~~~-------- '
Public Declare Sub UInt64ToStr Lib "UnsignedOps" Alias "UInt64_ToStr" (ByVal Value As Currency, ByVal pStr_out As LongPtr)
Public Declare Sub UInt64ToHex Lib "UnsignedOps" Alias "UInt64_ToHex" (ByVal Value As Currency, ByVal pStr_out As LongPtr)
Public Declare Sub UInt64ToBin Lib "UnsignedOps" Alias "UInt64_ToBin" (ByVal Value As Currency, ByVal pStr_out As LongPtr)
Public Declare Sub UInt64ToDec Lib "UnsignedOps" Alias "UInt64_ToDec" (ByVal Value As Long, ByRef Dec_out As Variant)
Public Declare Function UInt64_Parse Lib "UnsignedOps" (ByVal pStr As LongPtr, ByVal radix As Long) As Currency

'just some short aliases and additional stuff
Public Declare Function U4Add_ref Lib "UnsignedOps" Alias "UInt32_Add_ref" (ByRef pV1 As Long, ByRef pV2 As Long) As Long
'Public Declare Function U2Add_ref Lib "UnsignedOps" Alias "UInt32_UAdd_ref" (ByRef pV1 As Integer, ByRef pV2 As Integer) As Long 'Integer

Public Declare Function U4Add Lib "UnsignedOps" Alias "UInt32_Add" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function U4Sub Lib "UnsignedOps" Alias "UInt32_Sub" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function U4Mul Lib "UnsignedOps" Alias "UInt32_Mul" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function U4MulB Lib "UnsignedOps" Alias "UInt32_Mul" (ByVal v1 As Long, ByVal v2 As Long) As Currency
Public Declare Function U4Div Lib "UnsignedOps" Alias "UInt32_Div" (ByVal v1 As Long, ByVal v2 As Long) As Long
Public Declare Function U4DivB Lib "UnsignedOps" Alias "UInt32_Div" (ByVal v1 As Long, ByVal v2 As Long) As Currency
Public Declare Sub UInt64_Test Lib "UnsignedOps" (ByRef Dec_out As Variant) 'As Decimal
Public Declare Function UInt32_Add_ref Lib "UnsignedOps" (ByRef pV1 As Long, ByRef pV2 As Long) As Long

' --------~~~~~~~~========++++++++######## '  Unsigned Int16 input/output functions  ' ########++++++++========~~~~~~~~-------- '
Public Function UInt16_ToUInt32(ByVal Value As Integer) As Long
    Dim t2i As T2Int: t2i.Value0 = Value
    Dim tl  As TLng:     LSet tl = t2i
    UInt16_ToUInt32 = tl.Value
End Function
Public Function UInt16_ToStr(ByVal Value As Integer) As String
    UInt16_ToStr = Space(5) & vbNullChar
    UInt16ToStr UInt16_ToUInt32(Value), StrPtr(UInt16_ToStr)
    UInt16_ToStr = Trim0(UInt16_ToStr)
End Function
Public Function UInt16_ToHex(ByVal Value As Long) As String
    UInt16_ToHex = Space(4) & vbNullChar
    UInt16ToHex UInt16_ToUInt32(Value), StrPtr(UInt16_ToHex)
    UInt16_ToHex = "&H" & Trim0(UInt16_ToHex)
End Function
Public Function UInt16_ToBin(ByVal Value As Long) As String
    UInt16_ToBin = Space(16) & vbNullChar
    UInt16ToBin UInt16_ToUInt32(Value), StrPtr(UInt16_ToBin)
    UInt16_ToBin = "&B" & Trim0(UInt16_ToBin)
End Function
Public Function UInt16_Parse(ByVal Value As String) As Integer
    Value = Value & vbNullChar
    Dim tl As TLng, t2i As T2Int
    tl.Value = UInt16Parse(StrPtr(Value), 10)
    If tl.Value <= 65535 Then
        LSet t2i = tl
        UInt16_Parse = t2i.Value0
    Else
        'give error information, like overflow etc
    End If
End Function
Public Function UInt16_TryParse(ByVal Value As String, Val_out As Integer, Optional ByVal radix As Long = 10) As Boolean
    'radix=2: binary 01; radix=10: deciml 0123456789; radix=16: hexadecimal 0123456789ABCDEF
Try: On Error GoTo Catch
    If Len(Value) = 0 Then Exit Function
    Value = Value & vbNullChar
    Dim tl As TLng, t2i As T2Int
    tl.Value = UInt16Parse(StrPtr(Value), radix)
    If tl.Value <= 65535 Then
        LSet t2i = tl
        Val_out = t2i.Value0
    Else
        'give error information, overflow
    End If
    UInt16_TryParse = True
Catch:
End Function

' --------~~~~~~~~========++++++++######## '  Unsigned Int32 input/output functions  ' ########++++++++========~~~~~~~~-------- '
Public Function UInt32_ToUInt64(ByVal Value As Long) As Currency
    Dim t2l As T2Lng: t2l.Value0 = Value
    Dim tc  As TCur:     LSet tc = t2l
    UInt32_ToUInt64 = tc.Value
End Function
Public Function UInt32_ToStr(ByVal Value As Long) As String
    UInt32_ToStr = Space(10)
    UInt32ToStr Value, StrPtr(UInt32_ToStr)
    UInt32_ToStr = Trim$(UInt32_ToStr)
End Function
Public Function UInt32_ToHex(ByVal Value As Long) As String
    UInt32_ToHex = Space(8)
    UInt32ToHex Value, StrPtr(UInt32_ToHex)
    UInt32_ToHex = "&H" & Trim$(UInt32_ToHex)
End Function
Public Function UInt32_ToBin(ByVal Value As Long) As String
    UInt32_ToBin = Space(32)
    UInt32ToBin Value, StrPtr(UInt32_ToBin)
    UInt32_ToBin = "&B" & Trim$(UInt32_ToBin)
End Function
Public Function UInt32_ToDec(ByVal Value As Long) As Variant
    UInt32ToDec Value, UInt32_ToDec
End Function

Public Function UInt32_TryParse(ByVal Value As String, Val_out As Long, Optional radix As Long = 10) As Boolean
Try: On Error GoTo Catch
    If Len(Value) = 0 Then Exit Function
    Value = Value & vbNullChar
    Dim v As Long
    v = UInt32_Parse(StrPtr(Value), radix)
    Val_out = v
    UInt32_TryParse = True
    Exit Function
Catch:
End Function

' --------~~~~~~~~========++++++++######## '  Unsigned Int64 arithmetic operations  ' ########++++++++========~~~~~~~~-------- '
Public Function UInt64_Mul(ByVal v1 As Currency, ByVal v2 As Currency) As Variant
    UInt64Mul v1, v2, VarPtr(UInt64_Mul)
End Function
'Public Function UInt64_Div(ByVal V1 As Currency, ByVal V2 As Currency) As Variant
'    UInt64Div V1, V2, VarPtr(UInt64_Mul)
'End Function


' --------~~~~~~~~========++++++++######## '  Unsigned Int64 input/output functions  ' ########++++++++========~~~~~~~~-------- '
Public Function UInt64_ToStr(ByVal Value As Currency) As String
    UInt64_ToStr = Space(20)
    UInt64ToStr Value, StrPtr(UInt64_ToStr)
    UInt64_ToStr = Trim$(UInt64_ToStr)
End Function
Public Function UInt64_ToHex(ByVal Value As Currency) As String
    UInt64_ToHex = Space(16)
    UInt64ToHex Value, StrPtr(UInt64_ToHex)
    UInt64_ToHex = "&H" & Trim$(UInt64_ToHex)
End Function
Public Function UInt64_ToBin(ByVal Value As Currency) As String
    UInt64_ToBin = Space(64)
    UInt64ToBin Value, StrPtr(UInt64_ToBin)
    UInt64_ToBin = "&B" & Trim$(UInt64_ToBin)
End Function
Public Function UInt64_ToDec(ByVal Value As Currency) As Variant
    UInt64ToDec Value, UInt64_ToDec
End Function

