VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test Int32_UToStr"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test UnsignedOps-dll"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As Long)

' --------======== UInt32 Operations ========-------- '
Private Declare Function UInt32_Add_ref Lib "UnsignedOps" (ByRef pV1 As Long, ByRef pV2 As Long) As Long
Private Declare Function UInt32_Add Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_Sub Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_Mul Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_MulB Lib "UnsignedOps" Alias "UInt32_Mul" (ByVal v1 As Long, ByVal v2 As Long) As Currency
Private Declare Function UInt32_Div Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_DivB Lib "UnsignedOps" Alias "UInt32_Div" (ByVal v1 As Long, ByVal v2 As Long) As Currency

Private Declare Function UInt32_Shl Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Private Declare Function UInt32_Shr Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Private Declare Function UInt32_Sar Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Private Declare Function UInt32_Rol Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Private Declare Function UInt32_Rcl Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Private Declare Function UInt32_Ror Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Private Declare Function UInt32_Rcr Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Long
Private Declare Function UInt32_Shld Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Currency
Private Declare Function UInt32_Shrd Lib "UnsignedOps" (ByVal Value As Long, ByVal shifter As Long) As Currency

Private Declare Function UInt32_And Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_Or Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_Not Lib "UnsignedOps" (ByVal Value As Long) As Long
Private Declare Function UInt32_XOr Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_XNOr Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_NOr Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_NAnd Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long

Private Declare Sub UInt32_ToStr Lib "UnsignedOps" (ByVal Value As Long, ByVal pStr_out As LongPtr)
Private Declare Sub UInt32_ToHex Lib "UnsignedOps" (ByVal Value As Long, ByVal pStr_out As LongPtr)
Private Declare Sub UInt32_ToBin Lib "UnsignedOps" (ByVal Value As Long, ByVal pStr_out As LongPtr)

'just some short aliases
Private Declare Function U4Add_ref Lib "UnsignedOps" Alias "UInt32_UAdd_ref" (ByRef pV1 As Long, ByRef pV2 As Long) As Long
Private Declare Function U4Add Lib "UnsignedOps" Alias "UInt32_Add" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function U4Sub Lib "UnsignedOps" Alias "UInt32_Sub" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function U4Mul Lib "UnsignedOps" Alias "UInt32_Mul" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function U4MulB Lib "UnsignedOps" Alias "UInt32_Mul" (ByVal v1 As Long, ByVal v2 As Long) As Currency
Private Declare Function U4Div Lib "UnsignedOps" Alias "UInt32_Div" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function U4DivB Lib "UnsignedOps" Alias "UInt32_Div" (ByVal v1 As Long, ByVal v2 As Long) As Currency

' --------======== UInt64 Operations ========-------- '
Private Declare Function UInt64_Add Lib "UnsignedOps" (ByVal v1 As Currency, ByVal v2 As Currency) As Currency
Private Declare Function UInt64_Sub Lib "UnsignedOps" (ByVal v1 As Currency, ByVal v2 As Currency) As Currency
Private Declare Function UInt64_Mul Lib "UnsignedOps" (ByVal v1 As Currency, ByVal v2 As Currency) As Variant 'Decimal

Private Declare Sub UInt64_ToStr Lib "UnsignedOps" (ByVal Value As Currency, ByVal pStr_out As Long)
Private Declare Sub UInt64_ToHex Lib "UnsignedOps" (ByVal Value As Currency, ByVal pStr_out As Long)
Private Declare Sub UInt64_ToBin Lib "UnsignedOps" (ByVal Value As Currency, ByVal pStr_out As Long)
'
'errno_t _ultow_s(
'    unsigned long value,
'    wchar_t *str,
'    size_t sizeOfstr,
'    int radix
');

Private Type TVar
    Value As Variant
End Type

'https://www.vbarchiv.net/workshop/workshop_84-der-gebrauch-des-datentyps-decimal.html
Private Type TDec
    vt As Integer ' &HE = 14
    nnk As Byte   ' Anzahl der Nachkommastellen
    sgn As Byte   ' Vorzeichen
    LngVal3 As Long
    LngVal1 As Long
    LngVal2 As Long
End Type


'; int32_max = 2147483647
';uint32_max = 4294967294
';uint64_max = 18446744073709551615
'; int64_max = 9223372036854775807
';             922337203685477.5807

'; 79228162514264337593543950335
Private Sub Command3_Click()
    Dim c1 As Currency
    Dim c2 As Currency
    c1 = CCur(123456789012345#)
    c2 = CCur(123456789012345#)
    
    'Dim d As Variant
    'd = CDec(CDec(c1) * CDec(c2))
    
    'MsgBox d
    
    'MsgBox Hex(VarType(d)) '14 = &HE
    
    Dim tvv As TVar
    Dim ti8 As TDec
    
    tvv.Value = CDec("1")
    tvv.Value = CDec("-1")
    tvv.Value = CDec("2147483647")
    tvv.Value = CDec("2147483648")
    tvv.Value = CDec("4294967294")
    tvv.Value = CDec("4294967295")
    tvv.Value = CDec("9223372036854775807")
    tvv.Value = CDec("9223372036854775808")
    tvv.Value = CDec("18446744073709551615")
    tvv.Value = CDec("18446744073709551616")
    tvv.Value = CDec("79228162514264337593543950335") '* CDec("1024") ' + CDec("1023")
    
    RtlMoveMemory ti8, tvv, 16
    
    MsgBox "Biggest 96-bit unsigned value: " & vbCrLf & TDec_ToHex(ti8) & vbCrLf & CStr(tvv.Value)
    
End Sub

Private Function TDec_ToHex(Value As TDec) As String
    With Value
        TDec_ToHex = Hex4(.vt) & vbCrLf & Hex2(.nnk) & vbCrLf & Hex2(.sgn) & vbCrLf & _
                     Hex8(.LngVal3) & vbCrLf & Hex8(.LngVal1) & vbCrLf & Hex8(.LngVal2)
    End With
End Function
Private Sub Form_Load()
    Me.Caption = "Unsigned ops on signed Int32 - v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = 0
    Dim t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then Text1.Move L, t, W, H
End Sub

Private Sub Command1_Click()
    Text1.Text = ""
    TestDll
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Test_ToStr
End Sub

Sub Debug_Print(ByVal s As String)
    Text1.Text = Text1.Text & s & vbCrLf
End Sub

Sub TestDll()
    Dim v1 As Long
    Dim v2 As Long
    Dim lret As Long
    Dim cret As Currency
    Dim s As String
    
    v1 = 123
    v2 = 32
     
    lret = UInt32_Add_ref(v1, v2)
    Debug_Print lret '155
    
    lret = UInt32_Add(v1, v2)
    Debug_Print lret '155
    
    lret = UInt32_Sub(v1, v2)
    Debug_Print lret '91
    
    lret = UInt32_Mul(v1, v2)
    Debug_Print lret '3936
    
    v1 = 2147483647
    v2 = 100
    cret = UInt32_MulB(v1, v2)
    Debug_Print cret '21474836,4700
    
    v1 = 123456789
    v2 = 33
    lret = UInt32_Div(v1, v2)
    Debug_Print lret '3741114
    
    v1 = &HCAFE&
    lret = UInt32_Shl(v1, 8)
    Debug_Print "Shl(" & Hex(v1) & ", 8) = " & Hex(lret)
    
    v1 = UInt32_Shr(lret, 8)
    Debug_Print "Shr(" & Hex(lret) & ", 8) = " & Hex(v1)
    
    v1 = -v1
    lret = UInt32_Sar(v1, 8)
    Debug_Print "Sar(" & Hex(v1) & ", 8) = " & Hex(lret)
    
    v1 = &HCAFEBABE
    lret = UInt32_Rol(v1, 8)
    Debug_Print "Rol(" & Hex(v1) & ", 8) = " & Hex(lret)
    
    v1 = &HCAFEBABE
    lret = UInt32_Ror(v1, 8)
    Debug_Print "Ror(" & Hex(v1) & ", 8) = " & Hex(lret)
    
    v1 = &HCAFE&
    lret = UInt32_Rcl(v1, 12)
    Debug_Print "Rcl(" & Hex(v1) & ", 12) = " & Hex(lret)
    
    v1 = &HCAFE0000
    lret = UInt32_Rcr(v1, 12)
    Debug_Print "Rcr(" & Hex(v1) & ", 12) = " & Hex(lret)
    
    v1 = &HCAFEBABE
    v2 = &HB000&
    lret = UInt32_And(v1, v2)
    Debug_Print "And(" & Hex(v1) & ", " & Hex(v2) & ") = " & Hex(lret)
    
    v1 = &HCAFE0ABE
    v2 = &HB000&
    lret = UInt32_Or(v1, v2)
    Debug_Print "Or(" & Hex(v1) & ", " & Hex(v2) & ") = " & Hex(lret)
    
    v1 = &H35014541
    lret = UInt32_Not(v1)
    Debug_Print "Not(" & Hex(v1) & ") = " & Hex(lret)
    
    v1 = &HCAFE0ABE
    v2 = &HCAFEB000
    lret = UInt32_XOr(v1, v2)
    Debug_Print "XOr(" & Hex(v1) & ", " & Hex(v2) & ") = " & Hex(lret)
    
    v1 = &HCAFE0ABE
    v2 = &HCAFEB000
    lret = UInt32_XNOr(v1, v2)
    Debug_Print "XNOr(" & Hex(v1) & ", " & Hex(v2) & ") = " & Hex(lret)
    
    v1 = &HCAFE0ABE
    v2 = &HCAFEB000
    lret = UInt32_NOr(v1, v2)
    Debug_Print "NOr(" & Hex(v1) & ", " & Hex(v2) & ") = " & Hex(lret)
    
    v1 = &HCAFE0ABE
    v2 = &HCAFEB000
    lret = UInt32_NAnd(v1, v2)
    Debug_Print "NAnd(" & Hex(v1) & ", " & Hex(v2) & ") = " & Hex(lret)
    
    
    

    v1 = -1
    s = Space(10)
    UInt32_ToStr v1, StrPtr(s)
    Debug_Print Trim0(s)    '4294967295
    
    v1 = -1
    s = Space(8)
    UInt32_ToHex v1, StrPtr(s)
    Debug_Print Trim0(s)    '4294967295
    
    v1 = -1
    s = Space(32)
    UInt32_ToBin v1, StrPtr(s)
    Debug_Print Trim0(s)    '4294967295
    
    v1 = 32
    v2 = 65
    
    lret = U4Add(U4Sub(100, U4Mul(2, U4Div(v1, 4))), U4Div(v2, U4Add(6, 7)))
    s = Space(10): UInt32_ToStr lret, StrPtr(s)
    Debug_Print Trim0(s)
    
    lret = (100 - (2 * v1 / 4)) + (v2 / (6 + 7))
    s = Space(10): UInt32_ToStr lret, StrPtr(s)
    Debug_Print Trim0(s)
    
    Dim c1 As Currency
    Dim c2 As Currency
    
    c1 = CCur("822337203685477,5806")
    c2 = CCur("100000000000000,0001")
    
    cret = UInt64_Add(c1, c2)
    s = Space(20): UInt64_ToStr cret, StrPtr(s)
    Debug_Print Trim0(s)    '9223372036854775807
    
    cret = UInt64_Sub(c1, c2)
    s = Space(20): UInt64_ToStr cret, StrPtr(s)
    Debug_Print Trim0(s)    '7223372036854775805
    
    s = Space(20): UInt64_ToHex cret, StrPtr(s)
    Debug_Print Trim0(s)
    
    'cret = CCur("-922337203685477,5807")
    cret = CCur("-184467440737095")
    s = Space(66): UInt64_ToBin cret, StrPtr(s)
    Debug_Print Trim0(s)
End Sub

Sub Test_ToStr()

    Dim s As String
    Dim v As Long
    Dim ret As Long
    
    v = 123456789
    s = Int32_ToStr(v)
    Debug_Print s
    
    v = 2147483647
    s = Int32_ToStr(v)
    Debug_Print s
    
    v = &HCAFEBAB
    s = Int32_ToStr(v)
    Debug_Print s
    
    v = 123
    s = Int32_ToStr(v)
    Debug_Print s
    
    v = 4567
    s = Int32_ToStr(v)
    Debug_Print s
    
    
    Dim r As Long 'radix
    Dim L As Long
    
    v = 14021970
    r = 10
    L = LogN(v, r) + 1
    s = Space$(L)
    ret = Int32_ToStrR(v, StrPtr(s), r)
    Debug_Print s
    
    If ret <> L * 2 Then
        Debug_Print "error r<>l: r=" & r & "; l=" & L
    End If
    
    v = &H1CAFEBAB
    r = 16
    L = Ceiling(LogN(v, r))
    s = Space$(L)
    ret = Int32_ToStrR(v, StrPtr(s), r)
    
    Debug_Print s
    
    If ret <> L * 2 Then
        Debug_Print "error r<>l: r=" & r & "; l=" & L
    End If
    
    v = 987654321
    r = 10
    L = Ceiling(LogN(v, r))
    s = Space$(L)
    ret = Int32_ToStrR(v, StrPtr(s), r)
    
    Debug_Print s
    
    If ret <> L * 2 Then
        Debug_Print "error r<>l: r=" & r & "; l=" & L * 2
    End If
    
End Sub
