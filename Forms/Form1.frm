VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   10200
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Test Parsing"
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   4320
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
'
'errno_t _ultow_s(
'    unsigned long value,
'    wchar_t *str,
'    size_t sizeOfstr,
'    int radix
');


'; int32_max = 2147483647
';uint32_max = 4294967294
';uint64_max = 18446744073709551615
'; int64_max = 9223372036854775807
';             922337203685477.5807

'; 79228162514264337593543950335
Private Sub Command3_Click()
    Form2.Show
    'Dim c1 As Currency
    'Dim c2 As Currency
    'c1 = CCur(123456789012345#)
    'c2 = CCur(123456789012345#)
    
    'Dim d As Variant
    'd = CDec(CDec(c1) * CDec(c2))
    
    'MsgBox d
    
    'MsgBox Hex(VarType(d)) '14 = &HE
    
'    Dim tvv As TVar
'    Dim ti8 As TDec
'
'    tvv.Value = CDec("1")
'    tvv.Value = CDec("-1")
'    tvv.Value = CDec("2147483647")
'    tvv.Value = CDec("2147483648")
'    tvv.Value = CDec("4294967294")
'    tvv.Value = CDec("4294967295")
'    tvv.Value = CDec("9223372036854775807")
'    tvv.Value = CDec("9223372036854775808")
'    tvv.Value = CDec("18446744073709551615")
'    tvv.Value = CDec("18446744073709551616")
'    tvv.Value = CDec("79228162514264337593543950335") '* CDec("1024") ' + CDec("1023")
'
'    RtlMoveMemory ti8, tvv, 16
'
'    MsgBox "Biggest 96-bit unsigned value: " & vbCrLf & TDec_ToHex(ti8) & vbCrLf & CStr(tvv.Value)
    
End Sub

Private Sub Command4_Click()
'    Dim v As Variant
'
'    UInt64_Test v
'
'    MsgBox VarType(v)
End Sub

Private Sub Command5_Click()
    Dim s As String
    Dim i1 As Integer: i1 = 1234
    Dim i2 As Integer: i2 = 2345
    
    Dim isum As Integer: isum = UInt16_Add(i1, i2)
    
    MsgBox isum
    
    s = "65535"
    
    i1 = UInt16_Parse(s)
    MsgBox UInt16_ToStr(i1)
    MsgBox UInt16_ToHex(i1)
    MsgBox UInt16_ToBin(i1)
    
End Sub

Private Sub Command6_Click()
    Dim s As String
    Dim sVal As String
    
    Dim i As Integer
    sVal = "1010101010101010"
    If Not MUnsigned.UInt16_TryParse(sVal, i, 2) Then Exit Sub
    s = s & "s: " & sVal & " = v: " & MUnsigned.UInt16_ToStr(i) & vbCrLf
    
    sVal = "65535"
    If Not MUnsigned.UInt16_TryParse(sVal, i, 10) Then Exit Sub
    s = s & "s: " & sVal & " = v: " & MUnsigned.UInt16_ToStr(i) & vbCrLf
    
    sVal = "ABCD"
    If Not MUnsigned.UInt16_TryParse(sVal, i, 16) Then Exit Sub
    s = s & "s: " & sVal & " = v: " & MUnsigned.UInt16_ToStr(i) & vbCrLf
    
    Dim L As Long
    sVal = "10101010101010101010101010101010"
    If Not MUnsigned.UInt32_TryParse(sVal, L, 2) Then Exit Sub
    s = s & "s: " & sVal & " = v: " & MUnsigned.UInt32_ToStr(L) & vbCrLf
    
    sVal = "1234567890"
    If Not MUnsigned.UInt32_TryParse(sVal, L, 10) Then Exit Sub
    s = s & "s: " & sVal & " = v: " & MUnsigned.UInt32_ToStr(L) & vbCrLf
    
    sVal = "ABCDEF98"
    If Not MUnsigned.UInt32_TryParse(sVal, L, 16) Then Exit Sub
    s = s & "s: " & sVal & " = v: " & MUnsigned.UInt32_ToStr(L) & vbCrLf
    
    Text1.Text = s
End Sub

Private Sub Command7_Click()
Try: On Error GoTo Catch
    Dim divid As Currency: divid = UInt64_Parse("9223372036854775", 10)
    Dim divis As Currency: divis = UInt64_Parse("2", 10)
    Debug.Print divid
    Debug.Print "-------------------"
    Debug.Print "         " & divis
    Dim res As Currency
    
    res = UInt64_Div(divid, divis)
    
    MsgBox UInt64_ToStr(res) '767.1072
    Exit Sub
Catch:
   MsgBox "Error: " & Err.Number & vbCrLf & Err.LastDllError & vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
    Me.Caption = "Unsigned operations on signed Int32+Int64 - v" & App.Major & "." & App.Minor & "." & App.Revision
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
    Dim V1 As Long
    Dim V2 As Long
    Dim lret As Long
    Dim cret As Currency
    Dim s As String
    
    V1 = 123
    V2 = 32
     
    lret = UInt32_Add_ref(V1, V2)
    Debug_Print lret '155
    
    lret = UInt32_Add(V1, V2)
    Debug_Print lret '155
    
    lret = UInt32_Sub(V1, V2)
    Debug_Print lret '91
    
    lret = UInt32_Mul(V1, V2)
    Debug_Print lret '3936
    
    V1 = 2147483647
    V2 = 100
    cret = UInt32_Mul(V1, V2)
    Debug_Print cret '21474836,4700
    
    V1 = 123456789
    V2 = 33
    lret = UInt32_Div(V1, V2)
    Debug_Print lret '3741114
    
    V1 = &HCAFE&
    lret = UInt32_Shl(V1, 8)
    Debug_Print "Shl(" & Hex(V1) & ", 8) = " & Hex(lret)
    
    V1 = UInt32_Shr(lret, 8)
    Debug_Print "Shr(" & Hex(lret) & ", 8) = " & Hex(V1)
    
    V1 = -V1
    lret = UInt32_Sar(V1, 8)
    Debug_Print "Sar(" & Hex(V1) & ", 8) = " & Hex(lret)
    
    V1 = &HCAFEBABE
    lret = UInt32_Rol(V1, 8)
    Debug_Print "Rol(" & Hex(V1) & ", 8) = " & Hex(lret)
    
    V1 = &HCAFEBABE
    lret = UInt32_Ror(V1, 8)
    Debug_Print "Ror(" & Hex(V1) & ", 8) = " & Hex(lret)
    
    V1 = &HCAFE&
    lret = UInt32_Rcl(V1, 12)
    Debug_Print "Rcl(" & Hex(V1) & ", 12) = " & Hex(lret)
    
    V1 = &HCAFE0000
    lret = UInt32_Rcr(V1, 12)
    Debug_Print "Rcr(" & Hex(V1) & ", 12) = " & Hex(lret)
    
    V1 = &HCAFEBABE
    V2 = &HB000&
    lret = UInt32_And(V1, V2)
    Debug_Print "And(" & Hex(V1) & ", " & Hex(V2) & ") = " & Hex(lret)
    
    V1 = &HCAFE0ABE
    V2 = &HB000&
    lret = UInt32_Or(V1, V2)
    Debug_Print "Or(" & Hex(V1) & ", " & Hex(V2) & ") = " & Hex(lret)
    
    V1 = &H35014541
    lret = UInt32_Not(V1)
    Debug_Print "Not(" & Hex(V1) & ") = " & Hex(lret)
    
    V1 = &HCAFE0ABE
    V2 = &HCAFEB000
    lret = UInt32_XOr(V1, V2)
    Debug_Print "XOr(" & Hex(V1) & ", " & Hex(V2) & ") = " & Hex(lret)
    
    V1 = &HCAFE0ABE
    V2 = &HCAFEB000
    lret = UInt32_XNOr(V1, V2)
    Debug_Print "XNOr(" & Hex(V1) & ", " & Hex(V2) & ") = " & Hex(lret)
    
    V1 = &HCAFE0ABE
    V2 = &HCAFEB000
    lret = UInt32_NOr(V1, V2)
    Debug_Print "NOr(" & Hex(V1) & ", " & Hex(V2) & ") = " & Hex(lret)
    
    V1 = &HCAFE0ABE
    V2 = &HCAFEB000
    lret = UInt32_NAnd(V1, V2)
    Debug_Print "NAnd(" & Hex(V1) & ", " & Hex(V2) & ") = " & Hex(lret)
    
    
    

    V1 = -1
    's = Space(10)
    s = UInt32_ToStr(V1)  ', StrPtr(s)
    Debug_Print s 'Trim0(s)    '4294967295
    
    V1 = -1
    's = Space(8)
    s = UInt32_ToHex(V1)  ', StrPtr(s)
    Debug_Print s 'Trim0(s)    '4294967295
    
    V1 = -1
    's = Space(32)
    s = UInt32_ToBin(V1) '1, StrPtr(s)
    Debug_Print s 'Trim0(s)    '4294967295
    
    V1 = 32
    V2 = 65
    
    lret = U4Add(U4Sub(100, U4Mul(2, U4Div(V1, 4))), U4Div(V2, U4Add(6, 7)))
    's = Space(10): UInt32_ToStr lret, StrPtr(s)
    s = UInt32_ToStr(lret)
    Debug_Print s
    
    lret = (100 - (2 * V1 / 4)) + (V2 / (6 + 7))
    's = Space(10): UInt32_ToStr lret, StrPtr(s)
    s = UInt32_ToStr(lret)
    Debug_Print s
    
    Dim c1 As Currency
    Dim c2 As Currency
    
    c1 = CCur("822337203685477,5806")
    c2 = CCur("100000000000000,0001")
    
    cret = UInt64_Add(c1, c2)
    's = Space(20): UInt64_ToStr cret, StrPtr(s)
    s = UInt64_ToStr(cret)
    Debug_Print s    '9223372036854775807
    
    cret = UInt64_Sub(c1, c2)
    's = Space(20): UInt64_ToStr cret, StrPtr(s)
    s = UInt64_ToStr(cret) ', StrPtr(s)
    Debug_Print s    '7223372036854775805
    
    V1 = &HCAFEBABE
    Dim d: d = UInt32_ToDec(V1)
    Debug_Print VarType(d) & " " & CStr(d)  '& " " & Hex(d) '14 3405691582
    
    c1 = 12345.6789
    c2 = 98765.4321
    
    d = UInt64_Mul(c1, c2)
    Debug_Print VarType(d) & " " & CStr(d) '14 121932631112635269
    
    '2147483647 * 2147483647 = 4611686014132420609
    'c1 = 214748.3647
    'c2 = 214748.3647
    
    '12345678901234 * 23456789012345 = 289589985200405183661733730
    'c1 = 1234567890.1234
    'c2 = 2345678901.2345
    
    '628239210217683 * 41419963890366 = 26021645401728484450406541978
    'c1 = 62823921021.7683
    'c2 = 4141996389.0366
    
    '123456789012345 * 234567890123456 = 28958998520042431946358064320
    'c1 = 12345678901.2345
    'c2 = 23456789012.3456
    
    '289012345678901 * 289012345678901 = 79228162514264300664528174397
    c1 = 28901234567.8901
    c2 = 27413418042.1097
    
    d = UInt64_Mul(c1, c2)
    Debug_Print VarType(d) & " " & CStr(d)
    
    's = Space(20): UInt64_ToHex cret, StrPtr(s)
    s = UInt64_ToHex(cret)
    Debug_Print Trim0(s)
    
    'cret = CCur("-922337203685477,5807")
    cret = CCur("-184467440737095")
    's = Space(66): UInt64_ToBin cret, StrPtr(s)
    s = UInt64_ToBin(cret)
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

'Public Function U8Mul(ByVal v1 As Currency, v2 As Currency) As Variant
'    UInt64_Mul v1, v2, U8Mul
'End Function
