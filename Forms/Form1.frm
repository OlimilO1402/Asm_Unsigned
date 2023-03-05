VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command8 
      Caption         =   "Conversion 16-32-64"
      Height          =   375
      Left            =   11520
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
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
    
    Dim l As Long
    sVal = "10101010101010101010101010101010"
    If Not MUnsigned.UInt32_TryParse(sVal, l, 2) Then Exit Sub
    s = s & "s: " & sVal & " = v: " & MUnsigned.UInt32_ToStr(l) & vbCrLf
    
    sVal = "1234567890"
    If Not MUnsigned.UInt32_TryParse(sVal, l, 10) Then Exit Sub
    s = s & "s: " & sVal & " = v: " & MUnsigned.UInt32_ToStr(l) & vbCrLf
    
    sVal = "ABCDEF98"
    If Not MUnsigned.UInt32_TryParse(sVal, l, 16) Then Exit Sub
    s = s & "s: " & sVal & " = v: " & MUnsigned.UInt32_ToStr(l) & vbCrLf
    
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

Private Sub Command8_Click()
    Dim i As Integer
    Dim l As Long
    Dim c As Currency
    Dim s As String
    i = 12345:
    s = s & "i = " & i & vbCrLf
    
    l = MUnsigned.UInt16_ToUInt32(i)
    s = s & "l = MUnsigned.UInt16_ToUInt32(i)" & vbCrLf
    s = s & "l = " & l & vbCrLf
    
    c = MUnsigned.UInt16_ToUInt64(i)
    s = s & "c = MUnsigned.UInt16_ToUInt64(i)" & vbCrLf
    s = s & "c = " & c & vbCrLf
    
    
    i = 0: c = 0
    l = 12345
    s = s & "l = " & l & vbCrLf
    
    i = MUnsigned.UInt32_ToUInt16(l)
    s = s & "i = MUnsigned.UInt32_ToUInt16(l)" & vbCrLf
    s = s & "i = " & i & vbCrLf
    
    l = 123456789
    s = s & "l = " & l & vbCrLf
    c = MUnsigned.UInt32_ToUInt64(l)
    s = s & "c = MUnsigned.UInt32_ToUInt64(l)" & vbCrLf
    s = s & "c = " & c & vbCrLf
    
    
    i = 0
    l = 0
    c = 1.2345
    s = s & "c = " & c & vbCrLf
    i = MUnsigned.UInt64_ToUInt16(c)
    s = s & "i = MUnsigned.UInt64_ToUInt16(c)" & vbCrLf
    s = s & "i = " & i & vbCrLf
    
    c = 123.4567
    s = s & "c = " & c & vbCrLf
    l = MUnsigned.UInt64_ToUInt32(c)
    s = s & "l = MUnsigned.UInt64_ToUInt32(c)" & vbCrLf
    s = s & "l = " & l & vbCrLf
    
    Text1.Text = s
End Sub

Private Sub Form_Load()
    Me.Caption = "Unsigned operations on signed In16+Int32+Int64 - v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
    Dim l As Single: l = 0
    Dim t As Single: t = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then Text1.Move l, t, W, H
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
    cret = UInt32_Mul(v1, v2)
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
    's = Space(10)
    s = UInt32_ToStr(v1)  ', StrPtr(s)
    Debug_Print s 'Trim0(s)    '4294967295
    
    v1 = -1
    's = Space(8)
    s = UInt32_ToHex(v1)  ', StrPtr(s)
    Debug_Print s 'Trim0(s)    '4294967295
    
    v1 = -1
    's = Space(32)
    s = UInt32_ToBin(v1) '1, StrPtr(s)
    Debug_Print s 'Trim0(s)    '4294967295
    
    v1 = 32
    v2 = 65
    
    lret = U4Add(U4Sub(100, U4Mul(2, U4Div(v1, 4))), U4Div(v2, U4Add(6, 7)))
    's = Space(10): UInt32_ToStr lret, StrPtr(s)
    s = UInt32_ToStr(lret)
    Debug_Print s
    
    lret = (100 - (2 * v1 / 4)) + (v2 / (6 + 7))
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
    
    v1 = &HCAFEBABE
    Dim d: d = UInt32_ToDec(v1)
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
    Dim l As Long
    
    v = 14021970
    r = 10
    l = LogN(v, r) + 1
    s = Space$(l)
    ret = Int32_ToStrR(v, StrPtr(s), r)
    Debug_Print s
    
    If ret <> l * 2 Then
        Debug_Print "error r<>l: r=" & r & "; l=" & l
    End If
    
    v = &H1CAFEBAB
    r = 16
    l = Ceiling(LogN(v, r))
    s = Space$(l)
    ret = Int32_ToStrR(v, StrPtr(s), r)
    
    Debug_Print s
    
    If ret <> l * 2 Then
        Debug_Print "error r<>l: r=" & r & "; l=" & l
    End If
    
    v = 987654321
    r = 10
    l = Ceiling(LogN(v, r))
    s = Space$(l)
    ret = Int32_ToStrR(v, StrPtr(s), r)
    
    Debug_Print s
    
    If ret <> l * 2 Then
        Debug_Print "error r<>l: r=" & r & "; l=" & l * 2
    End If
    
End Sub

'Public Function U8Mul(ByVal v1 As Currency, v2 As Currency) As Variant
'    UInt64_Mul v1, v2, U8Mul
'End Function
