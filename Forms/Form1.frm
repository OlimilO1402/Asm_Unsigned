VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Test Int32_UToStr"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Int32UOps-dll"
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

Private Declare Function UInt32_Add_ref Lib "UnsignedOps" (ByRef pV1 As Long, ByRef pV2 As Long) As Long
Private Declare Function UInt32_Add Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_Sub Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_Mul Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_MulB Lib "UnsignedOps" Alias "UInt32_Mul" (ByVal v1 As Long, ByVal v2 As Long) As Currency
Private Declare Function UInt32_Div Lib "UnsignedOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UInt32_DivB Lib "UnsignedOps" Alias "UInt32_Div" (ByVal v1 As Long, ByVal v2 As Long) As Currency

Private Declare Function UInt64_Add Lib "UnsignedOps" (ByVal v1 As Currency, ByVal v2 As Currency) As Currency
Private Declare Function UInt64_Sub Lib "UnsignedOps" (ByVal v1 As Currency, ByVal v2 As Currency) As Currency

'just some short forms
Private Declare Function UAdd_ref Lib "UnsignedOps" Alias "UInt32_UAdd_ref" (ByRef pV1 As Long, ByRef pV2 As Long) As Long
Private Declare Function UAdd Lib "UnsignedOps" Alias "UInt32_Add" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function USub Lib "UnsignedOps" Alias "UInt32_Sub" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UMul Lib "UnsignedOps" Alias "UInt32_Mul" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UMulB Lib "UnsignedOps" Alias "UInt32_Mul" (ByVal v1 As Long, ByVal v2 As Long) As Currency
Private Declare Function UDiv Lib "UnsignedOps" Alias "UInt32_Div" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UDivB Lib "UnsignedOps" Alias "UInt32_Div" (ByVal v1 As Long, ByVal v2 As Long) As Currency

Private Declare Sub UInt32_ToStr Lib "UnsignedOps" (ByVal Value As Long, ByVal pStr_out As Long)
Private Declare Sub UInt64_ToStr Lib "UnsignedOps" (ByVal Value As Currency, ByVal pStr_out As Long)
'
'errno_t _ultow_s(
'    unsigned long value,
'    wchar_t *str,
'    size_t sizeOfstr,
'    int radix
');

Private Sub Form_Load()
    Me.Caption = "Unsigned ops on signed Int32 - v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = 0
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
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
    
    v1 = -1
    s = Space(10)
    UInt32_ToStr v1, StrPtr(s)
    Debug_Print Trim0(s)    '4294967295
    
    v1 = 32
    v2 = 65
    
    lret = UAdd(USub(100, UMul(2, UDiv(v1, 4))), UDiv(v2, UAdd(6, 7)))
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
