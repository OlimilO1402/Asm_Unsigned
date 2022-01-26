VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
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

Private Declare Function Int32_UAdd_ref Lib "Int32UOps" (ByRef pV1 As Long, ByRef pV2 As Long) As Long
Private Declare Function Int32_UAdd Lib "Int32UOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function Int32_USubtract Lib "Int32UOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function Int32_UMultiply Lib "Int32UOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function Int32_UMultiplyB Lib "Int32UOps" Alias "Int32_UMultiply" (ByVal v1 As Long, ByVal v2 As Long) As Currency
Private Declare Function Int32_UDivide Lib "Int32UOps" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function Int32_UDivideB Lib "Int32UOps" Alias "Int32_UDivide" (ByVal v1 As Long, ByVal v2 As Long) As Currency

'just some short forms
Private Declare Function UAdd_ref Lib "Int32UOps" Alias "Int32_UAdd_ref" (ByRef pV1 As Long, ByRef pV2 As Long) As Long
Private Declare Function UAdd Lib "Int32UOps" Alias "Int32_UAdd" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function USub Lib "Int32UOps" Alias "Int32_USubtract" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UMul Lib "Int32UOps" Alias "Int32_UMultiply" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UMulB Lib "Int32UOps" Alias "Int32_UMultiply" (ByVal v1 As Long, ByVal v2 As Long) As Currency
Private Declare Function UDiv Lib "Int32UOps" Alias "Int32_UDivide" (ByVal v1 As Long, ByVal v2 As Long) As Long
Private Declare Function UDivB Lib "Int32UOps" Alias "Int32_UDivide" (ByVal v1 As Long, ByVal v2 As Long) As Currency

Private Declare Sub Int32_UToStr Lib "Int32UOps" (ByVal Value As Long, ByVal pStr_out As Long)
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
    Dim l As Single: l = 0
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move l, T, W, H
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
     
    lret = Int32_UAdd_ref(v1, v2)
    Debug_Print lret '155
    
    lret = Int32_UAdd(v1, v2)
    Debug_Print lret '155
    
    lret = Int32_USubtract(v1, v2)
    Debug_Print lret '91
    
    lret = Int32_UMultiply(v1, v2)
    Debug_Print lret '3936
    
    v1 = 2147483647
    v2 = 100
    cret = Int32_UMultiplyB(v1, v2)
    Debug_Print cret '21474836,4700
    
    v1 = 123456789
    v2 = 33
    lret = Int32_UDivide(v1, v2)
    Debug_Print lret '3741114
    
    v1 = -1
    s = Space(10)
    Int32_UToStr v1, StrPtr(s)
    Debug_Print Trim0(s)    '4294967295
    
    v1 = 32
    v2 = 65
    
    lret = UAdd(USub(100, UMul(2, UDiv(v1, 4))), UDiv(v2, UAdd(6, 7)))
    s = Space(10): Int32_UToStr lret, StrPtr(s)
    Debug_Print Trim0(s)
    
    lret = (100 - (2 * v1 / 4)) + (v2 / (6 + 7))
    s = Space(10): Int32_UToStr lret, StrPtr(s)
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
