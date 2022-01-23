VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
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

Private Declare Sub mov Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As Long)

'Private Declare Function Int32_UAdd_ref Lib "Int32UOps" (ByRef pV1 As Long, ByRef pV2 As Long) As Long

Private Declare Function Int32_UAdd Lib "Int32UOpsDll\Int32UOps" (ByVal v1 As Long, ByVal v2 As Long) As Long

Private Declare Function Int32_USubtract Lib "Int32UOpsDll\Int32UOps" (ByVal v1 As Long, ByVal v2 As Long) As Long

Private Declare Function Int32_UMultiply Lib "Int32UOpsDll\Int32UOps" (ByVal v1 As Long, ByVal v2 As Long) As Long

Private Declare Function Int32_UMultiplyB Lib "Int32UOpsDll\Int32UOps" Alias "Int32_UMultiply" (ByVal v1 As Long, ByVal v2 As Long) As Currency

Private Declare Function Int32_UDivide Lib "Int32UOpsDll\Int32UOps" (ByVal v1 As Long, ByVal v2 As Long) As Long

Private Declare Function Int32_UDivideB Lib "Int32UOpsDll\Int32UOps" Alias "Int32_UDivide" (ByVal v1 As Long, ByVal v2 As Long) As Currency

'Private Declare Function Int32_UToStr Lib "Int32UOps" (ByVal v As Long) As String

Private Declare Sub Int32_UToStr Lib "Int32UOpsDll\Int32UOps" (ByVal value As Long, ByVal pStr_out As Long) ', ByVal wcharlen As Long, ByVal radix As Long)
'
'errno_t _ultow_s(
'    unsigned long value,
'    wchar_t *str,
'    size_t sizeOfstr,
'    int radix
');

Private Sub Command1_Click()
    Text1.Text = ""
    TestDll
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Test_ToStr
End Sub

Sub Debug_Print(ByVal s As String)
    Text1.Text = Text1.Text & s & vbCrLf
End Sub
Sub TestDll()
    Dim v1 As Long: v1 = 123
    Dim v2 As Long: v2 = 32
    Dim lret As Long
    Dim cret As Currency
        
'    lret = Int32_UAdd_ref(v1, v2)
'    Debug.Print ret '155
    
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
    lret = Int32_UDivide(v1, v2) 'Laufzeitfehler 6: Überlauf
    Debug_Print lret '3741114
    
    v1 = -1
    Dim s As String: s = Space(10)
    Int32_UToStr v1, StrPtr(s)
    Debug_Print s    '4294967295

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
    
    
    Dim r As Long
    Dim l As Long
    
    v = 14021970
    r = 10         'radix
    l = LogN(v, r) + 1
    s = Space$(l)
    ret = Int32_ToStrR(v, StrPtr(s), r)
    Debug_Print s
    
    If ret <> l * 2 Then
        Debug_Print "error r<>l: r=" & r & "; l=" & l
    End If
    
    v = &H1CAFEBAB
    r = 16         'radix
    l = Ceiling(LogN(v, r))
    s = Space$(l)
    ret = Int32_ToStrR(v, StrPtr(s), r)
    
    Debug_Print s
    
    If ret <> l * 2 Then
        Debug_Print "error r<>l: r=" & r & "; l=" & l
    End If
    
    
    v = 987654321
    r = 10         'radix
    l = Ceiling(LogN(v, r))
    s = Space$(l)
    ret = Int32_ToStrR(v, StrPtr(s), r)
    
    Debug.Print s
    
    If ret <> l * 2 Then
        Debug_Print "error r<>l: r=" & r & "; l=" & l * 2
    End If
    
    
End Sub

Function Int32_ToStrR(ByVal v As Long, ByRef pStr_out As Long, Optional ByVal radix As Long = 10) As Long
    
    Dim sl As Long
    Dim ll As Long
    Dim m As Long
    
    'read the length of the buffer and save it in sl
    mov sl, ByVal pStr_out - 4, 4
    
    ll = Ceiling(LogN(v, radix))       'x87-FPU: FYL2X
    ll = ll * 2
    
    If sl < ll Then
        'insufficient buffer-size
        'EAX = 0
        Exit Function
    End If
    
    Dim i As Long
    For i = ll To 1 Step -2
        
        If v = 0 Then Exit For
        
        m = v Mod radix
        
        If m > 9 Then m = m + &H7 'the gap between "9" and "A"
        
        m = m + &H30
        
        mov ByVal pStr_out + i - 2, m, 2
        
        v = v \ radix
        
    Next
    
    Int32_ToStrR = ll
    
End Function


Function Int32_ToStr_idiv(ByVal v As Long, buffer() As Byte) As Long
    
    ReDim buffer(0 To 9) As Byte
    
    Dim d As Long
    d = 1000000000 '10^9
    
    Dim pos As Long ': pos = 0
    
    Dim i As Long
    Dim tmp As Long
    For i = 0 To 9
    
        If v >= d Or pos > 0 Then
            
            tmp = v \ d
            buffer(pos) = &H30 + tmp
            pos = pos + 1
            v = v - tmp * d
        
        End If
        
        d = d \ 10
        
    Next
    
    Int32_ToStr = pos
End Function

Function Int32_ToStr(ByVal v As Long) As String
    
    Dim radix As Long: radix = 10
    Dim u As Long: u = 19
    ReDim buffer(0 To u) As Byte
    Dim m As Long
    
    Dim i As Long
    For i = u - 1 To 0 Step -2
        
        If v <> 0 Then
        
            m = v Mod radix
            buffer(i) = m + &H30
            v = v \ radix
            
        Else
            buffer(i) = &H20
        End If
    Next
    
    Int32_ToStr = buffer
    
End Function

Public Function Log10(ByVal d As Double) As Double
    Log10 = VBA.Math.Log(d) / VBA.Math.Log(10)
End Function

Public Function LN(ByVal d As Double) As Double
  LN = VBA.Math.Log(d)
End Function

Public Function LogN(ByVal x As Double, _
                     Optional ByVal N As Double = 10#) As Double
                     'n darf nicht eins und nicht 0 sein
    LogN = VBA.Math.Log(x) / VBA.Math.Log(N)
End Function

Public Function Ceiling(ByVal A As Double) As Double
    Ceiling = CDbl(Int(A))
    If A <> 0 Then If Abs(Ceiling / A) <> 1 Then Ceiling = Ceiling + 1
End Function

'Function Int32_ToStrR(ByVal v As Long, ByVal radix As Long) As String
'
'    Static chars(0 To 15) As Byte
'    Dim u As Long: u = 19
'    ReDim buffer(0 To u) As Byte
'    Dim m As Long
'
'    Dim i As Long
'    If chars(0) = 0 Then
'        For i = 0 To 9
'            chars(i) = &H30 + i
'        Next
'        For i = 10 To 15
'            chars(i) = &H37 + i
'        Next
'        'Debug.Print StrConv(chars, vbUnicode)
'    End If
'
'    For i = u - 1 To 0 Step -2
'
'        If v <> 0 Then
'
'            m = v Mod radix
'            buffer(i) = chars(m)
'            v = v \ radix
'
'        Else
'            buffer(i) = &H20
'        End If
'    Next
'
'    Int32_ToStrR = buffer
'
'End Function
'

