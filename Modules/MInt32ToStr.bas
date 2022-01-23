Attribute VB_Name = "MInt32ToStr"
Option Explicit

Private Declare Sub mov Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As Long)
'

Function Int32_ToStrR(ByVal v As Long, ByRef pStr_out As Long, Optional ByVal radix As Long = 10) As Long
    
    'with radix
    
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

'http://www.activevb.de/tipps/vb6tipps/tipp0714.html
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


