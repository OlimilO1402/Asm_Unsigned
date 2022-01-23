Attribute VB_Name = "Midiv"
Option Explicit

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
    
    Int32_ToStr_idiv = pos
End Function

