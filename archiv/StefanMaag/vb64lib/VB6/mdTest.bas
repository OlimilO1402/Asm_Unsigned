Attribute VB_Name = "mdTest"
Option Explicit

' Diese Teststrukt wird als Rückgabewert für Functions verwendet, um
' zu testen, ob die Pointerübergaben des Rückgabewertes funktionieren
Public Type TestStrukt
   p1 As Long
   p2 As Long
   p3 As Long
   p4 As Long
End Type


Public Function TestCall(ByVal Par1 As Long, ByVal Par2 As Long, ByVal Par3 As Long) As Long
   Debug.Print " "
   Debug.Print "Par1 = "; Par1
   Debug.Print "Par2 = "; Par2
   Debug.Print "Par3 = "; Par3
   
   TestCall = (Par3 - Par2) * Par1
   
End Function

' TestFunction zum Erzeugen einer TestStukt, um herauszufinden, wie
' VB den Rückgabeparameter handhabt

Public Function CreateStrukt(ByVal w1 As Long, _
                              ByVal w2 As Long, _
                              ByVal w3 As Long, _
                              ByVal w4 As Long) As TestStrukt
      With CreateStrukt
         .p1 = w1
         .p2 = w2
         .p3 = w3
         .p4 = w4
      End With
      
End Function

Public Sub Main()
   Dim cTS As New cTestCall
   
   Dim rs As TestStrukt
   Dim I As Long
   Dim J As Long
   
   I = &H1111
   I = &H1111
   I = &H2222
   I = &H2222
   
   J = cTS.TT(1, 2)
   'rs = CreateStrukt(1, 2, 3, 4)
   
   I = &H3333
   I = &H3333
   I = &H4444
   I = &H4444
End Sub

Public Function TT(ByVal X As Long, ByVal Y As Long) As TestStrukt
   TT.p1 = 1
   TT.p2 = 2
   TT.p3 = 3
   TT.p4 = 4
End Function

Public Function LT(ByVal Wert As Long) As Long
   LT = 1
End Function


