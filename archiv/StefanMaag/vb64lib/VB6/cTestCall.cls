VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit


' COM Standard-Aufrufe,
' der Rückgabewert wird immer als letzter Parameter mittels Zeiger übergeben wird
' Lediglich Functions und Propertys besitzen diesen versteckten Parameter, eine Sub nicht
' nachkontrollieren kann man das in den TypLibs
'
' Public Function Blub() As Long
' wird zu
' HRESULT __stdcall Blub(Long _this, Long* retVal)
' eine Sub dagegen
' Public Sub Blub()
' wird zu
' HRESULT __stdcall Blub(Long _this)

' ProcNr = 0
Public Function CallBackProc(ByVal Par1 As Long, _
                             ByVal Par2 As Long, _
                             ByVal Par3 As Long) As Long
   
   Debug.Print " "
   Debug.Print "Par1 = "; Par1
   Debug.Print "Par2 = "; Par2
   Debug.Print "Par3 = "; Par3
   
   CallBackProc = (Par3 - Par2) * Par1
   
End Function


Public Function TT(ByVal X As Long, ByVal Y As Long) As Long
   TT = 1
End Function

