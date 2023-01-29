Attribute VB_Name = "vb64Bit"
' ***************************************************************************
'  NAME: vb64Bit
'  DESC: 64Bit Funktionen für die Erweiterung von VB
'  DESC: die Aufrufe bzw. Rückgabewerte sind für VB6 optimiert
'  DESC: Die DLL ist komplett in Asembler geschrieben
'  DESC:
'  DESC: Da VB keine direkte ByVal Übergabe von UDTs untersützt
'  DESC: (diese müssten zerlegt in einzelnen häppchen übergeben werden),
'  DESC: erhalten hier Aufrufe über Pointer den Vorzug. Damit liegen dann
'  DESC: die Pointer auf dem Stack, nicht die Daten.
'  DESC: Ein weiterer Vorteil dadurch für VB ist, dass je nach dem wie der
'  DESC: DECLARE Aufruf gestaltet wird, unterschielichste VB-Daten
'  DESC: übergeben werden können. Es könnten somit z.B. auch
'  DESC: Byte/Integer/Long-Arrays, Currency und andere direkt als Large
'  DESC: interpretiert werden.
'  DESC:
'  DESC: Mein Dank geht an Udo Schmidt, der viele Infos und auch entsprechend
'  DESC: Code beigesteuert hat.

'  Author :  Stefan Maag
'  CoAuthor: Udo Schmidt
'  Create : 18.11.2006
'  Change : 08.03.2007  Sub in Functions geändert (ab DLL-Version 0.2)
'  Change : 10.03.2007  Funktionen von Udo Schmidt hinzugefügt
' ***************************************************************************

' Die Funktionen in der DLL können auf unterschiedliche weise definiert werden.
' Drei Definitionen stehen hier zur Auswahl.
' 1. pOP As Large     : By Reference Definition für den Datentyp Large
' 2. pOP As Any       : By Referende Definition für irgendeine 64Bit Datenstruktur
' 3. ByVal pOP As Any : By Value Definition für die direkte Übergabe von Pointer Werten

' ACHTUNG: bei den Definitionen 2 und 3 muß der Programmierer selbst dafür
'          sorgen, dass die übergebenen Pointer auf mindestens 64Bit große
'          Daten zeigen, sonst sind Programmabstütze bzw. Fehlfunktionen
'          unvermeidlich.

Option Explicit

' Definition für 64Bit Large Integer
Public Type Large
   Lo As Long
   Hi As Long
End Type

' Compileranweisung nur für Test der Library auf True setzen, wenn sich die
' DLL im MASM-Projekt-Verzeichnis befindet statt im System32-Verzeichnis.
' Grund dafür ist, dass VB bei nicht kompletter Pfadangabe die DLL im System32 Verzeichnis erwartet
#Const Test = True

'#Const DeclareType = "ANY"      ' Deklariert die Parameter als ANY
'#Const DeclareType = "ByRef"    ' Deklariert die Paraneter als LARGE
'#Const DeclareType = "ByVal"    ' Deklariert die Parameter ByVal als Pointer-Werte

#If Test Then

Public Declare Function Add64 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (pOP1 As Large, pOP2 As Large) As Large

Public Declare Function Sub64 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (pOP1 As Large, pOP2 As Large) As Large

Public Declare Function Mul64 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (pOP1 As Large, pOP2 As Large) As Large

Public Declare Function Neg64 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (pOP1 As Large) As Large

Public Declare Function Not64 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (pOP1 As Large) As Large
         
Public Declare Function Shr64 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (pOP1 As Large, ByVal nBit As Long) As Large

Public Declare Function Shl64 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (pOP1 As Large, ByVal nBit As Long) As Large

Public Declare Function Div64 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (pOP1 As Large, pOP2 As Large, Optional ByVal signed As Boolean = True) As Large

Public Declare Function Mod64 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (pOP1 As Large, pOP2 As Large, Optional ByVal signed As Boolean = True) As Large

Public Declare Function StrToLarge Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal lpString As String, Optional ByVal fct As Long = 10, Optional overflow As Boolean) As Large

Public Declare Function LargeToStr Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (OP1 As Large, Optional ByVal fct As Long = 10, Optional ByVal signed As Long = 1) As String

Public Declare Function LargeToDec Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (OP1 As Large, Optional ByVal signed As Long = 1) As Variant

Public Declare Function VarToLarge Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (OP1 As Variant) As Large

Public Declare Function SqrLarge Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (OP1 As Large) As Large

#Else

' ===========================================================================
'  NAME: Add64
'  DESC: Addiert zwei 64 Bit Large Integer (signed oder unsigned)
'  PARA(pOP1):   Pointer Operand1 64Bit Large Integer (Speicherformat Lo, Hi)
'  PARA(pOP2):   Pointer Operand2 64Bit Large Integer (Speicherformat Lo, Hi)
'  RET:          Rückgabewert Large
'                [RET] = Operand1 + Operand2
' ===========================================================================
Public Declare Function Add64 Lib "vbExt.dll" (pOP1 As Large, pOP2 As Large) As Large
'Public Declare Function Add64 Lib "vbExt.dll" (pOP1 As Any, pOP2 As Any) As Large
'Public Declare Function Add64 Lib "vbExt.dll" (ByVal pOP1 As Long, ByVal pOP2 As Long) As Large

' ===========================================================================
'  NAME: Sub64
'  DESC: Subtrahiert zwei 64 Bit Large Integer (signed oder unsigned)
'  PARA(pOP1):   Pointer Operand1 64Bit Large Integer (Speicherformat Lo, Hi)
'  PARA(pOP2):   Pointer Operand2 64Bit Large Integer (Speicherformat Lo, Hi)
'  RET:          Rückgabewert Large
'                [RET] = Operand1 - Operand2
' ===========================================================================
Public Declare Function Sub64 Lib "vbExt.dll" (pOP1 As Large, pOP2 As Large) As Large
'Public Declare Function Sub64 Lib "vbExt.dll" (pOP1 As Any, pOP2 As Any) As Large
'Public Declare Function Sub64 Lib "vbExt.dll" (ByVal pOP1 As Long, ByVal pOP2 As Long) As Large

' ===========================================================================
'  NAME: Mul64
'  DESC: Multipliziert zwei 64 Bit Large Integer (signed oder unsigned)
'  PARA(pOP1):   Pointer Operand1 64Bit Large Integer (Speicherformat Lo, Hi)
'  PARA(pOP2):   Pointer Operand2 64Bit Large Integer (Speicherformat Lo, Hi)
'  RET:          Rückgabewert Large
'                [RET] = Operand1 * Operand2
' ===========================================================================
Public Declare Function Mul64 Lib "vbExt.dll" (pOP1 As Large, pOP2 As Large) As Large
'Public Declare Function Mul64 Lib "vbExt.dll" (pOP1 As Any, pOP2 As Any) As Large
'Public Declare Function Mul64 Lib "vbExt.dll" (ByVal pOP1 As Long, ByVal pOP2 As Long) As Large

' ===========================================================================
'  NAME: Neg64 (NegateLarge)
'  DESC: Negiert einen 64 Bit Large Integer; Zweierkomplement, Vorzeichenwechsel
'  PARA(pOP1):   Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
'  RET:          Rückgabewert Large
' ===========================================================================
Public Declare Function Neg64 Lib "vbExt.dll" (pOP1 As Large) As Large
'Public Declare Function Neg64 Lib "vbExt.dll" (pOP1 As Any) As Large
'Public Declare Function Neg64 Lib "vbExt.dll" (ByVal pOP1 As Long)

' ===========================================================================
'  NAME: Not64
'  DESC: Negiert einen 64 Bit Large Integer, Einerkomplement
'  PARA(pOP1):   Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
'  RET:          Rückgabewert Large
' ===========================================================================
Public Declare Function Not64 Lib "vbExt.dll" (pOP1 As Large) As Large
'Public Declare Function Not64 Lib "vbExt.dll" (pOP1 As Any) As Large
'Public Declare Function Not64 Lib "vbExt.dll" (ByVal pOP1 As Long) As Large

' ===========================================================================
'  NAME: Shr64 (ShiftRightLarge)
'  DESC: Schiebt einen 64 Bit Large Integer um Bits nach rechts
'  PARA(pOP1):   Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
'  RET:          Rückgabewert Large
' ===========================================================================
Public Declare Function Shr64 Lib "vbExt.dll" (pOP1 As Large, ByVal nBit As Long) As Large
'Public Declare Function Shr64 Lib "vbExt.dll" (pOP1 As Any, ByVal nBit As Long) As Large
'Public Declare Function Shr64 Lib "vbExt.dll" (ByVal pOP1 As Long, ByVal nBit As Long) As Large

' ===========================================================================
'  NAME: Shl64 (ShiftLeftLarge)
'  DESC: Schiebt einen 64 Bit Large Integer um Bits nach links
'  PARA(pOP1):    Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
'  RET:           Rückgabewert Large
' ===========================================================================
Public Declare Function Shl64 Lib "vbExt.dll" (pOP1 As Large, ByVal nBit As Long) As Large
'Public Declare Function Shl64 Lib "vbExt.dll" (pOP1 As Any, ByVal nBit As Long) As Large
'Public Declare Function Shl64 Lib "vbExt.dll" (ByVal pOP1 As Long, ByVal Long As Byte) As Large


' ===========================================================================
'  NAME: Div64
'  DESC: dividiert zwei 64 Bit Large Integer
'  PARA(pOP1):   Pointer Operand1 64Bit Divident
'  PARA(pOP2):   Pointer Operand2 64Bit Divisor
'  RET:          Rückgabewert Large
'                [RET] = Operand1 * Operand2
' ===========================================================================
Public Declare Function Div64 Lib "vbExt.dll" (pOP1 As Large, pOP2 As Large, Optional ByVal signed As Boolean = True) As Large
'Public Declare Function Div64 Lib "vbExt.dll" (pOP1 As Any, pOP2 As Any, Optional ByVal signed As Boolean = True) As Large
'Public Declare Function Div64 Lib "vbExt.dll" (ByVal pOP1 As Long, ByVal pOP2 As Long, Optional ByVal signed As Boolean = True) As Large

' ===========================================================================
'  NAME: Mod64
'  DESC: Modulo Division für 64 Bit Integer
'  PARA(pOP1):   Pointer Divident 64Bit Integer (Speicherformat Lo, Hi)
'  PARA(pOP2):   Pointer Divisor  64Bit Integer (Speicherformat Lo, Hi)
'  RET:          Rückgabewert Large
'                [Ret] = Divisionsrest (Remainder)
' ===========================================================================
Public Declare Function Mod64 Lib "vbExt.dll" (pOP1 As Large, pOP2 As Large, Optional ByVal signed As Boolean = True) As Large
'Public Declare Function Mod64 Lib "vbExt.dll" (pOP1 As Any, pOP2 As Any,Optional ByVal signed As Boolean = True) As Large
'Public Declare Function Mod64 Lib "vbExt.dll" (ByVal pOP1 As Long, ByVal pOP2 As Long, Optional ByVal signed As Boolean = True) As Large

' ===========================================================================
'  NAME: StrToLarge
'  DESC:         Wandelt eine String in eine 64Bit Integer
'  DESC:         Akzeptiert auch prefixes &B, &O, &D and &H
'  PARA(lpStr):  VB-String
'  PARA(fct):    Faktor für Umwandlung bzw. Zahlensystem des String
'  RET:          Rückgabewert Large
' ===========================================================================

Public Declare Function StrToLarge Lib "vbExt.dll" (ByVal lpStr As String, Optional ByVal fct As Long = 10, Optional overflow As Boolean) As Large

' ===========================================================================
'  NAME: LargeToStr
'  DESC:         Wandelt einen 64Bit Integer in einen String
'  PARA(OP1):    Operand Large
'  PARA(fct):    Faktor für Umwandlung bzw. Zahlensystem des String
'  RET:          String
' ===========================================================================

Public Declare Function LargeToStr Lib "vbExt.dll" (OP1 As Large, Optional ByVal fct As Long = 10, Optional ByVal signed As Long = 1) As String

' ===========================================================================
'  NAME: LargeToDec
'  DESC:         Wandelt einen 64Bit Integer in einen Decimal (Variant)
'  PARA(OP1):    Operand Large
'  PARA(signed): mit Vorzeichen?
'  RET:          Variant, Decimal
' ===========================================================================

Public Declare Function LargeToDec Lib "vbExt.dll" (OP1 As Large, Optional ByVal signed As Long = 1) As Variant

' ===========================================================================
'  NAME: VarToLarge
'  DESC:         Wandelt einen Variant in einen 64Bit Large Integer
'  PARA(OP1):    Operand Variant
'                Akzeptiert int, lng, sgl, dbl, cur, dat, str, bol, dec, byte
'  RET:          64Bit Large Integer
' ===========================================================================

Public Declare Function VarToLarge Lib "vbExt.dll" (OP1 As Variant) As Large

' ===========================================================================
'  NAME: SqrLarge
'  DESC:         Berechnet die Quadratwurzel eines 64Bit Large Integer
'  PARA(OP1):    Operand Large
'  RET:          64Bit Large Integer
' ===========================================================================

Public Declare Function SqrLarge Lib "vbExt.dll" (OP1 As Large) As Large

#End If



