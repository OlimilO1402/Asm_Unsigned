Attribute VB_Name = "vbExt"
' ***************************************************************************
'  NAME: vbExt.dll
'  DESC: Erweiterungsfunktionen für VB komplett in Assembler geschrieben
'  DESC: Funktionen, die VB nicht bietet, aber immer wieder mal benötig
'  DESC: werden. Dazu gehören vor allem Bit-Operationen wie GetBit, SetBit
'  DESC: Schieboperationen SHL, SHR sowie Bitrotationsfunktionen ROL, ROR
'  DESC: Byte/Word Funktionen Get-/Set- Lo/Hi
'  DESC: Byte/Word Anordnung vertauschen zur Konvertierung zwischen den
'  DESC: Formaten Little- und Big-Endian
'  DESC: 'Hardcore'-Funktionen für das Array-Handling
'
'  Author : Stefan Maag
'  Create : 28.03.2007
'  Change : 11.11.2007
'  Change :
' ***************************************************************************


Option Explicit

' Compileranweisung nur für Test der Library auf True setzen, wenn sich die
' DLL im MASM-Projekt-Verzeichnis befindet statt im System32-Verzeichnis.
' Grund dafür ist, dass VB bei nicht kompletter Pfadangabe die DLL im System32 Verzeichnis erwartet
#Const Test = True
'
#If Test Then

Public Declare Function VbTest Lib "C:\masm32\Project\vbExt\vbExt.dll" _
               (ByVal Par1 As Long, ByVal Par2 As Long) As Long


Public Declare Function VbStackParVal Lib "C:\masm32\Project\vbExt\vbExt.dll" Alias "VbStackPar" _
                (ByVal ptr As Long) As Long

Public Declare Function CallProc Lib "C:\masm32\Project\vbExt\vbExt.dll" _
               (ptrParList As Any, _
                ByVal nPar As Long, _
                ByVal ptrProc As Long, _
                Optional ByVal ptrRet As Long = 0) As Long

Public Declare Function CallObjProc Lib "C:\masm32\Project\vbExt\vbExt.dll" _
               (ptrParList As Any, _
                ByVal nPar As Long, _
                ByVal ptrObj As Long, _
                ByVal ProcNr As Long, _
                Optional ByVal StdCall As Boolean = True) As Long

Public Declare Function GetObjProcPtr Lib "C:\masm32\Project\vbExt\vbExt.dll" _
               (ByVal ptrObj As Long, ByVal ProcNr As Long) As Long

Public Declare Function VbStackPar Lib "C:\masm32\Project\vbExt\vbExt.dll" _
               (vbPtr As Any) As Long

Public Declare Function IsArrayDim Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (vbArray() As Any) As Boolean

Public Declare Function ptrArrayStrukt Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (vbArray() As Any) As Long

Public Declare Sub XchgArray Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (vbArray1() As Any, vbArray2() As Any)
         
Public Declare Function UnHookArray Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (vbArray() As Any) As Long
         
Public Declare Function HookArray Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (vbArray() As Any, ByVal pSFA As Long) As Long
         
Public Declare Function GetBit16 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer, ByVal BitPos As Long) As Boolean
         
Public Declare Function GetBit32 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long, ByVal BitPos As Long) As Boolean
         
Public Declare Function SetBit16 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer, ByVal BitPos As Long) As Integer

Public Declare Function SetBit32 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long, ByVal BitPos As Long) As Long
         
Public Declare Function InvBit16 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer, ByVal BitPos As Long) As Integer
         
Public Declare Function InvBit32 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long, ByVal BitPos As Long) As Long
         
Public Declare Function SHL16 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer, ByVal Shift As Long) As Integer
         
Public Declare Function SHL32 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long, ByVal Shift As Long) As Long
         
Public Declare Function SHR16 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer, ByVal Shift As Long) As Integer
         
Public Declare Function SHR32 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long, ByVal Shift As Long) As Long
         
Public Declare Function ROL16 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer, ByVal rotate As Long) As Integer
         
Public Declare Function ROL32 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long, ByVal rotate As Long) As Long
         
Public Declare Function ROLR16 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer, ByVal rotate As Long) As Integer
         
Public Declare Function ROR32 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long, ByVal rotate As Long) As Long
        
Public Declare Function LSB16 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer) As Integer
       
Public Declare Function LSB32 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long) As Long
Public Declare Function MSB16 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer) As Integer
       
Public Declare Function MSB32 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long) As Long

Public Declare Function BitCount16 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer) As Integer
     
Public Declare Function BitCount32 Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long) As Long

Public Declare Function GetBitBlock Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long, ByVal FirstBit As Integer, ByVal LastBit As Integer) As Long
    
Public Declare Function GetLoByte Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer) As Byte
         
Public Declare Function GetHiByte Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer) As Byte
        
Public Declare Function SetLoByte Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer, ByVal NewLo As Byte) As Integer

Public Declare Function SetHiByte Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Integer, ByVal NewHi As Byte) As Integer
       
Public Declare Function GetLoWord Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long) As Integer

Public Declare Function GetHiWord Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long) As Integer

Public Declare Function SetLoWord Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long, ByVal NewLo As Integer) As Long

Public Declare Function SetHiWord Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long, ByVal NewHi As Integer) As Long

Public Declare Function ByteSwap Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long) As Long

Public Declare Function WordSwap Lib "C:\masm32\Project\vbExt\vbExt.dll" _
         (ByVal Wert As Long) As Long

 #Else  ' ENDE DER DEFINITIONEN FÜR DEN TEST


' ===========================================================================
'  NAME: VbStackPar
'  DESC: Gibt die von VB auf den Stack gelegte Adress bzw. Wert zurück
'  DESC: Damit kann man überprüfen, was VB in gewissen Fällen wirklich als
'  DESC: Paramter auf den Stack legt
'
'  PARA(VbAny):  VbAny-Parameter
'  RET(EAX):     Wert, der von VB auf den Stack gelegt wurde
' ===========================================================================

Public Declare Function VbStackParVal Lib "vbExt.dll" Alias "VbStackPar" _
                (ByVal ptr As Long) As Long


' ===========================================================================
' NAME: CallProc
' DESC: aufrufen beliebiger Proceduren über deren Adresse
' DESC: die aufzurufende Procedur muss aber im Adressraum des
' DESC: eigenen Programms liegen
'
' PARA(ptrParLst): Pointer der Parameterliste
' PARA(nPar)     : Anzahl der zu übergebenden Parameter
' PARA(ptrProc)  : Proceduradresse
' RET            : Rückgabewert der aufgerufenen Procedur
' ===========================================================================

Public Declare Function CallProc Lib "vbExt.dll" _
               (ptrParList As Any, _
                ByVal nPar As Long, _
                ByVal ptrProc As Long, _
                Optional ByVal ptrRet As Long = 0) As Long

' ===========================================================================
' NAME: CallObjProc
' DESC: aufrufen beliebiger Proceduren in Objecten über deren Adresse
' DESC: die aufzurufende Procedur muss aber im Adressraum des
' DESC: eigenen Programms liegen. Die Procedureadresse im Object ist
' DESC: vorher mit GetObjProcPtr anhand der Procedur-Nr. zu ermitteln
'
' PARA(ptrParLst): Pointer der Parameterliste
' PARA(nPar)     : Anzahl der zu übergebenden Parameter
' PARA(ptrProc)  : Proceduradresse
' RET            : Rückgabewert der aufgerufenen Procedur
' ===========================================================================

Public Declare Function CallObjProc Lib "vbExt.dll" _
               (ptrParList As Any, _
                ByVal nPar As Long, _
                ByVal ptrObj As Long, _
                ByVal ProcNr As Long, _
                Optional ByVal StdCall As Boolean = True) As Long

' ===========================================================================
' NAME: GetObjProcPtr
' DESC: ermittelt die Proceduradresse einer öffentlichen Procedur eines Objects
' DESC: anhand der Procedur-Nummer
'
' PARA(ptrObj): Object-Pointer
' PARA(ProcNr): Procedurnummer, fortlaufend, beginnend mit 0
' RET():        Proceduradresse
' ===========================================================================
Public Declare Function GetObjProcPtr Lib "vbExt.dll" _
               (ByVal ptrObj As Long, ByVal ProcNr As Long) As Long

' ===========================================================================
' NAME: VBStackPar
' DESC: Gibt die von VB auf den Stack gelegte Adress bzw. Wert zurück
' DESC: Damit kann man überprüfen, was VB in gewissen Fällen wirklich als
' DESC: Paramter auf den Stack legt
'
' PARA(VbAny):  VbAny-Parameter
' RET(EAX):     Wert, der von VB auf den Stack gelegt wurde
' ===========================================================================
Public Declare Function VbStackPar Lib "vbExt.dll" _
               (vbPtr As Any) As Long


' ===========================================================================
' NAME: IsArrayDim
' DESC: Ermittelt ob ein Array (SafeArray) dimensioniert ist
'
' PARA(VbArray): Array-Pointer
' RET(EAX):      vbFalse = undimensioniert; vbTrue=dimensioniert
' ===========================================================================

Public Declare Function IsArrayDim Lib "vbExt.dll" _
         (vbArray() As Any) As Boolean

' ===========================================================================
' NAME: PtrArrayStrukt
' DESC: Ermittelt den Pointer der zugehörigen SafeArrayStruktur des Arrays
' DESC: und gibt diesen zurück. Um den Pointer der SafeArrayStruktur zu setzen,
' DESC: die Funkion HookArray (hängt SafeArrayStruktur ein), verwenden!
'
' PARA(Array):
' RET:  Pointer auf SafeArrayStruktur
' ===========================================================================

Public Declare Function ptrArrayStrukt Lib "vbExt.dll" _
         (vbArray() As Any) As Long

' ===========================================================================
' NAME: XchgArray
' DESC: Tauscht 2 Arrays gegeneinander aus.
' DESC: es werden nur die Poitner der SafearrayStrukturen getauscht
' DESC: ACHTUNG! Keine normalen VB-Variablen übergeben, dies führt
' DESC: unweigerlich zu Speicherverletzungen
' PARA(Array1): Array1
' PARA(Array2): Array2
' ===========================================================================

Public Declare Sub XchgArray Lib "vbExt.dll" _
         (vbArray1() As Any, vbArray2() As Any)
         
' ===========================================================================
' NAME: UnHookArray
' DESC: Hängt die SafeArray-Struktur des Arrays aus, d.h. der Pointer
' DESC: auf die zugehörige SafeArray-Struktur wird gelöscht (:=0)
' DESC: Der Array ist somit undimensionert.
' DESC: Der ursprüngliche SafeArrayPointer wird als Funktionrückgabewert
' DESC: zurückgegeben.
' DESC: ACHTUNG: Wenn der Array ausgehängt wird, hängt der reservierte
' DESC: Speicherbereich in der Luft - VB weis davon im Prinzip nichts.
' DESC: Vor dem beenden des Programms bzw. vor auflösen des Arrays
' DESC: sollte der ursprüngliche Zustand wieder hergestellt werden.
'
' PARA(Array): Array
' RET:  Pointer der urspünglichen SafeArray-Struktur
' ===========================================================================
         
Public Declare Function UnHookArray Lib "vbExt.dll" _
         (vbArray() As Any) As Long
         
' ===========================================================================
'  NAME: HookArray
'  DESC: Hängt eine SafeArray-Struktur in den Array. Die einzuhängende
'  DESC: SafeArrayStruktur wird mit ihrem Pointer angegeben.
'  DESC: Der Pointer der ursprünglichen SafeArrray-Struktur wird als
'  DESC: Fuktionswert zurückgegeben
'  DESC: ACHTUNG: VB bekommt davon nichts mit. Vor beenden des Programms
'  DESC: bzw. vor dem Auflösen der Arrays sollte der ursprüngliche Zustand
'  DESC: wieder hergestellt werden.
'
'  PARA(Array): Array
'  PARA(pSFA):  Pointer der neuen SafeArray-Struktur
'  RET:  Pointer der urspünglichen SafeArray-Struktur
' ===========================================================================
         
Public Declare Function HookArray Lib "vbExt.dll" _
         (vbArray() As Any, ByVal pSFA As Long) As Long
         
         
' ===========================================================================
' NAME: GetBit
' DESC: Ermittelt den Wert eines Bits in einem Wert
' DESC: Die Funktion exisitiert als 16/32 Bit Version
'
' PARA(WERT):   Quellwert
' PARA(BitPos): Position des zu prüfenden Bits (0..15)
' RET:          Rückgabewert:  vbFalse(0), wenn Bit = 0
'                            vbTrue(-1), wenn Bit = 1
' ===========================================================================
         
Public Declare Function GetBit16 Lib "vbExt.dll" _
         (ByVal Wert As Integer, ByVal BitPos As Long) As Boolean
         
Public Declare Function GetBit32 Lib "vbExt.dll" _
         (ByVal Wert As Long, ByVal BitPos As Long) As Boolean
         
' ===========================================================================
' NAME: SetBit
' DESC: Setzt den Wert eines Bits nach Vorgabe auf True oder False
' DESC: Die Funktion exisitiert als 16/32 Bit Version
'
' PARA(WERT):   Quellwert
' PARA(Pos):    Position des zu prüfenden Bits
' PARA(NewVal): Neuer Wert des zu bearbeitenden Bits (VbTrue/VbFalse)
' RET:          Rückgabewert:  WERT mit entsprechend bearbeitetem Bit
' ===========================================================================

Public Declare Function SetBit16 Lib "vbExt.dll" _
         (ByVal Wert As Integer, ByVal BitPos As Long) As Integer

Public Declare Function SetBit32 Lib "vbExt.dll" _
         (ByVal Wert As Long, ByVal BitPos As Long) As Long
         
' ===========================================================================
' NAME: InvBit
' DESC: Invertiert den Wert eines angegeben Bits
' DESC: Die Funktion exisitiert als 16/32 Bit Version
'
' PARA(WERT):   Quellwert
' PARA(Pos):    Position des zu invertierenden Bits (0..x)
' PARA(NewVal): Neuer Wert des zu bearbeitenden Bits (VbTrue/VbFalse)
' RET:          Rückgabewert:  WERT mit entsprechend bearbeitetem Bit
' ===========================================================================

Public Declare Function InvBit16 Lib "vbExt.dll" _
         (ByVal Wert As Integer, ByVal BitPos As Long) As Integer
         
Public Declare Function InvBit32 Lib "vbExt.dll" _
         (ByVal Wert As Long, ByVal BitPos As Long) As Long
         
' ===========================================================================
' NAME: SHL
' DESC: Schieben links
' DESC: Die Funktion exisitiert als 16/32 Bit Version
'
' PARA(WERT):   Quellwert
' PARA(Shift):  Anzahl zu schiebender Bits
' RET:          Rückgabewert:  WERT entprechende Bits verschoben
' ===========================================================================

Public Declare Function SHL16 Lib "vbExt.dll" _
         (ByVal Wert As Integer, ByVal Shift As Long) As Integer
         
Public Declare Function SHL32 Lib "vbExt.dll" _
         (ByVal Wert As Long, ByVal Shift As Long) As Long
         
' ===========================================================================
' NAME: SHR
' DESC: Schieben rechts
' DESC: Die Funktion exisitiert als 16/32 Bit Version
'
' PARA(WERT):   Quellwert
' PARA(Shift):  Anzahl zu schiebender Bits
' RET:          Rückgabewert:  WERT entprechende Bits verschoben
' ===========================================================================

Public Declare Function SHR16 Lib "vbExt.dll" _
         (ByVal Wert As Integer, ByVal Shift As Long) As Integer
         
Public Declare Function SHR32 Lib "vbExt.dll" _
         (ByVal Wert As Long, ByVal Shift As Long) As Long
         
' ===========================================================================
' NAME: ROL
' DESC: Rotieren links
' DESC: Die Funktion exisitiert als 16/32 Bit Version
'
' PARA(WERT):   Quellwert
' PARA(rotate): Anzahl zu rotierender Bits
' RET:          Rückgabewert:  WERT entprechende Bits rotiert
' ===========================================================================

Public Declare Function ROL16 Lib "vbExt.dll" _
         (ByVal Wert As Integer, ByVal rotate As Long) As Integer
         
Public Declare Function ROL32 Lib "vbExt.dll" _
         (ByVal Wert As Long, ByVal rotate As Long) As Long
         
' ===========================================================================
' NAME: ROR
' DESC: Rotieren rechts
' DESC: Die Funktion exisitiert als 16/32 Bit Version
'
' PARA(WERT):   Quellwert
' PARA(rotate): Anzahl zu rotierender Bits
' RET:          Rückgabewert:  WERT entprechende Bits rotiert
' ===========================================================================
         
Public Declare Function ROLR16 Lib "vbExt.dll" _
         (ByVal Wert As Integer, ByVal rotate As Long) As Integer
         
Public Declare Function ROR32 Lib "vbExt.dll" _
         (ByVal Wert As Long, ByVal rotate As Long) As Long
        
        
' ===========================================================================
' NAME: LSB, least significant Bit
' DESC: sucht das niederwertigste Bit und gibt dessen Position zurück
' DESC: Die Funktion exisitiert als 16/32 Bit Version
'
' PARA(WERT):   Quellwert
' RET():        Rückgabewert:  Bit-Index
' ===========================================================================

Public Declare Function LSB16 Lib "vbExt.dll" _
         (ByVal Wert As Integer) As Integer
       
Public Declare Function LSB32 Lib "vbExt.dll" _
         (ByVal Wert As Long) As Long
        
' ===========================================================================
' NAME: MSB, most significant Bit
' DESC: sucht das hoechstwertigste Bit und gibt dessen Position zurück
' DESC: Die Funktion exisitiert als 16/32 Bit Version
'
' PARA(WERT):   Quellwert
' RET():        Rückgabewert:  Bit-Index
' ===========================================================================

Public Declare Function MSB16 Lib "vbExt.dll" _
         (ByVal Wert As Integer) As Integer
       
Public Declare Function MSB32 Lib "vbExt.dll" _
         (ByVal Wert As Long) As Long
        
' ===========================================================================
'  NAME: BitCount, Bits zählen
'  DESC: ermittelt die Anzahl der gesetzten Bits (=1) in einem Wert
'  DESC: Die Funktion exisitiert als 16/32 Bit Version
'
'  PARA(WERT):   Quellwert
'  RET():        Rückgabewert:  Anzahl der gesetzten Bits
' ===========================================================================
Public Declare Function BitCount16 Lib "vbExt.dll" _
         (ByVal Wert As Integer) As Integer
     
Public Declare Function BitCount32 Lib "vbExt.dll" _
         (ByVal Wert As Long) As Long
     
' ===========================================================================
'  NAME: GetBitBlock
'  DESC: gibt den Wert eines Bitbereichs (Block) in eiem 32-Bit-Wert zurueck
'  DESC: Value = 0000 1001 1100 0110
'  DESC; der Wert des BitBereichs 0 bis 3 ist '0110' also 6
'
'  PARA(WERT):   Quellwert 32 Bit
'  PARA(FirstBit): Nr. des ersten Bit (0-basiert)
'  PARA(LastBit):  Nr. des letzten Bits (0-basiert)
'  RET:          Rückgabewert: gewünschter BitBereich
' ===========================================================================
'
'   Maskierungswert für &-Verknüpfung:
'   0xFFFFFFFF enthält nur 1en, diese werden um die benötigen Bits links verschoben,
'   nach invertieren dieses Wertes ergeben sich rechts die benötigten 1en
'   Value muss nun noch um StartBit nach rechts verschoben werden und mit dem
'   Maskierungswert &-verknüpft, das ist dann der gesuchte Wert der
'   BitArea(StartBit bis EndBit)

Public Declare Function GetBitBlock Lib "vbExt.dll" _
         (ByVal Wert As Long, ByVal FirstBit As Integer, ByVal LastBit As Integer) As Long
        
' ===========================================================================
' NAME: GetLoByte
' DESC: gibt das niederwertige Byte eines 16 Bit Wertes zurück
'
' PARA(WERT):   Quellwert 16 Bit
' RET:          Rückgabewert: niederwertiges Byte
' ===========================================================================

Public Declare Function GetLoByte Lib "vbExt.dll" _
         (ByVal Wert As Integer) As Byte
         
' ===========================================================================
' NAME: GetHiByte
' DESC: gibt das hoeherwertige Byte eines 16 Bit Wertes zurück
'
' PARA(WERT):   Quellwert 16 Bit
' RET:          Rückgabewert: hoeherwertiges Byte
' ===========================================================================

Public Declare Function GetHiByte Lib "vbExt.dll" _
         (ByVal Wert As Integer) As Byte
        
' ===========================================================================
' NAME: SetLoByte
' DESC: setzt das niederwertige Byte eines 16Bit Wertes auf den
' DESC: übergeben Wert
'
' PARA(WERT):   Quellwert 16 Bit
' PARA(NewLo):  Neuer Wert des LoBytes, 8Bit
' RET:          neuer 16Bit-Wert
' ===========================================================================
        
Public Declare Function SetLoByte Lib "vbExt.dll" _
         (ByVal Wert As Integer, ByVal NewLo As Byte) As Integer
        
' ===========================================================================
' NAME: SetHiByte
' DESC: setzt das hoeherwertige Byte eines 16Bit Wertes auf den
' DESC: übergeben Wert
'
' PARA(WERT):   Quellwert 16 Bit
' PARA(NewHi):  Neuer Wert des HiBytes, 8Bit
' RET:          neuer 16Bit-Wert
' ===========================================================================

Public Declare Function SetHiByte Lib "vbExt.dll" _
         (ByVal Wert As Integer, ByVal NewHi As Byte) As Integer
       
' ===========================================================================
' NAME: GetLoWord
' DESC: gibt das niederwertige Word eines 32 Bit Wertes zurück
'
' PARA(WERT):   Quellwert 32 Bit
' RET:          Rückgabewert: niederwertiges Word
' ===========================================================================

Public Declare Function GetLoWord Lib "vbExt.dll" _
         (ByVal Wert As Long) As Integer

' ===========================================================================
' NAME: GetHiWord
' DESC: gibt das hoeherwertige Word eines 32 Bit Wertes zurück
'
' PARA(WERT):   Quellwert 32 Bit
' RET:          Rückgabewert: hoeherwertiges Word
' ===========================================================================

Public Declare Function GetHiWord Lib "vbExt.dll" _
         (ByVal Wert As Long) As Integer

' ===========================================================================
' NAME: SetLoWord
' DESC: setzt das niederwertige Word eines 32Bit Wertes auf den
' DESC: übergeben Wert
'
' PARA(WERT):   Quellwert 32 Bit
' PARA(NewLo):  Neuer Wert des LoWords, 16Bit
' RET:          neuer 32 Bit-Wert
' ===========================================================================

Public Declare Function SetLoWord Lib "vbExt.dll" _
         (ByVal Wert As Long, ByVal NewLo As Integer) As Long

' ===========================================================================
' NAME: SetHiWord
' DESC: setzt das hoeherwertige Word eines 32Bit Wertes auf den
' DESC: übergeben Wert
'
' PARA(WERT):   Quellwert 32 Bit
' PARA(NewLo):  Neuer Wert des HiWords, 16Bit
' RET:          neuer 32 Bit-Wert
' ===========================================================================

Public Declare Function SetHiWord Lib "vbExt.dll" _
         (ByVal Wert As Long, ByVal NewHi As Integer) As Long

' ===========================================================================
' NAME: ByteSwap
' DESC: vertauscht die Bytereihenfolgen eins 32 Bit Wertes
' DESC: Konvertierung zwischen den Formaten Little- und Big-Endian
'
' PARA(WERT):   Quellwert 32 Bit
' RET:          neuer 32 Bit-Wert mit vertauschter Bytereihenfolge
' ===========================================================================

Public Declare Function ByteSwap Lib "vbExt.dll" _
         (ByVal Wert As Long) As Long

' ===========================================================================
' NAME: WordSwap
' DESC: vertauscht die Hi- und Lo-Word eins 32 Bit Wertes
'
' PARA(WERT):   Quellwert 32 Bit
' RET:          neuer 32 Bit-Wert mit vertauschter Wordreihenfolge
' ===========================================================================

Public Declare Function WordSwap Lib "vbExt.dll" _
         (ByVal Wert As Long) As Long

#End If

