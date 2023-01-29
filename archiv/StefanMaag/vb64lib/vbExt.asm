;Virtual Alloc
;http://www.activevb.de/cgi-bin/forenarchive/forenarchive.pl?a=0&b=b&d=99473&e=12&f=search&g=ausf%FChrbarer+speicher&h=1

; ***************************************************************************
;  NAME: vbExt.dll
;  DESC: 64Bit Funktionen für die Erweiterung von VB
;  DESC: Die Aufrufe bzw. Rückgabewerte sind für VB6 optimiert
;  DESC:
;  DESC: Da VB keine direkte ByVal Übergabe von UDTs untersützt 
;  DESC: (diese müssten zerlegt in einzelnen Häppchen übergeben werden),
;  DESC: erhalten hier Aufrufe über Pointer den Vorzug. Damit liegen dann
;  DESC: die Pointer auf dem Stack, nicht die Daten.
;  DESC: Ein weiterer Vorteil dadurch für VB ist, dass je nach dem wie der 
;  DESC: DECLARE Aufruf gestaltet wird, unterschielichste VB-Daten
;  DESC: übergeben werden können. Es können somit z.B. auch 
;  DESC: Byte/Integer/Long-Arrays, Currency und andere als LARGE
;  DESC: interpretiert werden.
;
;  LANGUAGE: 32Bit x86 Assembler (MASM32) 
;  TABSIZE = 4, Monospaced Font
; ***************************************************************************
;
;     Author   : Stefan Maag
;     CoAuthor : Udo Schmidt

;     Create : 15.11.2006
;     Change : 08.03.2007  Rückgabewerte über Register EAX/EDX, so dass
;                          VB Functions verwendet werden können
;     Change : 10.03.2007  Code von Udo Schmidt hinzugefügt
;     Change : 11.11.2007  16/32 Bit-Operationen 
           
;     Version: 0.3

; ===========================================================================
;  rem: Quellenangaben:
; ===========================================================================
;  rem: Assembler Profireferenz, Franzis Verlag
;  rem:
;  rem: AMD Athlon Processor - x86 Code Optimization Guide
;  rem:     Publication No: 22007  Revision: K  Date: February 2002; AMD     
;  rem:     (Efficient 64-Bit Integer Arithmetic)  

; ===========================================================================
;  rem: Weitere benötigte Dateien:
; ===========================================================================
;  rem:   keine

.686
.model flat, stdcall
option casemap:none

;     include files
;     ~~~~~~~~~~~~~
     include \masm32\include\windows.inc

; ----------------------------------------
; prototypes for local procedures go here
; ----------------------------------------

.data?
    hInstance dd ?
    tmpString       db 70 dup (?)                   ; BSTR-hdr + up to 64 binary digits + 0
.code

LibMain proc instance:dword,reason:dword,unused:dword 

    .if reason == DLL_PROCESS_ATTACH
      push instance
      pop hInstance

    .elseif reason == DLL_PROCESS_DETACH

    .elseif reason == DLL_THREAD_ATTACH

    .elseif reason == DLL_THREAD_DETACH

    .endif

    mov eax, TRUE
    ret

LibMain endp


; ===========================================================================
;  NAME: VbStackPar
;  DESC: Gibt die von VB auf den Stack gelegte Adress bzw. Wert zurück 
;  DESC: Damit kann man überprüfen, was VB in gewissen Fällen wirklich als
;  DESC: Paramter auf den Stack legt
;
;  PARA(VbAny):  VbAny-Parameter
;  RET(EAX):     Wert, der von VB auf den Stack gelegt wurde
; ===========================================================================

VbStackPar proc
    mov eax, [esp+4]
    ret 4
VbStackPar endp

; ===========================================================================
; NAME: CallProc
; DESC: aufrufen beliebiger Proceduren über deren Adresse
; DESC: die aufzurufende Procedur muss aber im Adressraum des
; DESC: eigenen Programms liegen
;
; PARA(ptrParLst): Pointer der Parameterliste
; PARA(nPar)     : Anzahl der zu übergebenden Parameter
; PARA(ptrProc)  : Proceduradresse
; RET            : Rückgabewert der aufgerufenen Procedur
; ===========================================================================

CallProc proc   ptrParLst   : DWORD,        ; Adresse Paramterliste
                nPar        : DWORD,        ; Anzahl der Parameter
                ptrProc     : DWORD,        ; Proceduradresse
                ptrRet      : DWORD,        ; Pointer des Rückgabewertes
                 
        mov edx, ptrParLst      ; ptr Parameterliste
        mov ecx, nPar           ; Anzahl der Parameter
        test ecx, ecx           ; Wenn Anzahl der Parameter = 0, 
        jz @chkRET               ; dann gleich Funktion aufrufen

        ; Adresse des letzten Parameters berechnen, der muss als 1. auf den Stack
        ; Adresse (EDX) = ptrParLst + (nPar-1) * 4
        dec ecx                 ; nPar-1
        add edx, ecx            ; 4x (nPar-1) addieren
        add edx, ecx            ; das ist einfacher als mit
        add edx, ecx            ; 4 zu multiplizieren
        add edx, ecx   
        inc ecx                 ; Anzahl Parameter wieder herstellen

    @loop:                      ; n Paramter in Schleife auf Stack schieben
        mov eax, [edx]          ; Parameter in EAX laden
        push eax                ; auf Stack
        sub edx, 4              ; Adresse nächster Paramter              
        loop @loop              ; Anzahl noch zu bearbeitender Parameter
    @chkRET:
        mov eax, ptrRet         ; Option: Rückgabewert über Pointer handhaben
        test eax, eax           ; eax=0
        jz @call                ; wenn ja : Funktion aufrufen und Rückgabe über EAX
        push eax                ; RückgabePointer auf Stack

    @call:
        mov eax, ptrProc        ; Adresse Zielprocedur laden
        call eax                ; Indirekter Aufruf Zielprocedur
        ret
CallProc endp

; ===========================================================================
; NAME: CallObjProc
; DESC: aufrufen beliebiger Proceduren in Objecten über deren Adresse
; DESC: die aufzurufende Procedur muss aber im Adressraum des
; DESC: eigenen Programms liegen. Die Procedureadresse im Object ist
; DESC: vorher mit GetObjProcPtr anhand der Procedur-Nr. zu ermitteln
;
; PARA(ptrParLst): Pointer der Parameterliste
; PARA(nPar)     : Anzahl der zu übergebenden Parameter
; PARA(ptrProc)  : Proceduradresse
; RET            : Rückgabewert der aufgerufenen Procedur
; ===========================================================================

CallObjProc proc    ptrParLst   : DWORD,        ; Adresse Paramterliste
                    nPar        : DWORD,        ; Anzahl der Parameter
                    ptrObj      : DWORD,        ; ObjectAdresse
                    ProcNr      : DWORD,        ; Procedurnummer
                    ptrRet      : DWORD,        ; Pointer des Rückgabewertes

        local lRet : DWORD      ; Rückgabewert für Object-Function (benötigt unbedingt Pointer)
          
        mov edx, ptrRet         ; Pointer Rückgabewert
        test edx, edx           ; ptrRet = 0?
        jnz @1                  ; wenn nicht 0, dann Pointer als Rückgabewert verwenden
        lea edx, lRet           ; sonst Rückgabewert über Lokalvariable
      @1:
        push edx                ; Pointer Rückgabewert ist 1. Parameter auf Stack
        mov edx, ptrParLst      ; Pointer Parameterliste
        mov ecx, nPar           ; Anzahl der Parameter
        test ecx, ecx           ; Wenn Anzahl der Parameter = 0, 
        jz @call                ; dann gleich Funktion aufrufen

        ; Adresse des letzten Parameters berechnen, der muss bei stdCall
        ; als 1. auf den Stack
        ; Adresse (EDX) = ptrParLst + (nPar-1) * 4
        dec ecx                 ; nPar-1
        add edx, ecx            ; 4x (nPar-1) addieren
        add edx, ecx            ; das ist einfacher als mit
        add edx, ecx            ; 4 zu multiplizieren
        add edx, ecx   
        inc ecx                 ; Anzahl Parameter wieder herstellen

    @loop:                      ; n Paramter in Schleife auf Stack schieben
        mov eax, [edx]          ; Parameter in EAX laden
        push eax                ; auf Stack
        sub edx, 4              ; Adresse nächster Paramter              
        loop @loop              ; Anzahl noch zu bearbeitender Parameter in ECX-Register

    @call:                      ; ptr auf Proceduradresse = [ObjPtr] + &H1C + (4 * ProcNumber)
        mov edx, ptrObj         ; ptrObject
        push edx
        mov edx, [edx]          ; StartAdresse der Objektdaten
        add edx, 01Ch           ; Startoffset der Procedurliste addieren
        mov eax, ProcNr         ; Procedur-Nummer, ProcNr
        shl eax, 2              ; ProcNr*4 (=Anzahl Bytes)
        add edx, eax            ; = ptr auf Proceduradresse
        mov eax, [edx]          ; ProcedurAdresse aus Procedurliste laden
        call eax                ; Indirekter Aufruf Zielprocedur

        mov edx, ptrRet         ; Pointer Rückgabewert
        test edx, edx           ; ptrRet = 0?
        jnz @2                  ; wenn nicht 0, dann Pointer dann Ende
        mov eax, lRet           ; sonst lRet zurückgeben
    @2: ret
CallObjProc endp


GetObjProcPtr proc ptrObj : DWORD,
                   ProcNr : DWORD,
                   
    mov edx, ptrObj            ; ptrObject
    mov edx, [edx]             ; StartAdresse der Objektdaten
    add edx, 01Ch              ; Startoffset der Procedurliste addieren
    mov eax, ProcNr            ; Procedur-Nummer, ProcNr
    shl eax, 2                 ; ProcNr*4 (=Anzahl Bytes)
    add edx, eax               ; = ptr auf Proceduradresse
    mov eax, [edx]             ; ProcedurAdresse aus Procedurliste laden
    ret 
GetObjProcPtr endp


; ===========================================================================
;  NAME: IsArrayDim
;  DESC: Ermittelt ob ein Array (SafeArray) dimensioniert ist 
;
;  PARA(VbArray): Array-Pointer
;  RET(EAX):      vbFalse = undimensioniert; vbTrue=dimensioniert
; ===========================================================================

IsArrayDim proc
    mov edx, [esp+4]            ; von VB übergebene Adresse des Arrays in EDX
    mov eax, [edx]              ; an dieser Adresse steht der SafeArray-Pointer
    test eax, eax               ; SafeArray-Pointer <> 0 prüfen
    JNZ @dim                    ; ist <> 0 => Array ist dimensioniert
    xor eax, eax                ; undemiensioniert; 0 (vbFalse) zurückgeben
    jmp @ret
@dim:                           ; dimensioniert
    xor eax, eax                ; eax :=0
    not eax                     ; eax :=-1 (vbTrue)
@ret:    
    ret 4   
IsArrayDim endp

; ===========================================================================
;  NAME: ptrArrayStrukt
;  DESC: Ermittelt den Pointer der zugehörigen SafeArrayStruktur des Arrays
;  DESC: und gibt diesen zurück. Um den Pointer der SafeArrayStruktur zu setzen, 
;  DESC: die Funkion HookArray (hängt SafeArrayStruktur ein), verwenden!0
;
;  PARA(Array):
;  RET:  Pointer auf SafeArrayStruktur
; ===========================================================================

ptrArrayStrukt proc
    mov edx, [esp+4]            ; Adresse Array
    mov eax, [edx]              ; SafeArray-Pointer Array 
    ret 4
ptrArrayStrukt endp

; ===========================================================================
;  NAME: XchgArray
;  DESC: Tauscht 2 Arrays gegeneinander aus.
;  DESC: es werden nur die Poitner der SafearrayStrukturen getauscht
;
;  PARA(Array1): Array1
;  PARA(Array2): Array2
; ===========================================================================

XchgArray proc
    mov edx, [esp+4]            ; Adresse Array 1
    mov eax, [edx]              ; SafeArray-Pointer Array 1
    mov edx, [esp+8]            ; Adresse Array 2
    mov ecx, [edx]              ; SafeArray-Pointer Array 2
    mov [edx], eax              ; SafeArray-Pointer 2 = SafeArray-Pointer 1              
    mov edx, [esp+4]            ; Adresse Array 1
    mov [edx], ecx              ; SafeArray-Pointer 1 = SafeArray-Pointer 2
    ret 8
XchgArray endp

; ===========================================================================
;  NAME: UnHookArray
;  DESC: Hängt die SafeArray-Struktur des Arrays aus, d.h. der Pointer
;  DESC: auf die zugehörige SafeArray-Struktur wird gelöscht (:=0)
;  DESC: Der Array ist somit undimensionert.
;  DESC: Der ursprüngliche SafeArrayPointer wird als Funktionrückgabewert
;  DESC: zurückgegeben.
;  DESC: ACHTUNG: Wenn der Array ausgehängt wird, hängt der reservierte
;  DESC: Speicherbereich in der Luft - VB weis davon im Prinzip nichts.
;  DESC: Vor dem beenden des Programms bzw. vor auflösen des Arrays 
;  DESC: sollte der ursprüngliche Zustand wieder hergestellt werden.
;
;  PARA(Array): Array
;  RET:  Pointer der urspünglichen SafeArray-Struktur
; ===========================================================================

UnHookArray proc
    mov edx, [esp+4]            ; Adresse Array
    mov eax, [edx]              ; SafeArray-Pointer für Funktionsrückgabe nach EAX
    mov [edx], dword ptr 0      ; SafeArray-Pointer löschen (Array ist nun undimensioniert)
    ret 4
UnHookArray endp

; ===========================================================================
;  NAME: HookArray
;  DESC: Hängt eine SafeArray-Struktur in den Array. Die einzuhängende 
;  DESC: SafeArrayStruktur wird mit ihrem Pointer angegeben.
;  DESC: Der Pointer der ursprünglichen SafeArrray-Struktur wird als
;  DESC: Fuktionswert zurückgegeben
;  DESC: ACHTUNG: VB bekommt davon nichts mit. Vor beenden des Programms
;  DESC: bzw. vor dem Auflösen der Arrays sollte der ursprüngliche Zustand
;  DESC: wieder hergestellt werden.
;
;  PARA(Array): Array
;  PARA(pSFA):  Pointer der neuen SafeArray-Struktur
;  RET:  Pointer der urspünglichen SafeArray-Struktur
; ===========================================================================

HookArray proc
    mov edx, [esp+4]            ; Adresse Array
    mov eax, [edx]              ; SafeArray-Pointer für Funktionsrückgabe nach EAX
    mov ecx, [ESP+8]
    mov [edx], ecx              ; Pointer auf einzuhängende SafeArray-Struktur
    ret 8
HookArray endp

; ===========================================================================
;  NAME: GetBit 
;  DESC: Ermittelt den Wert eines Bits in einem Wert
;  DESC: Die Funktion exisitiert als 16/32 Bit Version	
;
;  PARA(WERT):   Quellwert
;  PARA(Pos):    Position des zu prüfenden Bits
;  RET(EAX):     Rückgabewert:  vbFalse(0), wenn Bit = 0 
;                		        vbTrue(-1), wenn Bit = 1
; ===========================================================================

GetBit16 proc 
	movzx ecx, word ptr [esp+4] ;     WERT
	mov edx, [esp+8] ;     Pos.
	xor eax, eax     ;     EAX := 0

	bt ecx, edx       ;    BitTest Wert, Bitposition : Ergebnis im CarryFlag
	jnc @16END        ;    Wenn Carry/Bit = 0, dann ENDE, EAX=0 zurück           
	not eax           ;    Wenn Carry/Bit = 1, dann EAX=-1 zurück (VB-True)
@16END: ret 8
GetBit16 endp

GetBit32 proc 
	mov ecx, [esp+4] 	; WERT
	mov edx, [esp+8] 	; Pos
	xor eax, eax     	; EAX := 0

	bt ecx, edx        	; BitTest Wert, Bitposition : Ergebnis im CarryFlag
	jnc @32END        	; Wenn Carry/Bit = 0, dann ENDE, EAX=0 zurück           
	not eax           	; Wenn Carry/Bit = 1, dann EAX=-1 zurück (VB-True)
  @32END:
      ret 8
GetBit32 endp

; ===========================================================================
;  NAME: SetBit 
;  DESC: Setzt den Wert eines Bits nach Vorgabe auf True oder False
;  DESC: Die Funktion exisitiert als 16/32 Bit Version	
;
;  PARA(WERT):   Quellwert
;  PARA(Pos):    Position des zu prüfenden Bits
;  PARA(NewVal): Neuer Wert des zu bearbeitenden Bits (VbTrue/VbFalse)
;  RET(EAX): 	 Rückgabewert:  WERT mit entsprechend bearbeitetem Bit
; ===========================================================================

SetBit16 proc
	movzx eax, word ptr [esp+4] 	      ; WERT
	mov edx, [esp+8] 		; Pos.
	movsx ecx, word ptr [esp+12] 		; Wert auf den das Bit gesetzt werden soll VbTrue/VbFalse
	test ecx, -1    			      ; auf <> Null prüfen
	jz @res16			            ; =0, dann Bit rücksetzen
	bts eax, edx                        ; sonst Bit setzen
	ret
   @res16:
      btr eax, edx	                  ; Bit reset
	ret 12	
SetBit16 endp

SetBit32 proc
	mov eax, [esp+4] 		; WERT
	mov edx, [esp+8] 		; Pos.
	movsx ecx, word ptr [esp+12] 		; Wert auf den das Bit gesetzt werden soll VbTrue/VbFalse
	test ecx, -1			      ; auf <> Null prüfen
	jz @Res32			            ; =0, dann Bit rücksetzen
	bts eax, edx		            ; sonst Bit setzen
	ret
   @Res32:
      btr eax, edx                      ; Bit reset
	ret 12	 
SetBit32 endp

; ===========================================================================
;  NAME: InvBit 
;  DESC: Invertiert den Wert eines angegeben Bits 
;  DESC: Die Funktion exisitiert als 16/32 Bit Version	
;
;  PARA(WERT):   Quellwert
;  PARA(Pos):    Position des zu invertierenden Bits (0..x)
;  RET(EAX):     Rückgabewert:  WERT mit entsprechend bearbeitetem Bit
; ===========================================================================

InvBit16 proc
    movzx eax, word ptr [esp+4]     ; zu bearbeitenden Wert in EAX laden
    mov ecx, [esp+8]                ; Bitposition
    btc eax, ecx                    ; Bit invertieren
    ret 8
InvBit16 endp

InvBit32 proc
    mov eax, [esp+4]                ; zu bearbeitenden Wert in EAX laden
    mov ecx, [esp+8]                ; Bitposition
    btc eax, ecx                    ; Bit invertieren
    ret 8
InvBit32 endp

; ===========================================================================
;  NAME: SHL 
;  DESC: Schieben links 
;  DESC: Die Funktion exisitiert als 16/32 Bit Version	
;
;  PARA(WERT):   Quellwert
;  PARA(Shift):  Anzahl zu schiebender Bits
;  RET(EAX):     Rückgabewert:  WERT entprechende Bits verschoben
; ===========================================================================

SHL16 proc
    movzx eax, word ptr [esp+4]     ; 16 Bit Wert in Akku laden
    mov ecx, [esp+8]                ; Anzahl zu schiebender Bits
    shl ax, cl                      ; Shift left
    ret 8
SHL16 endp

SHL32 proc
    mov eax, [esp+4]                ; 16 Bit Wert in Akku laden
    mov ecx, [esp+8]                ; Anzahl zu schiebender Bits
    shl eax, cl                     ; Shift left
    ret 8
SHL32 endp

; ===========================================================================
;  NAME: SHR 
;  DESC: Schieben rechts 
;  DESC: Die Funktion exisitiert als 16/32 Bit Version	
;
;  PARA(WERT):   Quellwert
;  PARA(Shift):  Anzahl zu schiebender Bits
;  RET(EAX):     Rückgabewert:  WERT entprechende Bits verschoben
; ===========================================================================

SHR16 proc
    movzx eax, word ptr [esp+4]     ; 32 Bit Wert in Akku laden
    mov ecx, [esp+8]                ; Anzahl zu schiebender Bits
    shr ax, cl                      ; Shift left
    ret 8
SHR16 endp

SHR32 proc
    mov eax, [esp+4]                ; 32 Bit Wert in Akku laden
    mov ecx, [esp+8]                ; Anzahl zu schiebender Bits
    shr eax, cl                     ; Shift left
    ret 8
SHR32 endp

; ===========================================================================
;  NAME: ROL 
;  DESC: Rotieren links 
;  DESC: Die Funktion exisitiert als 16/32 Bit Version	
;
;  PARA(WERT):   Quellwert
;  PARA(rotate): Anzahl zu rotierender Bits
;  RET(EAX):     Rückgabewert:  WERT entprechende Bits rotiert
; ===========================================================================

ROL16 proc
    movzx eax, word ptr [esp+4]     ; 16 Bit Wert in Akku laden
    mov ecx, [esp+8]                ; Anzahl zu rotierender Bits
    rol ax, cl                      ; rotate left
    ret 8
ROL16 endp

ROL32 proc
    mov eax, [esp+4]                ; 16 Bit Wert in Akku laden
    mov ecx, [esp+8]                ; Anzahl zu rotierender Bits
    rol eax, cl                     ; rotate right
    ret 8
ROL32 endp

; ===========================================================================
;  NAME: ROR 
;  DESC: Rotieren rechts 
;  DESC: Die Funktion exisitiert als 16/32 Bit Version	
;
;  PARA(WERT):   Quellwert
;  PARA(rotate): Anzahl zu rotierender Bits
;  RET(EAX):     Rückgabewert:  WERT entprechende Bits rotiert
; ===========================================================================

ROR16 proc
    movzx eax, word ptr [esp+4]     ; 16 Bit Wert in Akku laden
    mov ecx, [esp+8]                ; Anzahl zu rotierender Bits
    ror ax, cl                      ; rotate right
    ret 8
ROR16 endp

ROR32 proc
    mov eax, [esp+4]                ; 16 Bit Wert in Akku laden
    mov ecx, [esp+8]                ; Anzahl zu rotierender Bits
    ror eax, cl                     ; rotate right
    ret 8
ROR32 endp

; ===========================================================================
;  NAME: LSB, least significant Bit 
;  DESC: sucht das niederwertigste Bit und gibt dessen Position zurück
;  DESC: Die Funktion exisitiert als 16/32 Bit Version	
;
;  PARA(WERT):   Quellwert
;  RET():        Rückgabewert:  Bit-Index
; ===========================================================================

LSB16 proc
    movzx edx, word ptr [esp+4]     ; Wert in EDX laden
    xor eax, eax                    ; eax löschen
    bsf eax, edx                    ; erstes Bit in EDX suchen, Ergenis nach EAX
    ret 4
LSB16 endp

LSB32 proc
    mov edx, [esp+4]                ; Wert in EDX laden
    xor eax, eax                    ; eax löschen
    bsf eax, edx                    ; erstes Bit in EDX suchen, Ergenis nach EAX
    ret 4
LSB32 endp

; ===========================================================================
;  NAME: MSB, most significant Bit 
;  DESC: sucht das hoechstwertigste Bit und gibt dessen Position zurück
;  DESC: Die Funktion exisitiert als 16/32 Bit Version	
;
;  PARA(WERT):   Quellwert
;  RET():        Rückgabewert:  Bit-Index
; ===========================================================================

MSB16 proc
    movzx edx, word ptr [esp+4]              ; Wert in EDX laden
    xor eax, eax                    ; eax löschen
    bsr eax, edx                    ; letztes Bit in EDX suchen, Ergenis nach EAX
    ret 4
MSB16 endp

MSB32 proc
    mov edx, [esp+4]                ; Wert in EDX laden
    xor eax, eax                    ; eax löschen
    bsr eax, edx                    ; letztes Bit in EDX suchen, Ergenis nach EAX
    ret 4
MSB32 endp

; ===========================================================================
;  NAME: BitCount, Bits zählen 
;  DESC: ermittelt die Anzahl der gesetzten Bits (=1) in einem Wert
;  DESC: Die Funktion exisitiert als 16/32 Bit Version	
;
;  PARA(WERT):   Quellwert
;  RET():        Rückgabewert:  Anzahl der gesetzten Bits
; ===========================================================================

BitCount16 proc
        movzx edx, word ptr [esp+4]     ; Wert in EDX Register laden
        xor eax, eax                    ; Anzahl Bits loeschen
        xor ecx, ecx                    ; ECX löschen
        not cx                          ; Bitmaske Schleifenzähler 16-Bits auf 1 setzen
    @loop:
        shr ecx, 1                      ; Bitmaske der zu zählenden Bits schieben
        jnc @end                        ; Wenn keine Bits mehr in der Zählmaske, dann Ende

        shr edx, 1                      ; zu zählendes Bit in Carry-Flag schieben
        jnc @loop                       ; wenn Bit = 0, dann weiter, sonst
        inc eax                         ; Bit zählen
        jmp @loop                       ; weiter, Schleife
    @end:
        ret 4
BitCount16 endp

BitCount32 proc
        mov edx,  [esp+4]               ; Wert in EDX Register laden
        xor eax, eax                    ; Anzahl Bits loeschen
        xor ecx, ecx                    ; ECX löschen
        not ecx                         ; Bitmaske Schleifenzähler 32Bit auf 1 setzen
    @loop:
        shr ecx, 1                      ; Bitmaske der zu zählenden Bits schieben
        jnc @end                        ; Wenn keine Bits mehr in der Zählmaske, dann Ende

        shr edx, 1                      ; zu zählendes Bit in Carry-Flag schieben
        jnc @loop                       ; wenn Bit = 0, dann weiter, sonst
        inc eax                         ; Bit zählen
        jmp @loop                       ; weiter, Schleife
    @end:
        ret 4
BitCount32 endp

; ===========================================================================
;  NAME: GetBitBlock 
;  DESC: gibt den Wert eines Bitbereichs (Block) in eiem 32-Bit-Wert zurueck
;  DESC: Value = 0000 1001 1100 0110
;  DESC; der Wert des BitBereichs 0 bis 3 ist '0110' also 6
;
;  PARA(WERT):     Quellwert 32 Bit
;  PARA(FirstBit): Nr. des ersten Bit (0-basiert)
;  PARA(LastBit):  Nr. des letzten Bits (0-basiert)
;  RET:          Rückgabewert: gewünschter BitBereich
; ===========================================================================

;	Maskierungswert für &-Verknüpfung:
;	0xFFFFFFFF enthält nur 1en, diese werden um die benötigen Bits links verschoben,
;	nach invertieren dieses Wertes ergeben sich rechts die benötigten 1en
;	Value muss nun noch um StartBit nach rechts verschoben werden und mit dem
;	Maskierungswert &-verknüpft, das ist dann der gesuchte Wert der
;	BitArea(StartBit bis EndBit)

GetBitBlock proc
	mov dl, byte ptr [esp+8]          ; FirstBit (0-basiert)
	mov cl, byte ptr [esp+10]         ; EndBit
	sub cl, dl		                ; = EndBit-StartBit
	inc cl                            ; = EndBit-StartBit +1
	xor eax, eax		          ; EAX:=0
	not eax			          ; EAX:=OxFFFFFFFF
	shl eax, cl		                ; 1er Maske um (EndBit-StartBit +1) nach links
	mov ecx, edx                      ; FirstBit in C-Register, wegen Schiebeoperation                    
 	mov edx, [esp+4]	                ; Wert in EDX laden
      shr edx, cl                       ; Wert um StartBit nach rechts
	and eax, edx                      ; nun noch geschobenen Wert mit berechneter Maske &-verknuepfen 
	ret 8                             ; Thats it!
GetBitBlock endp

; ===========================================================================
;  NAME: GetLoByte 
;  DESC: gibt das niederwertige Byte eines 16 Bit Wertes zurück
;
;  PARA(WERT):   Quellwert 16 Bit
;  RET:          Rückgabewert: niederwertiges Byte  
; ===========================================================================

GetLoByte proc
    movzx eax, byte ptr [esp+4]
    ret 4
GetLoByte endp

; ===========================================================================
;  NAME: GetHiByte 
;  DESC: gibt das hoeherwertige Byte eines 16 Bit Wertes zurück
;
;  PARA(WERT):   Quellwert 16 Bit
;  RET:          Rückgabewert: hoeherwertiges Byte  
; ===========================================================================

GetHiByte proc
    movzx eax, byte ptr [esp+5]
    ret 4
GetHiByte endp

; ===========================================================================
;  NAME: SetLoByte 
;  DESC: setzt das niederwertige Byte eines 16Bit Wertes auf den
;  DESC: übergeben Wert
;
;  PARA(WERT):   Quellwert 16 Bit
;  PARA(NewLo):  Neuer Wert des LoBytes, 8Bit
;  RET:          neuer 16Bit-Wert
; ===========================================================================

SetLoByte proc
    movzx eax, word ptr [esp+4] ; Ausgangswert laden
    mov al, byte ptr [esp+6]    ; neues LoByte in LoByte des Akku laden
    ret 8                       ; neuen Wert in EAX zurückgeben
SetLoByte endp

; ===========================================================================
;  NAME: SetHiByte 
;  DESC: setzt das hoeherwertige Byte eines 16Bit Wertes auf den
;  DESC: übergeben Wert
;
;  PARA(WERT):   Quellwert 16 Bit
;  PARA(NewHi):  Neuer Wert des HiBytes, 8Bit
;  RET:          neuer 16Bit-Wert
; ===========================================================================

SetHiByte proc
    movzx eax, word ptr [esp+4]     ; Ausgangswert laden
    mov ah, byte ptr [esp+6]        ; neues HiByte in HiByte des Akku laden
    ret 8                           ; neuen Wert in EAX zurückgeben
SetHiByte endp

; ===========================================================================
;  NAME: GetLoWord 
;  DESC: gibt das niederwertige Word eines 32 Bit Wertes zurück
;
;  PARA(WERT):   Quellwert 32 Bit
;  RET:          Rückgabewert: niederwertiges Word  
; ===========================================================================

GetLoWord proc
    movzx eax, word ptr [esp+4]    ; nur das LoWord des Ausgangswertes laden
    ret 4
GetLoWord endp

; ===========================================================================
;  NAME: GetHiWord 
;  DESC: gibt das hoeherwertige Word eines 32 Bit Wertes zurück
;
;  PARA(WERT):   Quellwert 32 Bit
;  RET:          Rückgabewert: hoeherwertiges Word  
; ===========================================================================

GetHiWord proc
    movzx eax, word ptr [esp+6]    ; nur das HiWord des Ausgangswertes laden 
    ret 4
GetHiWord endp

; ===========================================================================
;  NAME: SetLoWord 
;  DESC: setzt das niederwertige Word eines 32Bit Wertes auf den
;  DESC: übergeben Wert
;
;  PARA(WERT):   Quellwert 32 Bit
;  PARA(NewLo):  Neuer Wert des LoWords, 16Bit
;  RET:          neuer 32 Bit-Wert
; ===========================================================================

SetLoWord proc
    mov eax, [esp+4]            ; ursprünglichen Wert in EAX laden
    mov ax, word ptr [esp+8]    ; neues LoWord dazuladen
    ret 8
SetLoWord endp

; ===========================================================================
;  NAME: SetHiWord 
;  DESC: setzt das hoeherwertige Word eines 32Bit Wertes auf den
;  DESC: übergeben Wert
;
;  PARA(WERT):   Quellwert 32 Bit
;  PARA(NewLo):  Neuer Wert des HiWords, 16Bit
;  RET:          neuer 32 Bit-Wert
; ===========================================================================

SetHiWord proc
    movzx eax, word ptr [esp+8]     ; neues HiWord laden
    shl eax, 16                     ; neues HiWord an Hi-Position verschieben 
    mov ax, word ptr [esp+4]        ; ursprüngliches LoWord dazuladen
    ret 8
SetHiWord endp

; ===========================================================================
;  NAME: ByteSwap 
;  DESC: vertauscht die Bytereihenfolgen eins 32 Bit Wertes
;  DESC: Konvertierung zwischen den Formaten Little- und Big-Endian
;
;  PARA(WERT):   Quellwert 32 Bit
;  RET:          neuer 32 Bit-Wert mit vertauschter Bytereihenfolge
; ===========================================================================

ByteSwap proc
    mov eax, [esp+4]    ; 32 Bit-Wert laden
    bswap eax           ; Bytereihenfolge vertauschen
    ret 4
ByteSwap endp

; ===========================================================================
;  NAME: WordSwap 
;  DESC: vertauscht Hi- und Lo-Word eins 32 Bit Wertes
;
;  PARA(WERT):   Quellwert 32 Bit
;  RET:          neuer 32 Bit-Wert mit vertauschter Wordreihenfolge
; ===========================================================================

WordSwap proc
    mov eax, [esp+4]    ; 32 Bit-Wert laden
    ror eax, 16         ; durch Rotation von 16 wird Hi- und Lo vertauscht
    ret 4
WordSwap endp    

; ===========================================================================
;  NAME: Add64 : Add Large
;  DESC: Addiert zwei 64 Bit Large Integer (signed oder unsigned)	
;
;  PARA(pOP1):   Pointer Operand 1 64Bit Integer (Speicherformat Lo, Hi)
;  PARA(pOP2):   Pointer Operand 2 64Bit Integer (Speicherformat Lo, Hi)
;  RET(EDX:EAX): Rückgabewert, 64Bit Integer 
;                EDX:EAX = Operand1 + Operand2
; ===========================================================================

Add64 proc  pOP1 : dword,    ; Pointer auf Operand 1
            pOP2 : dword,    ; Pointer auf Operand 2  

    push ebx                ;save EBX as per calling convention
       
    ; Operanden in Register laden
    ; Operand1 Hi:Lo, EDX:EAX
    ; Operand2 Hi:Lo, ECX:EBX
    mov ebx, pOP1           ;Adresse Operand1
    mov eax, [ebx]          ;Operand1 Lo
    mov edx, [ebx+4]        ;Operand1 Hi
    mov ebx, pOP2           ;Adresse Operand2
    mov ecx, [ebx+4]        ;Operand2 Hi
    mov ebx, [ebx]          ;Operand2 Lo
    
    ; addiert Operand in ECX:EBX zu Operand EDX:EAX
    ; Ergebnis EDX:EAX, Hi:Lo
    add eax, ebx            ; EAX= Lo1 + Lo2
    adc edx, ecx            ; EDX= Hi1 + Hi2 + Carry
    pop ebx
    ret
Add64 endp


; ===========================================================================
;  NAME: Sub64 
;  DESC: Subtrahiert zwei 64 Bit Large Integer (signed oder unsigned)	
;
;  PARA(pOP1):   Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
;  PARA(pOP2):   Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
;  RET(EDX:EAX): Rückgabewert, 64Bit Integer 
;                EDX:EAX = Operand1 - Operand2
; ===========================================================================

Sub64 proc pOP1 : dword,    ; Pointer auf Operand 1
           pOP2 : dword,    ; Pointer auf Operand 2  

    push ebx                ;save EBX as per calling convention
    
    ; Operanden in Register laden
    mov ebx, pOP1           ;Adresse Operand1
    mov eax, [ebx]          ;Operand1 Lo
    mov edx, [ebx+4]        ;Operand1 Hi
    mov ebx, pOP2           ;Adresse Operand2
    mov ecx, [ebx+4]        ;Operand2 Hi
    mov ebx, [ebx]          ;Operand2 Lo
    
    ;subtract operand in ECX:EBX from operand EDX:EAX, result in
    ; EDX:EAX
    sub eax, ebx
    sbb edx, ecx
    
    pop ebx
    ret
Sub64 endp


; ===========================================================================
;  NAME: Mul64 
;  DESC: Multipliziert zwei 64 Bit Large Integer (signed oder unsigned)	
;  PARA(pOP1):   Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
;  PARA(pOP2):   Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
;  RET(EDX:EAX): Rückgabewert, 64Bit Integer 
;                EDX:EAX = Operand1 * Operand2

;  Code from AMD's Code Optimization Guide (Efficient 64-Bit Integer Arithmetic)
; ===========================================================================

Mul64 proc pOP1 : dword,    ; Pointer auf Operand 1, Multiplicand
           pOP2 : dword,    ; Pointer auf Operand 2, Multiplier

    ; OUTPUT: EDX:EAX (multiplicand * multiplier) % 2^64

    push ebx                       ;save EBX as per calling convention
    
    mov ebx, pOP1                  ; Adresse Operand 1, Multiplicand
    mov edx, [ebx+4]               ; Multiplicand_hi
    mov ebx, pOP2                  ; Adresse Operand 2, Multiplier
    mov ecx, [ebx+4]               ; Multipliplier_hi
    or  edx, ecx                   ; Ein Operand >= 2^32?, dann 2 Multiplikationen
    mov edx, [ebx]                 ; multiplier_lo
    mov ebx, pOP1                  ; Adresse Operand 1, Multiplicand
    mov eax, [ebx]                 ; multiplicand_lo
    jnz $TwoMul                    ; ja, zwei Multiplikationen
    
    mul edx                        ; multiplicand_lo * multiplier_lo
    jmp $RET_MulLarge              ; done, return to caller
    
  $TwoMul:
    imul edx, [ebx+4]              ; p3_lo = multiplicand_hi*multiplier_lo
    imul ecx, eax                  ; p2_lo = multiplier_hi*multiplicand_lo
    add ecx, edx                   ; p2_lo + p3_lo
    mov ebx, pOP2                  ; Adresse Operand 1, Multiplicand
    mul dword ptr [ebx]            ; p1=multiplicand_lo*multiplier_lo
    add edx, ecx                   ; p1+p2lo+p3_lo = result in EDX:EAX
    
    ; Wert zurückschreiben
  $RET_MulLarge:
    pop ebx
    ret                           ; done, return to caller
Mul64 endp


; ===========================================================================
;  NAME: Neg64 (NegateLarge)
;  DESC: Negiert einen 64 Bit Large Integer; Vorzeichenwechsel; Zweierkomplement
;  PARA(pOP1):   Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
;  RET(EDX:EAX): Rückgabewert, 64Bit Integer 
; ===========================================================================

Neg64 proc pOP1 : dword,    ; Pointer auf Operand
        
    ; Operanden in Register laden
    mov ecx, pOP1           ;Adresse Operand1
    mov eax, [ecx]          ;Operand Lo
    mov edx, [ecx+4]        ;Operand Hi
    
    ;negate Operand in EDX:EAX
    not edx
    neg eax
    sbb edx, -1             ;fixup: increment hi-word if low-word was 0
    
    ret
Neg64 endp

; ===========================================================================
;  NAME: Not64 
;  DESC: Negiert einen 64 Bit Large Integer; Einerkomplement
;  PARA(pOP1):   Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
;  RETURN:       EAX/EDX 
; ===========================================================================

Not64 proc pOP1 : dword,    ; Pointer auf Operand
        
    ; Operanden in Register laden
    mov ecx, pOP1           ;Adresse Operand1
    mov eax, [ecx]          ;Operand Lo
    mov edx, [ecx+4]        ;Operand Hi
    
    ;invert Operand in EDX:EAX
    not edx
    not eax
    
    ret
Not64 endp

; ===========================================================================
;  NAME: Shr64 (ShiftRightLarge)
;  DESC: Schiebt einen 64 Bit Large Integer um Bits nach rechts 
;  PARA(pOP1):    Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
;  RET(EDX:EAX): Ergebnis, EAX = Lo (32Bit), EDX = Hi (32Bit)

;  Code by Stefan Maag
; ===========================================================================

Shr64 proc pOP1:dword,            ; Pointer auf Operand
bits: dword,                  	; Anzahl der zu schiebenden Bits

             
    ; Operanden in Register laden
    mov ecx, pOP1           ;Adresse Operand1
    mov eax, [ecx]          ;Operand Lo
    mov edx, [ecx+4]        ;Operand Hi
    mov ecx, bits
    
    ; SHRD: ab.386 : fasst zwei Register bei einer Schiebeopration zusammmen
    ;             SHRD Lo, Hi, nBits
    ;       ABER: SHLD Hi, Lo, nBits
    ;       Hi wird bei der Operation nicht gespeichert und muss 
    ;       Extra mit SHR Hi, nBits bearbeitet werden
    ;       Die Schiebeweite wird automatisch mit Modulo32 verrechnet und
    ;       ist somit auf max 31 begrenzt. Schiebeweiten ab 32 werden
    ;       durch zusätulichen Tausch der Register erreicht.

   
    cmp ecx, 32          ; Wenn Bits < 32
    jb $2_shift          ; dann 2 Schiebeoperationen
    cmp ecx, 64          ; Wenn Bits < 64 dann 
    jb $1_shift          ; 1 Schiebeoperation

    ; Ab 64 Bit ist das Ergebnis immer 0, dies muss
    ; extra bearbeitet werden, sonst entstehen durch die
    ; integrierte Modulo32 Verarbeitung des Schiebebefehls Fehler,
    ; Es wird dann um Mod32(Bits) geschoben
    xor eax, eax        ; EAX löschen
    xor edx, edx        ; EDX löschen
    ret

  $1_shift:
    mov eax, edx         ; Hi 32 Bit durch Registertausch in Lo 32 Bit 
    xor edx, edx         ; Hi 32 Bit löschen
    shr eax, cl          ; Bits schieben (wird vom Prozessor mit Modulo 32 verarbeitet)
    ret                  ; deswegen entfaellt eine manuelle Verrechnung der Schiebweite
    
  $2_shift:
    shrd eax, edx, cl   ; 64-Bit Schiebebefehl (SHRD Lo, Hi, n)Bits
    shr edx, cl         ; Hi muss nochmals extra bearteiter werden, da nich mitgespeichert wird
   ret

Shr64 endp

; ===========================================================================
;  NAME: Shl64 (ShiftLeftLarge)
;  DESC: Schiebt einen 64 Bit Large Integer um Bits nach links 
;  PARA(pOP1):    Pointer auf 64Bit Large Integer (Speicherformat Lo, Hi)
;  RET(EDX:EAX): Ergebnis, EAX = Lo (32Bit), EDX = Hi (32Bit)

;  Code by Stefan Maag
; ===========================================================================


Shl64 proc pOP1:dword,               ; Pointer auf Operand
           bits: dword               ; Anzahl der zu schiebenden Bits

; Operanden in Register laden
    mov ecx, pOP1           ;Adresse Operand1
    mov eax, [ecx]          ;Operand Lo
    mov edx, [ecx+4]        ;Operand Hi
    mov ecx, bits
    
    ; SHRD: ab.386 : fasst zwei Register bei einer Schiebeopration zusammmen
    ;             SHLD Hi, Lo, nBits
    ;       ABER: SHRD Lo, Hi, nBits
    ;       Hi wird bei der Operation nicht gespeichert und muss 
    ;       Extra mit SHL Lo, nBits bearbeitet werden
    ;       Die Schiebeweite wird automatisch mit Modulo32 verrechnet und
    ;       ist somit auf max 31 begrenzt. Schiebeweiten ab 32 werden
    ;       durch zusätulichen Tausch der Register erreicht.

    cmp ecx, 32          ; Wenn Bits < 32
    jb $2_shift          ; dann 2 Schiebeoperationen
    cmp ecx, 64          ; Wenn Bits < 64 dann 
    jb $1_shift          ; 1 Schiebeoperation

    ; Ab 64 Bit ist das Ergebnis immer 0, dies muss
    ; extra bearbeitet werden, sonst entstehen durch die
    ; integrierte Modulo32 Verarbeitung des Schiebebefehls Fehler
    ; Es wird dann um Mod32(Bits) geschoben
    xor eax, eax        ; EAX löschen
    xor edx, edx        ; EDX löschen
    ret
    
  $1_shift:  
    mov edx, eax         ; Lo 32 Bit durch Registertausch in Hi 32 Bit 
    xor eax, eax         ; Hi 32 Bit löschen
    shl edx, cl          ; Bits schieben (wird vom Prozessor mit Modulo 32 verarbeitet)
    ret                  ; deswegen entfaellt eine manuelle Verrechnung der Schiebweite

  $2_shift:
    shld edx, eax, cl   ; 64-Bit Schiebebefehl (SHLD Hi, Lo, n)Bits
    shl eax, cl         ; Hi muss nochmals extra bearteiter werden, da nich mitgespeichert wird
   ret
Shl64          endp

; ===========================================================================
;  NAME: Div64 
;  DESC: dividiert zwei 64 Bit Integer
;  PARA(pOP1):   Pointer Divident 64Bit Integer (Speicherformat Lo, Hi)
;  PARA(pOP2):   Pointer Divisor  64Bit Integer (Speicherformat Lo, Hi)
;  RET(EDX:EAX): Rückgabewert, 64Bit Integer 
;                EDX:EAX = Operand1 / Operand2
; ===========================================================================

Div64    proc   pOP1 : dword,    ; Pointer auf Operand 1, Divident
                pOP2 : dword,    ; Pointer auf Operand 2, Divisor
                SIGNED: word      ; mit Vorzeichen; VB Boolean Parameter false, true

            ; Divident auf STACK legen
            mov ecx, pOP1           ;Adresse Divident
            mov eax, [ecx]          ;Divident Lo
            push eax                ;auf STACK
            mov eax, [ecx+4]        ;Divident Hi
            push eax                ;auf STACK
    
            ; Divisor auf STACK legen
            mov ecx, pOP2           ;Adresse Divisor
            mov eax, [ecx]          ;Divisor Lo
            push eax                ;auf STACK
            mov eax, [ecx+4]        ;Divisor Hi
            push eax                ;auf Stack

            movzx ecx, SIGNED       ; VB-Boolean, mit Vorzeichen
            test ecx, 1             ; Wert auf 0 prüfen
            jz $Div_uns             ; wenn=0, dann unsigned

            call sDiv               ;eigentliche Dividier-Routine signed
            jmp $Div_end
$Div_uns:   call uDiv               ;eigentliche Dividier-Routine unsigned
                ; DIVISOR_HI  : DWORD,    ; Divisor Hi    : ESP+4
                ; DIVISOR_LO  : DWORD,    ; Dviisor Lo    : ESP+8
                ; DIVIDENT_HI : DWORD,    ; Devident Hi   : ESP+0C
                ; DIVIDENT_LO : DWORD,    ; Divident Lo   : ESP+10

$Div_end:   ret
 
Div64 endp

; ===========================================================================
;  NAME: Mod64 
;  DESC: Modulo Division für 64 Bit Integer
;  PARA(pOP1):   Pointer Divident 64Bit Integer (Speicherformat Lo, Hi)
;  PARA(pOP2):   Pointer Divisor  64Bit Integer (Speicherformat Lo, Hi)
;  RET(EDX:EAX): Rückgabewert, 64Bit Integer 
;                EDX:EAX = Divisionsrest (Remainder)
; ===========================================================================

Mod64   proc    pOP1 : dword,    ; Pointer auf Operand 1, Divident
                pOP2 : dword,    ; Pointer auf Operand 2, Divisor
                SIGNED: word     ; mit Vorzeichen; VB Boolean Parameter false, true


            ; Divident auf STACK legen
            mov ecx, pOP1           ;Adresse Divident
            mov eax, [ecx]          ;Divident Lo
            push eax                ;auf STACK
            mov eax, [ecx+4]        ;Divident Hi
            push eax                ;auf STACK
    
            ; Divisor auf STACK legen
            mov ecx, pOP2           ;Adresse Divisor
            mov eax, [ecx]          ;Divisor Lo
            push eax                ;auf STACK
            mov eax, [ecx+4]        ;Divisor Hi
            push eax                ;auf Stack

            movzx ecx, SIGNED       ; VB-Boolean, mit Vorzeichen
            test ecx, 1             ; Wert auf 0 prüfen
            jz $Mod_uns             ; wenn=0, dann unsigned
            call sMOD               ; eigentliche Modulo Division signed
            jmp $Mod_end
$Mod_uns:   call uMOD               ;eigentliche Modulo Division unsigned
                ; DIVISOR_HI  : DWORD,    ; Divisor Hi    : ESP+4
                ; DIVISOR_LO  : DWORD,    ; Dviisor Lo    : ESP+8
                ; DIVIDENT_HI : DWORD,    ; Devident Hi   : ESP+0C
                ; DIVIDENT_LO : DWORD,    ; Divident Lo   : ESP+10

$Mod_end:   ret
   
Mod64 endp

; ===============================================================================
; Sqr64         create square root of LARGE
; -------------------------------------------------------------------------------
;               in      pOP1: ->LARGE            source value
;
;               ret     EDX:EAX                 sqare root (0 On Error)

; Code by Udo Schmidt
; ===============================================================================

Sqr64           proc
                pop     ecx                     ; Rücksprungadresse in ECX laden
                pop     eax                     ; pOP1 in EAX laden
                push    0                       ; temporäres EDX auf Stack erzeugen
                push    0                       ; temporäres EAX auf dem Stack erzeugen
                test    byte ptr [eax+7],80h    ; Vorzeichen OP1, wenn -, dann
                jnz     @1                      ; 0 ausgeben
                fild    qword ptr [eax]         ; Large in die FPU laden
                fsqrt
                fistp   qword ptr [esp]         ; Ergebnis in die temporären EAX:EDX auf Stack schreiben
@1:             pop     eax                     ; Ergebnis Low in EAX laden
                pop     edx                     ; Ergebnis High in EDX laden
                jmp     ecx                     ; Rücksprung ausführen
Sqr64           endp


; ===============================================================================
; StrToLarge:      convert STRING to Large
; -------------------------------------------------------------------------------
;               in      OP1: ->STRING           Value to convert
;                       FCT:   LONG             Factor to use (2 ... 36)
;               out     FLG: ->BOOLEAN          Overflow flag
;
;               ret     EDX:EAX                 LARGE value (0 on error)
;
;               - Accepts prefixes &B, &O, &D and &H

; Code by Udo Schmidt
; ===============================================================================

StrToLarge     proc
                push    esi
                push    edi
                push    ebx
                
                mov     esi,[esp+16]
                mov     edi,[esp+20]
                xor     eax,eax
                mov     ebx,eax
                mov     ecx,eax
                push    eax
                push    eax
                push    eax
                
                cmp     edi,2
                jc      VAL40_err
                sub     esi,1
                jc      VAL40_err

@1:             inc     esi
                movzx   edx,byte ptr [esi]
                cmp     dl,'a'
                jc      @2
                cmp     dl,'z'+1
                jc      @2
                and     dl,0DFh

@2:             or      bl,bl
                jns     @3
                or      dl,dl
                jz      VAL40_err
                mov     bl,2
                mov     edi,2
                cmp     dl,'B'
                jz      @1
                add     edi,6
                cmp     dl,'O'
                jz      @1
                add     edi,2
                cmp     dl,'D'
                jz      @1
                add     edi,6
                cmp     dl,'H'
                jz      @1
                jmp     VAL40_err

@3:             or      dl,dl
                jz      VAL40_end
                cmp     dl,'.'
                jz      VAL40_end

                cmp     dl,32
                jnz     @4
                or      bl,bl
                jz      @1
                jmp     VAL40_end

@4:             cmp     dl,'+'
                jnz     @5
                cmp     bl,2
                jnc     VAL40_err
                or      bl,1
                jmp     @1

@5:             cmp     dl,'-'
                jnz     @6
                cmp     bl,2
                jnc     VAL40_err
                or      bl,1
                xor     bh,80h
                jmp     @1
                
@6:             cmp     dl,'&'
                jnz     @7
                cmp     bl,4
                jnc     VAL40_err
                mov     bl,80h
                jmp     @1

@7:             sub     dl,'0'
                jc      VAL40_err
                cmp     dl,10
                jc      @8
                sub     dl,7
                jc      VAL40_err
@8:             cmp     edx,edi
                jnc     VAL40_err
                inc     bh
                or      bl,4
                push    edx
                xchg    eax,ecx
                mul     edi
                xchg    eax,ecx
                push    edx
                mul     edi
                add     ecx,edx
                pop     edx
                adc     edx,0
                pop     edx
                jnz     VAL40_err
                add     eax,edx
                jmp     @1

VAL40_end:      test    bh,7Fh
                jz      VAL40_err
                shl     bh,1
                sbb     ebx,ebx
                jz      VAL40_end_1
                xor     ecx,ebx
                xor     eax,ebx
                sub     eax,ebx
                sbb     ecx,ebx
VAL40_end_1:    inc     dword ptr [esp]
                mov     [esp+4],eax
                mov     [esp+8],ecx
    
VAL40_err:      pop     ecx
                pop     eax
                pop     edx
                mov     ebx,[esp+24]
                or      ebx,ebx
                jz      VAL40_err_1
                dec     ecx
                mov     [ebx],cx

VAL40_err_1:    pop     ebx
                pop     edi
                pop     esi
                ret     12
StrToLarge      endp


; ===============================================================================
; LargeToStr:  convert LARGE to STRING
; -------------------------------------------------------------------------------
;               in      OP1: ->LARGE            Value to convert
;                       FCT:   LONG             Factor to use (2 ... 36)
;                       FLG:   LONG             Bit 0: unsigned/signed
;
;               ret     EAX                     ->BSTR

; Code by Udo Schmidt
; ===============================================================================

LargeToStr      proc
                push    esi
                push    edi
                push    ebx
                mov     ecx,[esp+16]
                mov     ebx,[ecx]
                mov     ecx,[ecx+4]
                xor     esi,esi

                test    byte ptr [esp+24],1
                jz      @0
                mov     esi,ecx
                sar     esi,31
                jz      @0
                xor     ecx,esi
                xor     ebx,esi
                sub     ebx,esi
                sbb     ecx,esi

@0:             push    esi
                lea     edi,[tmpString+99]
                mov     byte ptr [edi],0
                mov     esi,[esp+24]
                cmp     esi,2
                jc      @3
                             
@1:             dec     edi
                xor     edx,edx
                mov     eax,ecx
                div     esi
                mov     ecx,eax
                mov     eax,ebx
                div     esi
                mov     ebx,eax 
                add     dl,30h
                cmp     dl,3Ah
                jc      @2
                add     dl,7
@2:             mov     [edi],dl

                mov     eax,ebx
                or      eax,ecx
                jnz     @1

@3:             pop     esi
                lea     eax,[tmpString+99]
                sub     eax,edi
                sub     eax,esi
                add     edi,esi
                jnc     @4
                mov     byte ptr [edi],'-'
@4:             mov     [edi-4],eax
                mov     eax,edi

                pop     ebx
                pop     edi
                pop     esi
                ret     12
LargeToStr     endp

 ; In den 1. 2 Byte steht der Variablentyp,
 ; die nächsten 2 Byte beinhalten, falls es ein Decimal ist,
 ; Vorzeicheninformation und Zahl der Nachkommastellen, ansonsten sind sie leer.
 ; Ab Byte 5 stehen beim Decimal die hochwertigen 4 Byte.
 ; Ab Byte 9 die niederwertigen 8 Byte. Letzteres gilt auch für Byte, Integer, Long, Double etc. im Variant.
 ; Wenn ein Array im Variant steht, steht am Offset 8 die Adresse des Safearrays.


; ===============================================================================
; LargegToDec   create Variant (decimal) from LARGE
; -------------------------------------------------------------------------------
;               in      pRet: ->Variant         Buffer Set by VB (hidden parameter)
;                       pOP1: ->LARGE           Value To convert
;                       FLG: ->FLAG             Bit 0: Sign flag
;
;               ret     ---                     Variant Set To decimal

; Code by Udo Schmidt
; ===============================================================================


LargeToDec     proc    pRET : dword,           ; ptr: return value
                        pOP1 : dword,           ; ptr: large
                        flg  : dword            ; flg: Bit 0: signed

                variant_type    equ     <word ptr [eax]>
                variant_flag    equ     <word ptr [eax+3]>
                variant_dataLO  equ     <dword ptr [eax+8]>
                variant_dataHI  equ     <dword ptr [eax+12]>
                vbDecimal       equ     14

                mov     eax,pOP1
                mov     ecx,[eax]               ; ecx =  LARGE.lo
                mov     edx,[eax+4]             ; edx =  LARGE.hi
                mov     eax,pRET                ; eax => ret (VARIANT)
                mov     variant_type,vbDecimal   ; set VARIANT type

                or      edx,edx                 ; if LARGE is greater than or equals to zero
                jns     @1                      ; go on at @1

                test    flg,1                   ; the same, if sign flag is not set
                jz      @1

                or      variant_flag,8000h      ; otherwise set the VARIANT's sign flag

@1:             mov     variant_dataLO,ecx      ; set result
                mov     variant_dataHI,edx
                ret
LargeToDec     endp


; ===============================================================================
; VarToLarge    create LARGE from VARIANT
; -------------------------------------------------------------------------------
;               in      OP1: ->VARIANT          Value to convert
;
;               ret     EDX:EAX                 LARGE value (0 on error)
;
;               - Accepts int, lng, sgl, dbl, cur, dat, str, bol, dec, byte

; Code by Udo Schmidt
; ===============================================================================

VarToLarge      proc
                push    ebx                     ; store ebx

                xor     eax,eax                 ; initialize return value
                xor     edx,edx
                mov     ecx,[esp+8]             ; get source Variant
                or      ecx,ecx
                jz      @2                      ; no source, no work -> finish

                movzx   ebx,word ptr [ecx]      ; get Variant's data type
                test    bh,20h
                jnz     @2                      ; if array flag is set -> finish

                add     ecx,8                   ; set ecx to point to Variant's data

                test    bh,40h                  ; if variant was passed by reference ...
                jz      @1
                mov     ecx,[ecx]               ; ... get "real" data pointer

@1:             xor     bh,bh
                cmp     bl,18                   
                jnc     @2                      ; data type greater than 17 -> finish

                call    dword ptr [4*ebx+(offset GET40_vct -4)] ; jump to acc. routine

@2:             pop     ebx                     ; restore ebx
                ret     4                       ; return


GET40_byt:      mov     al,[ecx]                ; "BYTE": set AL only
                ret

GET40_int:
GET40_bol:      movsx   eax,word ptr [ecx]      ; "INTEGER/BOOLEAN": get signed AX
                cdq                             ; adapt edx acc.
                ret

GET40_cur:      mov     eax,[ecx]               ; "CURRENCY": get EDX:EAX and divide by 10000
                mov     edx,[ecx+4]
                xor     ecx,ecx
                mov     ebx,10000
                jmp     rouDIV

GET40_lng:      mov     eax,[ecx]               ; "LONG": get EAX
                cdq                             ; adapt edx acc.
GET40_dmy:      ret

GET40_sgl:      fld     dword ptr [ecx]         ; "SINGLE": load data into FPU
GET40_sgl1:     push    edx                     ; reduce stack by 2 DWords
                push    eax
                fistp   qword ptr [esp]         ; put FPU onto stack
                pop     eax                     ; get result from stack
                pop     edx
                ret
                
GET40_str:      mov     ecx,[ecx]               ; "STRING"
                jecxz   GET40_dmy               ; no string pointer => finish
                cmp     dword ptr [ecx-4],8     ; string too short   => finish
                jc      GET40_dmy

GET40_var:      mov     eax,[ecx]               ; "VARIANT": get EDX:EAX
                mov     edx,[ecx+4]
                ret

GET40_dat:                                      ; "DOUBLE/DATE"
GET40_dbl:      fld     qword ptr [ecx]         ; load value into FPU
                jmp     GET40_sgl1              ; go on at "SINGLE"

GET40_dec:      mov     eax,[ecx]               ; "DECIMAL"
                mov     edx,[ecx+4]             ; get EDX:EAX
                mov     cl,[ecx-5]
                shl     cl,1
                sbb     ecx,ecx
                jz      GET40_dmy               ; -> finish, if sign flag is not set

                xor     eax,ecx                 ; NEG EDX:EAX otherwise
                xor     edx,ecx
                sub     eax,ecx
                sbb     edx,ecx
                ret
                                                
                               
GET40_vct       dd      GET40_dmy,GET40_int,GET40_lng,GET40_sgl
                dd      GET40_dbl,GET40_cur,GET40_dat,GET40_str
                dd      GET40_dmy,GET40_dmy,GET40_bol,GET40_var
                dd      GET40_dmy,GET40_dec,GET40_dmy,GET40_dmy
                dd      GET40_byt                

VarToLarge      endp



; ***************************************************************************
; *             THE FOLLOWING FUNCTIONS ARE NOT EXPORTED                    *
; ***************************************************************************


; ===========================================================================
;  NAME: uDiv 
;  DESC: Dividier-Routine für 64 Bit Unsigned Integer
;  RET(EDX:EAX) : Quotient

;  Code from AMD's Code Optimization Guide (Efficient 64-Bit Integer Arithmetic)
; ===========================================================================

uDiv proc   DIVISOR_HI  : dword,    ; Divisor Hi    : ESP+4
            DIVISOR_LO  : dword,    ; Dviisor Lo    : ESP+8
            DIVIDENT_HI : dword,    ; Devident Hi   : ESP+0C
            DIVIDENT_LO : dword,    ; Divident Lo   : ESP+10

    ;OUTPUT: EDX:EAX quotient of division
    ;
    ;DESTROYS: EAX,ECX,EDX,EFlags
  
    push ebx                ;save EBX as per calling convention
    mov ecx, DIVISOR_HI     ;divisor_hi
    mov ebx, DIVISOR_LO     ;divisor_lo
    mov edx, DIVIDENT_HI    ;dividend_hi
    mov eax, DIVIDENT_LO    ;dividend_lo
    test ecx, ecx           ;divisor > 2^32-1?
    jnz $big_divisor        ;yes, divisor > 32^32-1
    
    cmp edx, ebx            ;only one division needed? (ECX = 0)
    jae $two_divs           ;need two divisions
    div ebx                 ;EAX = quotient_lo
    mov edx, ecx            ;EDX = quotient_hi = 0 (quotient in EDX:EAX)
    pop ebx                 ;restore EBX as per calling convention
    ret                     ;done, return to caller
    
  $two_divs:
    mov ecx, eax            ;save dividend_lo in ECX
    mov eax, edx            ;get dividend_hi
    xor edx, edx            ;zero extend it into EDX:EAX
    div ebx                 ;quotient_hi in EAX
    xchg eax, ecx           ;ECX = quotient_hi, EAX = dividend_lo
    div ebx                 ;EAX = quotient_lo
    mov edx, ecx            ;EDX = quotient_hi (quotient in EDX:EAX)
    pop ebx                 ;restore EBX as per calling convention
    ret
    
  $big_divisor:
    push edi                ;save EDI as per calling convention
    mov edi, ecx            ;save divisor_hi
    shr edx, 1              ;shift both divisor and dividend right
    rcr eax, 1              ; by 1 bit
    ror edi, 1
    rcr ebx, 1
    bsr ecx, ecx                ;ECX = number of remaining shifts
    shrd ebx, edi, cl           ;scale down divisor and dividend
    shrd eax, edx, cl           ;such that divisor is
    shr edx, cl                 ;less than 2^32 (i.e. fits in EBX)
    rol edi, 1                  ;restore original divisor_hi
    div ebx                     ;compute quotient
    mov ebx, DIVIDENT_LO        ;dividend_lo
    mov ecx, eax                ;save quotient
    imul edi, eax               ;quotient * divisor hi-word (low only)
    mul dword ptr [DIVISOR_LO]  ;quotient * divisor lo-word
    add edx, edi                ;EDX:EAX = quotient * divisor
    sub ebx, eax                ;dividend_lo - (quot.*divisor)_lo
    mov eax, ecx                ;get quotient
    mov ecx, DIVIDENT_HI        ;dividend_hi
    sbb ecx, edx                ;subtract divisor * quot. from dividend
    sbb eax, 0                  ;adjust quotient if remainder negative
    xor edx, edx                ;clear hi-word of quot(EAX<=FFFFFFFFh)
    pop edi                     ;restore EDI as per calling convention
    pop ebx                     ;restore EBX as per calling convention
  
    ret                         ;done, return to caller
uDiv endp



; ===========================================================================
;  NAME: sDiv 
;  DESC: Dividier-Routine für 64 Bit Signed Integer
;  RET(EDX:EAX) : Quotient

;  Code from AMD's Code Optimization Guide (Efficient 64-Bit Integer Arithmetic)
; ===========================================================================

sDiv proc   DIVISOR_HI  : dword,    ; Divisor Hi    : ESP+4
            DIVISOR_LO  : dword,    ; Dviisor Lo    : ESP+8
            DIVIDENT_HI : dword,    ; Devident Hi   : ESP+0C
            DIVIDENT_LO : dword,    ; Divident Lo   : ESP+10

    ; OUTPUT: EDX:EAX quotient of division
    ;
    ; DESTROYS: EAX,ECX,EDX,EFlags
    
    push ebx                ;save EBX as per calling convention
    push esi                ;save ESI as per calling convention
    push edi                ;save EDI as per calling convention
    mov ecx, DIVISOR_HI     ;divisor-hi
    mov ebx, DIVISOR_LO     ;divisor-lo
    mov edx, DIVIDENT_HI    ;dividend-hi
    mov eax, DIVIDENT_LO    ;dividend-lo
    mov esi, ecx            ;divisor-hi
    xor esi, edx            ;divisor-hi ^ dividend-hi
    sar esi, 31             ;(quotient < 0) ? -1 : 0
    mov edi, edx            ;dividend-hi
    sar edi, 31             ;(dividend < 0) ? -1 : 0
    xor eax, edi            ;if (dividend < 0)
    xor edx, edi            ;compute 1's complement of dividend
    sub eax, edi            ;if (dividend < 0)
    sbb edx, edi            ;compute 2's complement of dividend
    mov edi, ecx            ;divisor-hi
    sar edi, 31             ;(divisor < 0) ? -1 : 0
    xor ebx, edi            ;if (divisor < 0)
    xor ecx, edi            ;compute 1's complement of divisor
    sub ebx, edi            ;if (divisor < 0)
    sbb ecx, edi            ;compute 2's complement of divisor
    jnz $big_divisor        ;divisor > 2^32-1
    cmp edx, ebx            ;only one division needed ? (ECX = 0)
    jae $two_divs           ;need two divisions
    div ebx                 ;EAX = quotient-lo
    mov edx, ecx            ;EDX = quotient-hi = 0
    ; (quotient in EDX:EAX)
    xor eax, esi            ;if (quotient < 0)
    xor edx, esi            ;compute 1's complement of result
    sub eax, esi            ;if (quotient < 0)
    sbb edx, esi            ;compute 2's complement of result
    pop edi                 ;restore EDI as per calling convention
    pop esi                 ;restore ESI as per calling convention
    pop ebx                 ;restore EBX as per calling convention
    ret                     ;done, return to caller
    
  $two_divs:
    mov ecx, eax                ;save dividend-lo in ECX
    mov eax, edx                ;get dividend-hi
    xor edx, edx                ;zero extend it into EDX:EAX
    div ebx                     ;quotient-hi in EAX
    xchg eax, ecx               ;ECX = quotient-hi, EAX = dividend-lo
    div ebx                     ;EAX = quotient-lo
    mov edx, ecx                ;EDX = quotient-hi
    ; (quotient in EDX:EAX)
    jmp $make_sign              ;make quotient signed
    
  $big_divisor:
    mov DIVIDENT_LO, eax        ;dividend-lo
    mov DIVISOR_LO, ebx         ;divisor-lo
    mov DIVIDENT_HI, edx        ;dividend-hi
    mov edi, ecx                ;save divisor-hi
    shr edx, 1                  ;shift both
    rcr eax, 1                  ;divisor and
    ror edi, 1                  ;and dividend
    rcr ebx, 1                  ;right by 1 bit
    bsr ecx, ecx                ;ECX = number of remaining shifts
    shrd ebx, edi, cl           ;scale down divisor and
    shrd eax, edx, cl           ;dividend such that divisor
    shr edx, cl                 ;less than 2^32 (i.e. fits in EBX)
    rol edi, 1                  ;restore original divisor-hi
    div ebx                     ;compute quotient
    mov ebx, DIVIDENT_LO        ;dividend-lo
    mov ecx, eax                ;save quotient
    imul edi, eax               ;quotient * divisor hi-word (low only)
    mul dword ptr [DIVISOR_LO]  ;quotient * divisor lo-word
    add edx, edi                ;EDX:EAX = quotient * divisor
    sub ebx, eax                ;dividend-lo - (quot.*divisor)-lo
    mov eax, ecx                ;get quotient
    mov ecx, DIVIDENT_HI        ;dividend-hi
    sbb ecx, edx                ;subtract divisor * quot. from dividend
    sbb eax, 0                  ;adjust quotient if remainder negative
    xor edx, edx                ;clear hi-word of quotient
  $make_sign:
    xor eax, esi                ;if (quotient < 0)
    xor edx, esi                ;compute 1's complement of result
    sub eax, esi                ;if (quotient < 0)
    sbb edx, esi                ;compute 2's complement of result
    pop edi                     ;restore EDI as per calling convention
    pop esi                     ;restore ESI as per calling convention
    pop ebx                     ;restore EBX as per calling convention
    ret                         ;done, return to caller
sDiv endp


; ===========================================================================
;  NAME: uMOD 
;  DESC: Modulo Division für 64 Bit Unsigned Integer
;  RET(EDX:EAX) : Divisionsrest (Remainder)

;  Code from AMD's Code Optimization Guide (Efficient 64-Bit Integer Arithmetic)
; ===========================================================================

uMOD proc   DIVISOR_HI  : dword,    ; Divisor Hi    : ESP+4
            DIVISOR_LO  : dword,    ; Divisor Lo    : ESP+8
            DIVIDENT_HI : dword,    ; Devident Hi   : ESP+0C
            DIVIDENT_LO : dword,    ; Divident Lo   : ESP+10

    ; divides two unsigned 64-bit integers, and returns
    ; the remainder.
    ;
    ;OUTPUT: EDX:EAX remainder of division
    ;
    ;DESTROYS: EAX,ECX,EDX,EFlags
    
    push ebx                    ;save EBX as per calling convention
    mov ecx, DIVISOR_HI         ;divisor_hi
    mov ebx, DIVISOR_LO         ;divisor_lo
    mov edx, DIVIDENT_HI        ;dividend_hi
    mov eax, DIVIDENT_LO        ;dividend_lo
    test ecx, ecx               ;divisor > 2^32-1?
    jnz $r_big_divisor          ;yes, divisor > 32^32-1
    
    cmp edx, ebx                ;only one division needed? (ECX = 0)
    jae $r_two_divs             ;need two divisions
    div ebx                     ;EAX = quotient_lo
    mov eax, edx                ;EAX = remainder_lo
    mov edx, ecx                ;EDX = remainder_hi = 0
    pop ebx                     ;restore EBX as per calling convention
    ret                         ;done, return to caller
    
  $r_two_divs:
    mov ecx, eax                ;save dividend_lo in ECX
    mov eax, edx                ;get dividend_hi
    xor edx, edx                ;zero extend it into EDX:EAX
    div ebx                     ;EAX = quotient_hi, EDX = intermediate
    
    ; remainder
    mov eax, ecx                ;EAX = dividend_lo
    div ebx                     ;EAX = quotient_lo
    mov eax, edx                ;EAX = remainder_lo
    xor edx, edx                ;EDX = remainder_hi = 0
    pop ebx                     ;restore EBX as per calling convention
    ret                         ;done, return to caller
    
  $r_big_divisor:
    push edi                    ;save EDI as per calling convention
    mov edi, ecx                ;save divisor_hi
    shr edx, 1                  ;shift both divisor and dividend right
    rcr eax, 1                  ; by 1 bit
    ror edi, 1
    rcr ebx, 1
    bsr ecx, ecx                ;ECX = number of remaining shifts
    shrd ebx, edi, cl           ;scale down divisor and dividend such
    shrd eax, edx, cl           ; that divisor is less than 2^32
    shr edx, cl                 ; (i.e. fits in EBX)
    rol edi, 1                  ;restore original divisor (EDI:ESI)
    div ebx                     ;compute quotient
    mov ebx, DIVIDENT_LO        ;dividend lo-word
    mov ecx, eax                ;save quotient
    imul edi, eax               ;quotient * divisor hi-word (low only)
    mul dword ptr DIVISOR_LO    ;quotient * divisor lo-word
    add edx, edi                ;EDX:EAX = quotient * divisor
    sub ebx, eax                ;dividend_lo - (quot.*divisor)-lo
    mov ecx, DIVIDENT_HI        ;dividend_hi
    mov eax, DIVISOR_LO         ;divisor_lo
    sbb ecx, edx                ;subtract divisor * quot. from
    ; dividend
    sbb edx, edx                ;(remainder < 0)? 0xFFFFFFFF : 0
    and eax, edx                ;(remainder < 0)? divisor_lo : 0
    and edx, DIVISOR_HI         ;(remainder < 0)? divisor_hi : 0
    add eax, ebx                ;remainder += (remainder < 0)?
    adc edx, ecx                ; divisor : 0
    pop edi                     ;restore EDI as per calling convention
    pop ebx                     ;restore EBX as per calling convention
    ret                         ;done, return to caller
uMOD endp


; ===========================================================================
;  NAME: sMOD 
;  DESC: Modulo Division für 64 Bit Signed Integer
;  RET(EDX:EAX) : Divisionsrest (Remainder)

;  Code from AMD's Code Optimization Guide (Efficient 64-Bit Integer Arithmetic)
; ===========================================================================

sMOD proc   DIVISOR_HI  : dword,    ; Divisor Hi    : ESP+4
            DIVISOR_LO  : dword,    ; Divisor Lo    : ESP+8
            DIVIDENT_HI : dword,    ; Devident Hi   : ESP+0C
            DIVIDENT_LO : dword,    ; Divident Lo   : ESP+10

    ; divides two signed 64-bit numbers and returns the remainder
    
    ; OUTPUT: EDX:EAX remainder of division
    
    ; DESTROYS: EAX,ECX,EDX,EFlags
    
    push ebx                    ;save EBX as per calling convention
    push esi                    ;save ESI as per calling convention
    push edi                    ;save EDI as per calling convention
    mov ecx, DIVISOR_HI         ;divisor-hi
    mov ebx, DIVISOR_LO         ;divisor-lo
    mov edx, DIVIDENT_HI        ;dividend-hi
    mov eax, DIVIDENT_LO        ;dividend-lo
    mov esi, edx                ;sign(remainder) == sign(dividend)
    sar esi, 31                 ;(remainder < 0) ? -1 : 0
    mov edi, edx                ;dividend-hi
    sar edi, 31                 ;(dividend < 0) ? -1 : 0
    xor eax, edi                ;if (dividend < 0)
    xor edx, edi                ;compute 1's complement of dividend
    sub eax, edi                ;if (dividend < 0)
    sbb edx, edi                ;compute 2's complement of dividend
    mov edi, ecx                ;divisor-hi
    sar edi, 31                 ;(divisor < 0) ? -1 : 0
    xor ebx, edi                ;if (divisor < 0)
    xor ecx, edi                ;compute 1's complement of divisor
    sub ebx, edi                ;if (divisor < 0)
    sbb ecx, edi                ;compute 2's complement of divisor
    jnz $sr_big_divisor         ;divisor > 2^32-1
    
    cmp edx, ebx                ;only one division needed ? (ECX = 0)
    jae $sr_two_divs            ;nope, need two divisions
    div ebx                     ;EAX = quotient_lo
    mov eax, edx                ;EAX = remainder_lo
    mov edx, ecx                ;EDX = remainder_lo = 0
    xor eax, esi                ;if (remainder < 0)
    xor edx, esi                ;compute 1's complement of result
    sub eax, esi                ;if (remainder < 0)
    sbb edx, esi                ;compute 2's complement of result
    pop edi                     ;restore EDI as per calling convention
    pop esi                     ;restore ESI as per calling convention
    pop ebx                     ;restore EBX as per calling convention
    ret                         ;done, return to caller
  
  $sr_two_divs:
    mov ecx, eax                ;save dividend_lo in ECX
    mov eax, edx                ;get_dividend_hi
    xor edx, edx                ;zero extend it into EDX:EAX
    div ebx                     ;EAX = quotient_hi,
    ;EDX = intermediate remainder
    mov eax, ecx                ;EAX = dividend_lo
    div ebx                     ;EAX = quotient_lo
    mov eax, edx                ;remainder_lo
    xor edx, edx                ;remainder_hi = 0
    jmp $sr_makesign            ;make remainder signed
  
  $sr_big_divisor:
    ; lokale Variablen werden nicht benötigt, dafür wird in Orignalvariablen geschrieben
    ;SUB ESP, 16                 ;create three local variables
    mov DIVIDENT_LO, eax        ;dividend_lo
    mov DIVISOR_LO, ebx         ;divisor_lo
    mov DIVIDENT_HI, edx        ;dividend_hi
    mov DIVISOR_HI, ecx         ;divisor_hi
    mov edi, ecx                ;save divisor_hi
    shr edx, 1                  ;shift both
    rcr eax, 1                  ;divisor and
    ror edi, 1                  ;and dividend
    rcr ebx, 1                  ;right by 1 bit
    bsr ecx, ecx                ;ECX = number of remaining shifts
    shrd ebx, edi, cl           ;scale down divisor and
    shrd eax, edx, cl           ;dividend such that divisor
    shr edx, cl                 ;less than 2^32 (i.e. fits in EBX)
    rol edi, 1                  ;restore original divisor_hi
    div ebx                     ;compute quotient
    mov ebx, DIVIDENT_LO        ;dividend_lo
    mov ecx, eax                ;save quotient
    imul edi, eax               ;quotient * divisor hi-word (low only)
    mul dword ptr [DIVISOR_LO]  ;quotient * divisor lo-word
    add edx, edi                ;EDX:EAX = quotient * divisor
    sub ebx, eax                ;dividend_lo - (quot.*divisor)-lo
    mov ecx, DIVIDENT_HI        ;dividend_hi
    sbb ecx, edx                ;subtract divisor * quot. from dividend
    sbb eax, eax                ;remainder < 0 ? 0xffffffff : 0
    mov edx, DIVISOR_HI         ;divisor_hi
    and edx, eax                ; remainder < 0 ? divisor_hi : 0
    and eax, DIVISOR_LO         ;remainder < 0 ? divisor_lo : 0
    add eax, ebx                ;remainder_lo
    add edx, ecx                ;remainder_hi
    ;ADD ESP, 16                ;remove local variables
    
  $sr_makesign:
    xor eax, esi                ;if (remainder < 0)
    xor edx, esi                ;compute 1's complement of result
    sub eax, esi                ;if (remainder < 0)
    sbb edx, esi                ;compute 2's complement of result
    pop edi                     ;restore EDI as per calling convention
    pop esi                     ;restore ESI as per calling convention
    pop ebx                     ;restore EBX as per calling convention
    ret                         ;done, return to caller
sMOD endp

; ===============================================================================
; rouDIV        devide LARGEs
; -------------------------------------------------------------------------------
;               in      EDX:EAX                 divident
;                       ECX:EBX                 divisior
;
;               out     EDX:EAX                 quotient
;                       ECX:EBX                 remainder

; Code by Udo Schmidt
; ===============================================================================
rouDIV          proc    near
                or      ecx,ecx
                jnz     @2
                cmp     edx,ebx
                jc      @1
                xchg    ecx,edx
                xchg    eax,ecx
                div     ebx
                xchg    eax,ecx
@1:             div     ebx
                mov     ebx,edx
                mov     edx,ecx
                xor     ecx,ecx
                ret

@2:             push    0
                push    edi
                push    ebx
                push    ecx
                push    edx
                push    eax
                push    ebx
                mov     edi,ecx
                bsr     ecx,ecx
                cmp     cl,31
                jnz     @3
                mov     ebx,edi
                mov     eax,edx
                xor     edx,edx
                jmp     @4

@3:             inc     ecx
                shrd    ebx,edi,cl
                shrd    eax,edx,cl
                shr     edx,cl
                
@4:             div     ebx
                mov     ebx,eax
                imul    edi,eax
                pop     eax
                mul     ebx
                add     edi,edx
                
                pop     edx
                sub     edx,eax
                                
                pop     eax
                pop     ecx
                sbb     eax,edi
                sbb     edi,edi
                add     ebx,edi
                xchg    ebx,[esp]
                and     ecx,edi
                and     ebx,edi
                add     ebx,edx
                adc     ecx,eax
                pop     eax
                pop     edi
                pop     edx
                ret                                
rouDIV          endp

end LibMain