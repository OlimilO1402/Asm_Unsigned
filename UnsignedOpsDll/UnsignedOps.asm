include C:\masm32\include\masm32rt.inc 
.data?
    hInstance dd ?
	
.code

DllEntry proc instance: dword, reason: dword, unused: dword
    
    .if reason == DLL_PROCESS_ATTACH
        
        mrm hInstance, instance       ; copy local to global
        mov eax, TRUE                 ; return TRUE so DLL will start
        
    .elseif reason == DLL_PROCESS_DETACH
		;
    .elseif reason == DLL_THREAD_ATTACH
		;
    .elseif reason == DLL_THREAD_DETACH
		;
    .endif
	
    ret
	
DllEntry endp

;123456789012345 * 123456789012345
;           = 15241578753238669120562399025
;uint64_max = 18446744073709551615
; int64_max =  9223372036854775807

;The manual to Intel assembler syntax can be found here:
;https://software.intel.com/content/dam/develop/public/us/en/documents/325462-sdm-vol-1-2abcd-3abcd.pdf


OPTION PROLOGUE:NONE
OPTION EPILOGUE:NONE

;Align 8

; ARITHMETIC Operations
; --------========  Unsigned Int16 operations  ========--------
UInt16_Add proc 
    
    mov  ax, [esp+4]  ; copy the first  uint16 value from stack to register AX
    mov  cx, [esp+8]  ; copy the second uint16 value from Stack to register CX
    add  ax,  cx      ; add the value in CX to the value in register AX
                      ; the return value of a function in VB must be in register EAX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt16_Add endp

UInt16_Sub proc 
    
    mov  ax, [esp+4]  ; copy the first  uint16 value from stack to register AX
    mov  cx, [esp+8]  ; copy the second uint16 value from Stack to register CX
    sub  ax,  cx      ; subtract the value in CX from the value in register AX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt16_Sub endp

UInt16_Mul proc 
    
    mov  ax, [esp+4]  ; copy the first  uint16 value from stack to register AX
    mov  cx, [esp+8]  ; copy the second uint16 value from Stack to register CX
    mul  cx           ; multipliy the value in register CX with the value in register EAX
                      ; the result as UInt32 is returned in register EAX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt16_Mul endp

UInt16_Div proc
    
    mov  ax, [esp+4]  ; copy the first  uint16 value from stack to register AX (Dividend)
    mov  cx, [esp+8]  ; copy the second uint16 value from Stack to register CX (Divisor )
                      ; before div we must ensure that register edx is empty 
                      ; you can do this either using mov, sub or xor
    xor edx, edx      ; xor: null all ones if there is one
    div  cx           ; divide the value in register AX by the value in register CX 
                      ; the quotient is in EAX the remainder (rest) in EDX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt16_Div endp


; --------========  Unsigned Int32 operations  ========--------

;page 605
UInt32_Add proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    add eax, ecx      ; add the value in ECX to the value in register EAX
                      ; the return value of a function in VB must be in register EAX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Add endp

;page 1857
UInt32_Sub proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    sub eax, ecx      ; subtract the value in ECX from the value in register EAX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Sub endp

;page 1332
UInt32_Mul proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    mul ecx           ; multipliy the value in register ECX with the value in register EAX
                      ; the result as UInt64 is returned in EDX:EAX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Mul endp

;page 891
UInt32_Div proc
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX (Dividend)
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX (Divisor )
                      ; before div we must ensure that register edx is empty 
                      ; you can do this either using mov, sub or xor
    xor edx, edx      ; xor: null all ones if there is one
    div ecx           ; divide the value in register EAX by the value in register ECX 
                      ; the quotient is in EAX the remainder (rest) in EDX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Div endp


; just as an example, for "ByRef"
UInt32_Add_ref proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    mov eax, [eax]    ; dereference the pointer in EAX and copy den value back to EAX
    add eax, [ecx]    ; dereference the pointer in ECX und add the value from it
                      ;	to the value in register EAX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Add_ref endp

; --------========  Unsigned Int64 operations  ========--------

;http://masm32.com/board/index.php?topic=5264.0
;https://www.jj2007.eu/Masm32_Tips_Tricks_and_Traps.htm
;https://foren.activevb.de/archiv/vb-classic/thread-404293/beitrag-404304/Re-Errata/

;pages 605 + 600
UInt64_Add proc
    
    mov eax, [esp+4]   ; copy the lower part of the first uint64 value from stack to register EAX
    mov edx, [esp+8]   ; copy the upper part of the first uint64 value from Stack to register EDX
    add eax, [esp+12]  ; add the lower part of the second uint64 value from stack to the value in register EAX 
    adc edx, [esp+16]  ; add the upper part of the second uint64 value from stack to the value in register EDX by using the carry flag
    ret 16             ; return to caller, remove 16 bytes from stack (->stdcall)
    
UInt64_Add endp

;pages 1857 + 1789
UInt64_Sub proc
    
    mov eax, [esp+4]   ; copy the lower part of the first uint64 value from stack to register EAX
    mov edx, [esp+8]   ; copy the upper part of the first uint64 value from Stack to register EDX
    sub eax, [esp+12]  ; subtract the lower part of the second uint64 value from stack to the value in register EAX 
    sbb edx, [esp+16]  ; subtract the upper part of the second uint64 value from stack to the value in register EDX by using the borrow flag
    ret 16
    
UInt64_Sub endp

;pages 1857 + 1789
;https://www.plantation-productions.com/Webster/www.artofasm.com/Windows/HTML/AdvancedArithmetica2.html#1007619
;https://stackoverflow.com/questions/87771/how-can-i-multiply-two-64-bit-numbers-using-x86-assembly-language
;https://docs.microsoft.com/en-us/cpp/cpp/argument-passing-and-naming-conventions?view=msvc-170
;https://docs.microsoft.com/en-us/windows/win32/api/wtypes/ns-wtypes-decimal-r1
;                ByVal V1 As Currency, ByVal V2 As Currency, ByRef Dec_out As Variant; is the same as:
;                ByVal V1 As Currency, ByVal V2 As Currency, ByVal pDec_out As LongPtr
;                      V1: 8-Byte,           V2: 8-Byte,           pDec_out: 4-Byte
;                multiplier(64)    , multiplicand(64)  , pDestination(32)
;                ByVal mp As UInt64, ByVal mc as Uint64, ByRef prd As UInt128
;                ByVal V1 As Currency, ByVal V2 As Currency, ByVal pDec_out As LongPtr
UInt64_Mul proc
    
    ;int 3
    push  ebp               ; save the register ebp to the stack
    mov   ebp    ,  esp     ; save the current stack pointer to register EBP
    push  edi               ; save the register edi to the stack
    ; 1. prepare Variant as type Decimal
    mov   edi    , [ebp+24] ; copy the pointer to a Variant from stack to register EDI
    mov    dx    ,      14  ; copy the vartype-word for decimal it is 14 = &HE to the register dx	
    mov  [edi+ 0],   dx     ; copy the value in register dx to the Variant to the first int16-slot
    mov    dl    ,       0  ; copy the number of precision it is 0 to the register dl (->integer, otherwise between 0 and 255)
    mov    dh    ,       0  ; copy the sign-word it is 0 to the register dh (->unsigned, otherwise &H80 is negative sign)
    mov  [edi+ 2],   dx     ; copy the value in register dx to the Variant into the second int16-slot
    push  ebx               ; save the register ebx to the stack
    ;mit 0 vorbelegen
    ;mov   ecx    ,       0  ; copy 0 to the register ECX
    ;mov  [edi+ 4],  ecx     ; copy the register ECX to the Variant into the second Int32 slot 
    ;mov  [edi+ 8],  ecx
    ;mov  [edi+12],  ecx	
	
    ; 2. Multiply the Lo dword of multiplier times Lo dword of multiplicand.  
    mov   eax    , [ebp+ 8] ; copy the Lo dword of multiplier to register EAX
    mul  dword ptr [ebp+16] ; Multiply Lo dwords EAX * ECX    
    mov  [edi+ 8],  eax     ; Save Lo dword of product to the third Variant slot
    mov   ecx    ,  edx     ; Save Hi dword of product into the register EBX
    
     ; 3. Multiply Lo dword of multiplier times Hi dword of Multiplicand
    mov   eax    , [ebp+ 8] ; copy the lo dword of multiplier to register EAX
    mul  dword ptr [ebp+20] ; multiply register EAX times Hi dword of Multiplicand
    add   eax    ,  ecx     ; Add to the partial product
    adc   edx    ,       0  ; add the carryflag to EDX which contains the hi dword of the multiplication-result
    mov   ebx    ,  eax     ; Save partial product for now
    mov   ecx    ,  edx     ;
    
    ; 4. Multiply the Hi dword of multiplier times Lo dword of Multiplicand
    mov   eax    , [ebp+12] ; Get Hi dword of Multiplier 
    mul  dword ptr [ebp+16] ; Multiply by Lo dword of Multiplicand
    add   eax    ,  ebx     ; Add to the partial product. 
    mov  [edi+12],  eax     ; Save the partial product.    
    adc   ecx    ,  edx     ; Add in the carry!
    pushfd                  ; Save carry out here.
    
    ; 5. Multiply the two Hi dwords together,
    ;mov   eax    , [esp+ 8] ; Get Hi dword of Multiplier
    mov   eax    , [ebp+12]
    ;mul  dword ptr [esp+16] ; Multiply by Hi dword of Multiplicand
    mul  dword ptr [ebp+20]
    popfd                   ; Retrieve carry from above
    adc   eax    ,  ecx     ; Add in partial product from above. 
    adc   edx    ,       0  ; Don't forget the carry!
    mov  [edi+ 4],  eax     ; Save the partial product.
    ;mov  [edi+ 0],  edx     ; Nope! sorry we only have 96 Bit
    pop   ebx	
    pop   edi               ; recover register edi
    pop   ebp               ; recover register ebp
    ret       20            ; return remove 20 bytes from call stack
    
UInt64_Mul endp

;OK, jetzt muss ich noch die Additionen erledigen und 3 uint32 in den Variant moven (Decimal=96bit).
;ich überleg wie man das Ganze noch verbessern kann, folgende Fragen sausen mir durch den Kopf
;* wieviele lokale Variablen bzw zusätzlichen Stackspeicher braucht man mindestens
;* was macht man am besten wann oder zuerst,
;* wie verwendet man die Register und lokale variablen am sinnvollsten
;also lassen wir uns die Sachen nochmal durch den Kopf gehen.
;wir haben Grundsätzlich folgende Register zur Verfügung
;* EAX
;* ECX
;* EDX
;* EBX
;* EBP erhält den Zeiger für die Übergabeparameter
;* ESP ist dann der Zeiger auf die lokalen Variablen
;EAX, ECX und EDX werden bisher zum Multiplizieren (und Addieren) verwendet.
;dann könnte man doch EBX verwenden um gleich von vornherein, noch bevor man lokale Variablen anlegt den Zeiger auf den Variant zu halten, oder würde da irgendwas dagegen sprechen?



;UInt64_Add proc
;    
;    mov eax, [esp+4]   ; copy the lower part of the first uint64 value from stack to register EAX
;    mov edx, [esp+8]   ; copy the upper part of the first uint64 value from Stack to register EDX
;    add eax, [esp+12]  ; add the lower part of the second uint64 value from stack to the value in register EAX 
;    adc edx, [esp+16]  ; add the upper part of the second uint64 value from stack to the value in register EDX by using the carry flag
;    ret 16             ; return to callee, remove 16 bytes from stack (->stdcall)
;    
;UInt64_Add endp

UInt64_Div proc
	;
UInt64_Div endp


;               v1: UInt32 (4-Byte), pLng_out: 4Byte
UInt64_Test proc
    
    ; open a stackframe
    mov   ebp    ,  esp       ; copy the current stack pointer in esp to ebp 
    mov   eax    , [esp+ 4]
    mov  [eax+ 0],   dx
    mov   edx    ,    1
    mov  [eax+ 4],  edx
    mov   edx    ,    2
    mov  [eax+ 8],  edx
    mov   edx    ,    3
    mov  [eax+12],  edx
    
    ret 8
	
UInt64_Test endp






; ADDITIONAL Operations
; ------======  Unsigned Int32 shifting operations  ======------

;Thanks Creel
;https://www.youtube.com/watch?v=wXGZ7o9HNss
UInt32_Shl proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    shl eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Shl endp 

UInt32_Shld proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ebx, [esp+8]  ; copy the first  uint32 value from stack to register EBX
    mov ecx, [esp+12] ; copy the second uint32 value from Stack to register ECX
    shld eax, ebx, cl ; shift the value in register EAX about the value in the register CL
    mov ecx, ebx
    ret 12            ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Shld endp 

UInt32_Shr proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    shr eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Shr endp 

UInt32_Shrd proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ebx, [esp+8]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+12] ; copy the second uint32 value from Stack to register ECX
    shrd eax, ebx, cl ; shift the value in register EAX about the value in the register CL
    mov ecx, ebx
    ret 12            ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Shrd endp 

UInt32_Sar proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    sar eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Sar endp 

UInt32_Rol proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    rol eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Rol endp 

UInt32_Rcl proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    rcl eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Rcl endp 

UInt32_Ror proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    ror eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Ror endp 

UInt32_Rcr proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    rcr eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Rcr endp 

; --------======== Boolean Operations  ========-------- ;
;True-table for And
; A B  Out
; 0 0  0
; 0 1  0
; 1 0  0
; 1 1  1
UInt32_And proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    and eax, ecx      ; AND the value in register EAX with the value in the register ECX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_And endp 

;True-table for Or
; A B  Out
; 0 0  0
; 0 1  1
; 1 0  1
; 1 1  1
UInt32_Or proc 
    
    mov eax, [esp+4]  ; copy the first  uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    or  eax, ecx      ; OR the value in register EAX with the value in the register ECX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_Or endp 

;True-table for Not
; A  Out
; 0  1
; 1  0
UInt32_Not proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    not eax           ; NOT the value in register EAX, every Bit gets flipped
    ret 4             ; return to caller, remove 4 bytes from stack (->stdcall)
    
UInt32_Not endp 

;True-table for XOr
; A B  Out
; 0 0  0
; 0 1  1
; 1 0  1
; 1 1  0
UInt32_XOr proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    xor eax, ecx      ; XOR the value in register EAX with the value in the register ECX
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_XOr endp 

;True-table for XNOr
; A B  Out
; 0 0  1
; 0 1  0
; 1 0  0
; 1 1  1
UInt32_XNOr proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    xor eax, ecx      ; XOR the value in register EAX with the value in the register ECX
    not eax           ; NOT the value in register EAX, every bit gets flipped
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_XNOr endp 

;True-table for NOr
; A B  Out
; 0 0  1
; 0 1  0
; 1 0  0
; 1 1  0
UInt32_NOr proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    or  eax, ecx      ; OR the value in register EAX with the value in the register ECX	
    not eax           ; NOT the value in register EAX, every bit gets flipped
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_NOr endp 

;True-table for NAnd
; A B  Out
; 0 0  1
; 0 1  1
; 1 0  1
; 1 1  0
UInt32_NAnd proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    and eax, ecx      ; AND the value in register EAX with the value in the register ECX
    not eax           ; NOT the value in register EAX, every bit gets flipped
    ret 8             ; return to caller, remove 8 bytes from stack (->stdcall)
    
UInt32_NAnd endp 


; INPUT / OUTPUT functions
  ; ------======  Unsigned Int16 input/output functions  ======------
;                  v1: UInt16 (4-Byte), pDec_out: 4-Byte

UInt16_ToDec proc
    
    mov   edi    , [esp+ 8]    ; copy the pointer to a Variant from stack to register EAX
    mov    dx    ,      14     ; copy the vartype-word for decimal it is 14 = &HE to the register dx
    mov  [edi+ 0],   dx        ; copy the value in register dx (=vartype) to the Variant to the first int16-slot
    mov    dx    ,       0     ; copy the sign-word, it is 0 as we want to have unsigned of course
    mov  [edi+ 2],   dx        ; copy the value in register dx to the Variant 
    mov   edx    ,       0     ; copy the value 0 to the register EDX
    mov  [edi+ 4],  edx        ; copy the value in register EDX to the Variant to the second int32-slot
    mov    dx    , [esp+ 4]    ; copy the UInt16-value v1 to register DX 
    mov  [edi+ 8],   dx        ; copy the UInt16-Value v1 to the Variant
    mov   edx    ,       0     ; copy the value 0 to the register EDX
    mov  [edi+12], edx         ; copy the value from register EDX to the Variant to the fourth int32-slot
    ret        8
    
UInt16_ToDec endp

;                  v1: UInt32 (4-Byte), pDec_out: 4-Byte
UInt32_ToDec proc
    
    mov   edi    , [esp+ 8]    ; copy the pointer to a Variant from stack to register EAX
    mov    dx    ,      14     ; copy the vartype-word for decimal it is 14 = &HE to the register dx
    mov  [edi+ 0],   dx        ; copy the value in register dx to the Variant to the first int16-slot
    mov    dx    ,       0     ; copy the sign-word, it is 0 as we want to have unsigned of course
    mov  [edi+ 2],   dx        ; copy the value in register dx to the Variant 
    mov   edx    ,       0     ; copy the value 0 to the register EDX
    mov  [edi+ 4],  edx        ; copy the value in register EDX to the Variant to the second int32-slot
    mov   edx    , [esp+ 4]    ; copy the UInt32-value v1 to register EDX 
    mov  [edi+ 8],  edx        ; copy the UInt32-Value v1 to the Variant
    mov   edx    ,       0     ; copy the value 0 to the register EDX
    mov  [edi+12], edx         ; copy the value from register EDX to the Variant to the fourth int32-slot
    ret        8
    
UInt32_ToDec endp

UInt64_ToDec proc
	;
	ret
UInt64_ToDec endp


OPTION EPILOGUE:EpilogueDef
OPTION PROLOGUE:PrologueDef

; ------======  Unsigned Int16 input/output functions  ======------
UInt16_ToStr proc v1: dword, pStr: dword 
    
    invoke crt__ultow, v1, pStr, 10 ; rradix=10
    ret 8
    
UInt16_ToStr endp

UInt16_ToHex proc v1: dword, pStr: dword 
    
    invoke crt__ultow, v1, pStr, 16 ; rradix=16
    ret 8
	
UInt16_ToHex endp

UInt16_ToBin proc v1: dword, pStr: dword 
    
    invoke crt__ultow, v1, pStr, 2  ; rradix=2
    ret 8
	
UInt16_ToBin endp

UInt16_Parse proc pStr: dword, v1: dword
	;invoke crt_strtoul
	invoke crt_wcstoul, pStr, 0, 10, v1
	ret 4
	
UInt16_Parse endp



;https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2013/kk8w4t5t(v=vs.120)
; to string aka to decimal but decimal is narrative with VB.Decimal
UInt32_ToStr proc v1: dword, pStr: dword 
    
    invoke crt__ultow, v1, pStr, 10 ; rradix=10
    ret 8
    
UInt32_ToStr endp

UInt32_ToHex proc v1: dword, pStr: dword 
    
    invoke crt__ultow, v1, pStr, 16 ; rradix=16
    ret 8
    
UInt32_ToHex endp

UInt32_ToBin proc v1: dword, pStr: dword 
    
    invoke crt__ultow, v1, pStr, 2  ; rradix=2
    ret 8
    
UInt32_ToBin endp

;unsigned long wcstoul(
;   const wchar_t *strSource,
;   wchar_t **endptr,
;   int base
;);

UInt32_Parse proc pStr: dword, v1: dword
	;invoke crt_strtoul
	invoke crt_wcstoul, pStr, 0, 10, v1
	ret 4
	
UInt32_Parse endp



UInt64_ToStr proc v1: qword, pStr: dword 
    
    invoke crt__ui64tow, v1, pStr, 10 ; rradix=10
    ret 12
    
UInt64_ToStr endp

UInt64_ToHex proc v1: qword, pStr: dword 
    
    invoke crt__ui64tow, v1, pStr, 16 ; rradix=16
    ret 12
    
UInt64_ToHex endp

UInt64_ToBin proc v1: qword, pStr: dword 
    
    invoke crt__ui64tow, v1, pStr, 2  ; rradix=2
    ret 12
    
UInt64_ToBin endp

UInt64_Parse proc pStr: dword, v1: qword
	;invoke crt_strtoul
	;invoke crt__wtoui64, pStr, 0, 10, v1
	ret 4
UInt64_Parse endp



End DllEntry