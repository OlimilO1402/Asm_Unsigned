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
;=15241578753238669120562399025
;uint64_max = 18446744073709551615
; int64_max = 9223372036854775807

;The manual to Intel assembler syntax can be found here:
;https://software.intel.com/content/dam/develop/public/us/en/documents/325462-sdm-vol-1-2abcd-3abcd.pdf


OPTION PROLOGUE:NONE
OPTION EPILOGUE:NONE

Align 8

; --------========  Unsigned Int32 operations  ========--------

;page 605
UInt32_Add proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    add eax, ecx      ; add the value in ECX to the value in register EAX
                      ; the return value of a function in VB must be in register EAX
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)

UInt32_Add endp

;page 1857
UInt32_Sub proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    sub eax, ecx      ; subtract the value in ECX from the value in register EAX
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Sub endp

;page 1332
UInt32_Mul proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    mul ecx           ; multipliy the value in register ECX with the value in register EAX
                      ; the result as UInt64 is returned in EDX:EAX
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Mul endp

;page 891
UInt32_Div proc
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX (Dividend)
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX (Divisor )
                      ; before div we must ensure that register edx is empty 
                      ; you can do this either using mov, sub or xor
    xor edx, edx      ; xor: null all ones if there is one
    div ecx           ; divide the value in register EAX by the value in register ECX 
                      ; the quotient is in EAX the remainder (rest) in EDX
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Div endp

; just as an example, for "ByRef"
UInt32_Add_ref proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    mov eax, [eax]    ; dereference the pointer in EAX and copy den value back to EAX
    add eax, [ecx]    ; dereference the pointer in ECX und add the value from it
                      ;	to the value in register EAX
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Add_ref endp

;Thanks Creel
;https://www.youtube.com/watch?v=wXGZ7o9HNss
UInt32_Shl proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    shl eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Shl endp 

UInt32_Shld proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ebx, [esp+8]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+12] ; copy the second uint32 value from Stack to register ECX
    shld eax, ebx, cl ; shift the value in register EAX about the value in the register CL
    mov ecx, ebx
    ret 12            ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Shld endp 

UInt32_Shr proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    shr eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Shr endp 

UInt32_Shrd proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ebx, [esp+8]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+12] ; copy the second uint32 value from Stack to register ECX
    shrd eax, ebx, cl ; shift the value in register EAX about the value in the register CL
    mov ecx, ebx
    ret 12            ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Shrd endp 

UInt32_Sar proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    sar eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Sar endp 

UInt32_Rol proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    rol eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Rol endp 

UInt32_Rcl proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    rcl eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Rcl endp 

UInt32_Ror proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    ror eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Ror endp 

UInt32_Rcr proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    rcr eax, cl       ; shift the value in register EAX about the value in the register CL
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Rcr endp 

; --------======== Boolean Operations  ========-------- ;
;True-table for And
; A B  Out
; 0 0  0
; 0 1  0
; 1 0  0
; 1 1  1
UInt32_And proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    and eax, ecx      ; AND the value in register EAX with the value in the register ECX
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_And endp 

;True-table for Or
; A B  Out
; 0 0  0
; 0 1  1
; 1 0  1
; 1 1  1
UInt32_Or proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    mov ecx, [esp+8]  ; copy the second uint32 value from Stack to register ECX
    or  eax, ecx      ; OR the value in register EAX with the value in the register ECX
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_Or endp 

;True-table for Not
; A  Out
; 0  1
; 1  0
UInt32_Not proc 
    
    mov eax, [esp+4]  ; copy the first uint32 value from stack to register EAX
    not eax           ; NOT the value in register EAX, every Bit gets flipped
    ret 4             ; return to callee, remove 4 bytes from stack (->stdcall)
    
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
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
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
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
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
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
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
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
    
UInt32_NAnd endp 

;                 v1: UInt32 (4-Byte), pDec_out: pointer to Variant (4-Byte)
UInt32_ToDec proc
    
    mov   eax, [esp+ 8] ; copy the pointer to a Variant from stack to register EAX
	mov    dx,  14      ; copy the vartype for decimal it is 14 = &HE to the register dx
	mov  [eax+ 0],  dx  ; copy the value in register dx to the Variant 
	mov   edx, [esp+ 4] ; copy the UInt32-value v1 to register edx 
    mov  [eax+ 8], edx  ; copy the UInt32-Value v1 to the Variant
	ret 8
    
UInt32_ToDec endp

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
    ret 16             ; return to callee, remove 16 bytes from stack (->stdcall)
    
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
;                ByVal V1 As Currency, ByVal V2 As Currency, ByRef Dec_out As Variant; is the same as:
;                ByVal V1 As Currency, ByVal V2 As Currency, ByVal pDec_out As LongPtr
;                      V1: 8-Byte,           V2: 8-Byte,           pDec_out: 4-Byte
UInt64_Mul proc
    
	;prepare stack frame
	push ebp     ; save the register ebp to the stack
    mov ebp, esp ; save the current stack pointer to register EBP
    sub esp, 32  ; allocate 32 Bytes of new stack memory for local variables = 4 UInt64 (4*8 Bytes)
    
    ; 1. 4 Multiplications
	
    ; 1.1 Multiply the upper 32 Bit 
	
    mov  eax, [ebp+12]    ; copy the upper part of the first  uint64 value from Stack to register EAX
    mov  ecx, [ebp+20]    ; copy the upper part of the second uint64 value from Stack to register ECX
    mul  ecx              ; multipliy the value in register ECX With the value in register EAX
    mov  [esp+ 0], eax    ; zum Zwischenspeichern kopiere EAX auf den Stack
    mov  [esp+ 4], edx    ; zum Zwischenspeichern kopiere EDX auf den Stack
    
    ; 1.2 Multiply the mixed products 32 Bit
	
    ; 1.2a mixed product 1	
    mov  eax, [ebp+ 8]    ; copy the lower part of the first  uint64 value from Stack to register EAX
    mov  ecx, [ebp+20]    ; copy the upper part of the second uint64 value from Stack to register ECX
    mul  ecx              ; multipliy the value in register ECX With the value in register EAX
    mov  [esp+ 8], eax    ; zum Zwischenspeichern kopiere EAX auf den Stack
    mov  [esp+12], edx    ; zum Zwischenspeichern kopiere EDX auf den Stack
    
    ; 1.2b mixed product 2
    mov  eax, [ebp+12]    ; copy the upper part of the first  uint64 value from Stack to register EAX
    mov  ecx, [ebp+16]    ; copy the lower part of the second uint64 value from Stack to register ECX
    mul  ecx              ; multipliy the value in register ECX With the value in register EAX
    mov  [esp+16], eax    ; zum Zwischenspeichern kopiere EAX auf den Stack
    mov  [esp+20], edx    ; zum Zwischenspeichern kopiere EDX auf den Stack 
    
    ; 1.3 multiply the lower 32 Bit
	
    mov  eax, [ebp+ 8]    ; copy the lower part of the first  uint64 value from Stack To register EAX
    mov  ecx, [ebp+16]    ; copy the lower part of the second uint64 value from Stack To register ECX
    mul  ecx              ; multipliy the value in register ECX With the value in register EAX
    mov  [esp+24], eax    ; zum Zwischenspeichern kopiere EAX auf den Stack
    mov  [esp+28], edx    ; zum Zwischenspeichern kopiere EDX auf den Stack 
    
    ; 2. 6 Additionen
    ; TODO TODO TODO  
    
	; copy the results to variant-pointer 
	;
	mov  eax, [ebp+24]    ; copy the pointer to a Variant from stack to register EAX
	mov   dx, 14          ; copy the vartype for decimal it is 14 = &HE to the register dx
	mov [eax+ 0], dx      ; copy the value in register dx to the Variant 
	mov  edx, [esp+24]    ; copy from local variable back to register edx
	mov [eax+ 8], edx     ; copy the value from register edx to the Variant
	mov  edx, [esp+28]    ; copy from local variable back to register edx
	mov [eax+ 12], edx     ; copy the value from register edx to the Variant
    
	add esp, 32
	mov esp, ebp  ; copy the old stackpointer from ebp to esp back again
	pop ebp       ; restore register ebp from stack
    ret 20        ; return remove 20 bytes from call stack
    
UInt64_Mul endp


;                  v1: UInt32 (4-Byte), pDec_out: 4-Byte
;UInt32_ToDec proc
;    
;    mov   eax, [esp+ 8] ; copy the pointer to a Variant from stack to register EAX
;    mov    dx,  14      ; copy the vartype for decimal it is 14 = &HE to the register dx
;    mov  [eax+ 0],  dx  ; copy the value in register dx to the Variant 
;    mov   edx, [esp+ 4] ; copy the UInt32-value v1 to register edx 
;    mov  [eax+ 8], edx  ; copy the UInt32-Value v1 to the Variant
;    ret 8
;    
;UInt32_ToDec endp




;;               v1: UInt32 (4-Byte), pLng_out: 4Byte
UInt64_Test proc
    
	mov   ebp, esp    ; open a stackframe
	mov   eax    , [esp+4]
	mov  [eax+ 0], dx
	mov   edx    , 1
	mov  [eax+ 4], edx
	mov   edx    , 2
	mov  [eax+ 8], edx
	mov   edx    , 3
	mov  [eax+12], edx
    
	ret 8
	
UInt64_Test endp


OPTION EPILOGUE:EpilogueDef
OPTION PROLOGUE:PrologueDef

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

End DllEntry