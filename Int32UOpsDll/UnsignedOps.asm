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


;The manual to Intel assembler syntax can be found here:
;https://software.intel.com/content/dam/develop/public/us/en/documents/325462-sdm-vol-1-2abcd-3abcd.pdf

;page 605

OPTION PROLOGUE:NONE
OPTION EPILOGUE:NONE

Align 8

; --------========  Unsigned Int32 operations  ========--------

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

; --------========  Unsigned Int64 operations  ========--------

UInt64_Add proc
	
	mov eax, [esp+4]   ; copy the lower part of the first uint64 value from stack to register EAX
	mov edx, [esp+8]   ; copy the upper part of the first uint64 value from Stack to register EDX
	add eax, [esp+12]  ; add the lower part of the second uint64 value from stack to the value in register EAX 
	adc edx, [esp+16]  ; add the upper part of the second uint64 value from stack to the value in register EDX by using the carry flag
	ret 16             ; return to callee, remove 16 bytes from stack (->stdcall)
	
UInt64_Add endp

UInt64_Sub proc
	
	mov eax, [esp+4]   ; copy the lower part of the first uint64 value from stack to register EAX
	mov edx, [esp+8]   ; copy the upper part of the first uint64 value from Stack to register EDX
	sub eax, [esp+12]  ; subtract the lower part of the second uint64 value from stack to the value in register EAX 
	sbb edx, [esp+16]  ; subtract the upper part of the second uint64 value from stack to the value in register EDX by using the borrow flag
	ret 16
	
UInt64_Sub endp

OPTION EPILOGUE:EpilogueDef
OPTION PROLOGUE:PrologueDef

;https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2013/kk8w4t5t(v=vs.120)
UInt32_ToStr proc v1: dword, pStr: dword 
    
    invoke crt__ultow, v1, pStr, 10 ; rradix=10
    ret 8
    
UInt32_ToStr endp

UInt64_ToStr proc v1: qword, pStr: dword 
    
    invoke crt__ui64tow, v1, pStr, 10 ; rradix=10
    ret 12
    
UInt64_ToStr endp

End DllEntry