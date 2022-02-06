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
    
    mov eax, [esp+4]  ; copy the value v1 from stack to register EAX
    mov ecx, [esp+8]  ; copy the value v2 from Stack to register ECX
    add eax, ecx      ; add the value in ECX to the value in register EAX
                      ; the return value of a function in VB must be in register EAX
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
                      ; normally ret is enough, because the size of the stack 
                      ; is determined automatically
UInt32_Add endp


;page 1857
UInt32_Sub proc 
    
    mov eax, [esp+4]       ; copy the value v1 from stack to register EAX
    mov ecx, [esp+8]       ; copy the value v2 from Stack to register ECX
    sub eax, ecx      ; subtract the value in ECX from the value in register EAX
    ret 8             ; return to callee, remove 8 bytes from stack
    
UInt32_Sub endp

;page 1332
UInt32_Mul proc 
    
    mov eax, [esp+4]       ; copy the value v1 from stack to register EAX
    mov ecx, [esp+8]       ; copy the value v2 from Stack to register ECX
    mul ecx           ; multipliy the value in register ECX with the value in register EAX
    ret 8             ; return
    
UInt32_Mul endp

;page 891
UInt32_Div proc
    
    mov eax, [esp+4] ; copy the value v1 from stack to register EAX (Dividend)
    mov ecx, [esp+8] ; copy the value v2 from Stack to register ECX (Divisor )
                     ; before div we must ensure that register edx is empty 
                     ; you can do this either using mov, sub or xor
    xor edx, edx     ; xor: null all ones if there is one
    div ecx          ; divide the value in register EAX by the value in register ECX 
                     ; the quotient is in EAX the remainder (rest) in EDX
    ret 8            ; return
    
UInt32_Div endp

; just as an example, for "ByRef"
UInt32_Add_ref proc 
    
    mov eax, [esp+4]  ; kopiere den obersten Wert vom Stack in das EAX-Register
    mov ecx, [esp+8]  ; kopiere den nächsten Wert vom Stack in das ECX-Register
    mov eax, [eax]    ; Dereferenziere den Zeiger in EAX und kopieren den Wert in EAX
    add eax, [ecx]    ; Dereferenzieren den Zeiger in ECX und addiere den Wert daraus
                      ;	auf den Wert im Register EAX
    ret 8             ; zurückkehren
    
UInt32_Add_ref endp

; --------========  Unsigned Int64 operations  ========--------

UInt64_Add proc
	
	mov eax, [esp+4]
	mov edx, [esp+8] 
	add eax, [esp+12]
	adc edx, [esp+16]
	ret 16
	
UInt64_Add endp

UInt64_Sub proc
	
	mov eax, [esp+4]
	mov edx, [esp+8] 
	sub edx, [esp+16]
	sub eax, [esp+12]
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