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
Int32_UAdd proc v1: dword, v2: dword
    
    mov eax, v1       ; copy the value v1 from stack to register EAX
    mov ecx, v2       ; copy the value v2 from Stack to register ECX
	add eax, ecx      ; add the value in ECX to the value in register EAX
	                  ; the return value of a function in VB must be in register EAX
    ret 8             ; return to callee, remove 8 bytes from stack (->stdcall)
                      ; normally ret is enough, because the size of the stack 
					  ; is determined automatically
Int32_UAdd endp

;page 1857
Int32_USubtract proc v1: dword, v2: dword
    
    mov eax, v1       ; copy the value v1 from stack to register EAX
    mov ecx, v2       ; copy the value v2 from Stack to register ECX
    sub eax, ecx      ; subtract the value in ECX from the value in register EAX
    ret 8             ; return to callee, remove 8 bytes from stack
    
Int32_USubtract endp

;page 1332
Int32_UMultiply proc v1: dword, v2: dword
    
    mov eax, v1       ; copy the value v1 from stack to register EAX
    mov ecx, v2       ; copy the value v2 from Stack to register ECX
    mul ecx           ; multipliy the value in register ECX with the value in register EAX
    ret 8             ; return
    
Int32_UMultiply endp

;page 891
Int32_UDivide proc v1: dword, v2: dword
    
    mov eax, v1 ; copy the value v1 from stack to register EAX (Dividend)
    mov ecx, v2 ; copy the value v2 from Stack to register ECX (Divisor )
                ; before div we must ensure that register edx is empty 
	            ; you can do this either using mov, sub or xor
	;mov edx, 0   ; mov: copy 0 to the register edx
	;sub edx, edx ; sub: subtract the value itself -> Tipp von Ripler
	xor edx, edx  ; xor: null all ones if there is one
    div ecx     ; divide the value in register EAX by the value in register ECX 
	            ; the quotient is in EAX the remainder (rest) in EDX
    ret 8       ; return
    
Int32_UDivide endp

;https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2013/kk8w4t5t(v=vs.120)
Int32_UToStr proc v1: dword, pStr: dword 
	
	invoke crt__ultow, v1, pStr, 10 ; rradix=10
	ret
	
Int32_UToStr endp

; just as an example, for "ByRef"
;Int32_UAdd_ref proc pV1: dword, pV2: dword
;    
;    mov eax, pV1      ; kopiere den obersten Wert vom Stack in das EAX-Register
;    mov ecx, pV2      ; kopiere den nächsten Wert vom Stack in das ECX-Register
;    mov eax, [eax]    ; Dereferenziere den Zeiger in EAX und kopieren den Wert in EAX
;    add eax, [ecx]    ; Dereferenzieren den Zeiger in ECX und addiere den Wert daraus
;                      ;	auf den Wert im Register EAX
;    ret 8             ; zurückkehren
;    
;Int32_UAdd_ref endp

End DllEntry