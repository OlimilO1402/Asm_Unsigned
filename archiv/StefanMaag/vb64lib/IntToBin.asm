; --------------------------------------------------------------------------------------
; Function:		Number64String
;
; Paraneters:	[IN] lpString (LPSTR) = Pointer auf den String
; 		[IN] Number (LARGE_INTEGER) = 64 Bit Ganzzahl

; Rückgabewert: Anzahl Zeichen die benötigt werden oder kopiert wurden exklusive dem NULL-CHAR
; --------------------------------------------------------------------------------------
; uses ebx esi edi

Number64String	PROC		lpString:LPSTR, Number:LARGE_INTEGER
	Local szTemp[50]:Byte
	Local TempNumber:LARGE_INTEGER
	
	xor eax, eax
	mov ecx, 50
	lea edi, szTemp[0]
	rep stosb
	
	m2m TempNumber.LowPart,Number.LowPart
	m2m TempNumber.HighPart,Number.HighPart
		
	lea edi,szTemp[48]
	mov ebx,10
	
	xor ecx,ecx
	.While True
		mov eax, TempNumber.HighPart
		xor edx,edx
		.If eax
			div ebx
			mov TempNumber.HighPart, eax
		.endif
		mov eax, TempNumber.LowPart
		div ebx
		mov TempNumber.LowPart, eax
		add edx,30h	
		mov Byte ptr [edi],dl
		Sub edi,1
		add ecx,1
		.break .If eax == 0
	.endw
	
	push ecx
	.If ecx && lpString
		lea esi, [edi+1] 
		mov edi,lpString
		inc ecx
		rep movsb
	.endif
	pop eax
	ret
Number64String endp

{------------------------------------------------------------------------------}
{ IntToBin  (integer)                                                          }
{ Wandelt value in einen String mit der Länge len um, der die unteren len Bits }
{ von value in binärer Darstellung zeigt.                                      }
{ Wenn len<=0 oder len>32 ist, entscheidet die Position des höchsten 1-Bits    }
{ in value über die Länge des Strings, wobei aber, wenn value leer ist,        }
{ mindestens ein Zeichen ausgegeben wird                                       }
{------------------------------------------------------------------------------}
FUNCTION IntToBin(value,len:integer):string;
const bits:array[0..63] of char=
         '0000000100100011010001010110011110001001101010111100110111101111';
asm
               // EAX Wert
               // EDX Länge
               // ECX Zeiger Result
               push     edi
               mov      edi,ecx           // Zeiger auf Result
               or       edx,edx
               je       @DefineLen        // len = 0, benötigte Länge holen
               cmp      edx,32
               jbe      @Start            // 0 < len <= 32
@DefineLen:    bsr      edx,eax           // Index höchstes Bit in Value
               jnz      @AdjustLen
               xor      edx,edx
@AdjustLen:    add      edx,1
@Start:        push     eax               // value
               mov      eax,edi           // Zeiger Result
               call     System.@LStrSetLength
               mov      edi,[edi]         // Zeiger String
               mov      ecx,[edi-4]       // Länge
               pop      eax               // value
               sub      ecx,4
               jc       @EndDW            // weniger als 4 bits
               // Jeweils 4 Bits umwandeln
@LoopDW:       mov      edx,eax
               and      edx,$F
               mov      edx,DWORD [bits+edx*4]
               mov      [edi+ecx],edx
               shr      eax,4
               je       @FillZero
               sub      ecx,4
               jnc      @LoopDW           // nochmal 4 Bits
@EndDW:        // Restline 1..3 Bytes
               and      eax,$F
               mov      eax,DWORD [bits+eax*4]
               cmp      cl,$FE
               ja       @3Bytes
               je       @2Bytes
               jnp      @End
               rol      eax,8
               mov      [edi],al
               jmp      @End
@2Bytes:       shr      eax,16
               mov      [edi],ax
               jmp      @End
@3Bytes:       mov      [edi],ah
               shr      eax,16
               mov      [edi+1],ax
               jmp      @End
               // Mit 0en auffüllen
               // ECX = Anzahl
@FillZero:     mov      eax,'0000'
               mov      edx,ecx
               shr      ecx,2
               rep      stosd
               mov      ecx,edx
               and      ecx,3
               rep      stosb
@End:          pop      edi
end; 

{------------------------------------------------------------------------------}
{ IntToBin  (int64)                                                            }
{ Wandelt value in einen String mit der Länge len um, der die unteren len Bits }
{ von value in binärer Darstellung zeigt.                                      }
{ Wenn len<=0 oder len>64 ist, entscheidet die Position des höchsten 1-Bits    }
{ in value über die Länge des Strings, wobei aber, wenn value leer ist,        }
{ mindestens ein Zeichen ausgegeben wird                                       }
{------------------------------------------------------------------------------}
FUNCTION IntToBin(value:int64; len:integer):string;
const bits:array[0..63] of char=
         '0000000100100011010001010110011110001001101010111100110111101111';
asm
               // EAX Länge
               // EDX Zeiger Result
               // [EBP+8]  = LoValue
               // [EBP+12] = HiValue
               push     ebx
               push     edi
               mov      edi,edx        // Zeiger Result
               mov      ebx,[ebp+8]    // LoValue
               mov      ecx,[ebp+12]   // HiValue
               // len prüfen
               mov      edx,eax
               or       edx,edx
               je       @DefineLen     // len = 0, benötigte Länge holen
               cmp      edx,64
               jbe      @Start
@DefineLen:    mov      eax,33
               bsr      edx,ecx        // EAX=Höchstes Bit in HiValue
               jnz      @AdjustLen     // HiValue nicht leer
               mov      eax,1
               bsr      edx,ebx        // EAX=Höchstes Bit in LoValue
               jnz      @AdjustLen     // LoValue nicht leer
               xor      edx,edx
@AdjustLen:    add      edx,eax        // Länge=BitPos+1
@Start:        mov      eax,edi        // Zeiger Result
               call     System.@LStrSetLength
               mov      edi,[edi]         // Zeiger String
               mov      ecx,[edi-4]       // Länge
               mov      edx,[ebp+12]      // HiValue
               // Jeweils 4 Bits umwandeln
@LoopDW:       sub      ecx,4
               jc       @EndDW            // weniger als 4 bits
               mov      eax,ebx           // LoValue
               and      eax,$F
               mov      eax,DWORD [bits+eax*4]
               mov      [edi+ecx],eax
               shrd     ebx,edx,4         // Value 4 Bits nach unten
               shr      edx,4
               jne      @LoopDW           // HiValue nicht leer
               or       ebx,ebx
               jne      @LoopDW           // LoValue nicht leer
               // Mit 0en auffüllen
@FillZero:     mov      eax,'0000'
               mov      edx,ecx
               shr      ecx,2
               rep      stosd
               mov      ecx,edx
               and      ecx,3
               rep      stosb
               jmp      @End
@EndDW:        // Restline 1..3 Bytes
               and      ebx,$F
               mov      eax,DWORD [bits+ebx*4]
               cmp      cl,$FE
               ja       @3Bytes
               je       @2Bytes
               jnp      @End
               rol      eax,8
               mov      [edi],al
               jmp      @End
@2Bytes:       shr      eax,16
               mov      [edi],ax
               jmp      @End
@3Bytes:       mov      [edi],ah
               shr      eax,16
               mov      [edi+1],ax
@End:          pop      edi
               pop      ebx
end;