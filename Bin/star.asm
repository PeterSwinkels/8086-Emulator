ORG 0x0100

MOV AX, CS               ; Set the data segment to the code segment.
MOV DS, AX	         ;

MOV SS, AX	         ; Set up the stack.
LEA SP, [Stack + 0x10]   ;
MOV BP, SP   


                MOV     AX,0B000H
                MOV     ES,AX           ;manda es aponta segment video (CGA) :(

				CALL Initialize          ; Initialize.
				call CLRSCR
				                    
                MOV     SI, 0           ;linhas impressas
                mov     ah, 3dh         ;Open the file
                mov     al, 0           ;Open for reading
                lea     dx, Filename    ;Presume DS points at filename
                int     21h             ; segment.
                jc      BadOpen
                mov     [FHndl], ax       ;Save file handle

LP:             mov     ah,3fh          ;Read data from the file
                lea     dx, Buffer      ;Address of data buffer
                mov     cx, 1           ;Read one byte
                mov     bx, [FHndl]       ;Get file handle value
                int     21h
                jc      ReadError
                cmp     ax, cx          ;EOF reached?
                jne     EOF
                mov     al, [Buffer]      ;Get character read
                
                cmp     al,10
                je      conta_linhas
                jmp     imprime
                
conta_linhas:
                inc     si                
                cmp     si,13           ;imprimiu 13 linhas
                je      volta_cursor_0_0
                jmp     imprime
                
volta_cursor_0_0:      
                PUSH AX
                PUSH BX
                PUSH CX
                MOV AH,    02H
                MOV BH, 0 ;    video page number (0-based)
                MOV DH, 0 ;   row (0-based)
                MOV DL,0  ;    column (0-based)
                INT 10H
                POP CX
                POP BX
                POP AX
                MOV SI,0
                               
                mov ah, 1
                int 21h
                
                CALL CLRSCR   ;APAGA TELA

imprime:                

                call write                    ;Print it
                jmp     LP              ;Read next byte

EOF:            mov     bx, FHndl
                mov     ah, 3eh         ;Close file
                int     21h
                ;jc      CloseError
                
    lea dx, pkey
    mov ah, 9
    int 21h        ; output string at ds:dx
    
    ; wait for any key....    
    mov ah, 1
    int 21h
    
    mov ax, 4c00h ; exit to operating system.
    int 21h     
    
write:
		    push ax 
		    push bx 
		    push cx 
		    push dx
		    mov ah, 2
		    mov dl, [Buffer] ;char to be printed
		    int 21h
	        pop ax 
	        pop bx 
	        pop cx 
	        pop dx
		    ret

BadOpen:
	mov ah,9
	mov dx, ERROR_OPEN
	int 21h
    mov ax, 4c00h ; exit to operating system.
    int 21h     
	
ReadError:
	mov ah,9
	mov dx, ERROR_READ
	int 21h
    mov ax, 4c00h ; exit to operating system.
    int 21h     

Initialize:
   MOV AX ,0x0007       ; Set the screen mode.
   INT 0x10             ;
                        ;
   MOV AH, 0x01         ; Disable the cursor.
   MOV CX, 0x2000       ;
   INT 0x10             ;
RET
                
CLRSCR:
        PUSH AX
        MOV DI,0
APAGANDO:
        MOV AL," "
        MOV ES:[DI],AL
        ADD DI,2
        CMP DI, 4000   ; 25X80 + ATRIBUTO DE COR PARA CADA BYTE
        JG SAI
        JMP APAGANDO
SAI:       
        POP AX
        RET

    Filename db 'star.txt'
	FHndl dw 0
	Buffer db 80h dup(0)    ; add your data here!
    pkey db "press any key...$"
	ERROR_OPEN db "FILE NOT FOUND$"
	ERROR_READ db "READ ERROR$"
Stack TIMES 0x10 DB 0x00                                  ; The stack buffer.
