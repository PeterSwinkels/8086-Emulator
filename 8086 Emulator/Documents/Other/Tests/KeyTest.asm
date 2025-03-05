ORG 0x0100

MOV AX, CS               ; Set the data segment to the code segment.
MOV DS, AX	         ;

MOV SS, AX	         ; Set up the stack.
LEA SP, [Stack + 0x20]   ;
MOV BP, SP               ;

CALL Initialize          ; Initialize.

Main:
   CALL CheckForKey               ; Checks for a key stroke.
                                  ;
   CMP AL, 0x00                   ; Display a key's character, if a key has been pressed.
   JE .EndIf1                     ;
      PUSH AX                     ; Save the key stroke.
                                  ; 
      MOV AH, 0x02                ; Set the cursor.
      MOV BH, 0x00                ;
      MOV DX, 0x0306              ;
      INT 0x10                    ;
                                  ;      
      MOV AH, 0x09                ; Display the key's character, if any.
      MOV BX, 0x0002              ;
      MOV CX, 0x0001              ;
      INT 0x10                    ;      
                                  ;
      POP AX                      ; Retrieve the key stroke.
   .EndIf1:                       ;
                                  ;
   CMP AX, 0x0000                 ; Check for a key stroke.
   JE .EndIf2                     ;
      PUSH AX                     ; Save the key stroke.
                                  ;
      MOV DX, AX                  ; Convert the key stroke's scancode to text.
      LEA DI, [KeyAsText]         ;
      CALL NumberToHexText        ;
                                  ;
      MOV AH, 0x02                ; Set the cursor.
      MOV BH, 0x00                ;
      MOV DX, 0x0308              ;
      INT 0x10                    ;
                                  ;
      LEA SI, [KeyAsText]         ; Display the key stroke scancode's text.
      CALL DisplayText            ;
                                  ;
      POP AX                      ; Retrieve the key stroke.
   .EndIf2:                       ;
                                  ;
   CMP AL, 0x1B                   ; Quit if Escape was pressed.
   JE Quit                        ;
JMP Main                          ;

Quit:
   MOV AX, 0x4C00   ; Quit.
   INT 0x21         ;

CheckForKey:
   MOV AH, 0x01     ; Check whether a key has been pressed.
   INT 0x16         ;
   JNZ .EndIf       ;
      XOR AX, AX    ;
      RET           ;
   .EndIf:          ;
   MOV AH, 0x00     ; Read a key's character if one has been pressed.
   INT 0x16         ;
RET

DisplayText:
   MOV AH, 0x03             ; Get the cursor position.
   MOV BH, 0x00             ;
   INT 0x10                 ;
                            ;
   MOV BL, 0x02             ; Display the text at DS:SI.
   MOV CX, 0x0001           ;
                            ;
   .DisplayCharacter:       ;
       CMP BYTE [SI], "$"   ;
       JE .Done             ;
       MOV AH, 0x09         ;
       MOV AL, [SI]         ;
       INT 0x10             ;
       INC DL               ;
       MOV AH, 0x02         ;
       INT 0x10             ;   
       INC SI               ;
   JMP .DisplayCharacter    ;
  .Done:                    ;
RET

Initialize:
   MOV AX ,0x0007       ; Set the screen mode.
   INT 0x10             ;
                        ;
   MOV AH, 0x01         ; Disable the cursor.
   MOV CX, 0x2000       ;
   INT 0x10             ;
RET

NumberToHexText:
   MOV AX, DS              ; Set the result to the default value.
   MOV ES, AX              ;
   MOV AL, " "             ;
   MOV CX, 0x0004          ;
   REP STOSB               ;
   ES                      ;
   MOV BYTE [DI], "0"      ;
                           ;
   MOV BX, 0x0010          ;
   MOV CX, DX              ; Retrieve the number to be converted to text.
   .NextDigit:             ;
      CMP CX, 0x0000       ;
      JE .Done             ;
         MOV AX, CX        ;
         XOR DX, DX        ;
         DIV BX            ;
                           ;
         CMP DX, 0x09      ;
         JA .EndIf01       ;
            ADD DX, "0"    ;
            JMP .EndIf02   ;
         .EndIf01          ;
                           ;
         CMP DX, 0x0A      ;
         JB .EndIf02       ;
            ADD DX, "7"    ;
         .EndIf02          ;
                           ;
         MOV [DI], DL      ;
         DEC DI            ;
         MOV CX, AX        ;
      JMP .NextDigit       ;
  .Done:                   ;
RET

KeyAsText DB "    0$"      ; A key's scancode represented as text.
Stack TIMES 0x20 DB 0x00   ; The stack buffer.
