ORG 0x0100

MOV AX, CS              ; Ensure data segment and code segment are the same.
MOV DS, AX	        ;

MOV SS, AX		; Set up the stack.
LEA SP, [Stack + 0x08]  ;
MOV BP, SP              ;

MOV AX, 0x0007          ; Use Hercules/MDA mode.
INT 0x10                ;

MOV AH, 0x01            ; Large cursor.
MOV CX, 0x000D          ;
INT 0x10                ;

MOV BX, 0x0002          ; Video page and character attribute.
MOV DX, 0x0407          ; Initial cursor position.

LEA SI, [Text]          ; Location of the text to be displayed.

WriteText:
 MOV AH, 0x02           ; Position the cursor.
 INT 0x10               ;
 MOV AL, [SI]           ; Retrieve a character.
 CMP BYTE AL, "$"       ; Quit at end of text.
 JE Quit                ;
 MOV CX, 0x0001         ;
 MOV AH, 0x09           ;
 INT 0x10               ;
 INC DL                 ; Next column.
 INC SI                 ; Next character.
JMP SHORT WriteText     ;

Quit:
MOV AX, 0x4C00          ; Quit.
INT 0x21                ;

Stack TIMES 0x08 DB 0x00   ; The stack buffer.

Text DB "Hello World - by: Peter Swinkels, ***2025***$"   ; The text to be displayed.
