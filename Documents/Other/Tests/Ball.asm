ORG 0x0100

MOV AX, CS               ; Set the data segment to the code segment.
MOV DS, AX	         ;

MOV SS, AX	         ; Set up the stack.
LEA SP, [Stack + 0x10]   ;
MOV BP, SP               ;

CALL Initialize          ; Initialize.

CALL TitleScreen         ; Display the title screen.

CALL ClearScreen         ; Set up the initial game screen.
CALL DisplayStatus       ;
MOV AL, "="              ;
CALL DisplayPaddle       ;

Main:
   MOV AL, "*"                    ; Display the ball.
   CALL DisplayBall               ;
   CALL Delay                     ; Delay.
   MOV AL, " "                    ; Erase the ball.
   CALL DisplayBall               ;
   CALL VerticalBallMove          ; Move the ball.
   CALL HorizontalBallMove        ; 
                                  ;
   CMP BYTE [PaddleYD], 0x00      ; Move the paddle if necessary. 
   JE .EndIf01                    ;
      CALL MovePaddle             ;
   .EndIf01:                      ;
                                  ;
   CALL CheckForKey               ; Checks for a keystroke.
                                  ;
   CMP AL, 0x1B                   ; Quit if Escape was pressed.
   JE Quit                        ;
                                  ;
   CMP AL, "2"                    ; Stops the paddle if "2" was pressed.
   JNE .EndIf02                   ;
      MOV BYTE [PaddleYD], 0x00   ;
   .EndIf02:                      ;
                                  ;
   CMP AL, "4"                    ; Set the paddle's direction to left if "4" was pressed.
   JNE .EndIf03                   ;
      MOV  BYTE [PaddleYD], 0xFF  ;
   .EndIf03:                      ;
                                  ;
   CMP AL, "6"                    ; Set the paddle's direction to right if "6" was pressed.
   JNE .EndIf04                   ;
      MOV BYTE [PaddleYD], 0x01   ;
   .EndIf04:                      ;
                                  ;
   CMP BYTE [Lives], 0x00         ; Quit if the player is out of lives.
   JE Quit                        ;
JMP Main                          ;

Quit:
   MOV AX, 0x4C00   ; Quit.
   INT 0x21         ;

CheckForKey:
   MOV AH, 0x01     ; Check whether a key has been pressed.
   INT 0x16         ;
   JNZ .EndIf       ;
      XOR AL, AL    ;
      RET           ;
   .EndIf:          ;
   MOV AH, 0x00     ; Read a key's character if one has been pressed.
   INT 0x16         ;
RET

CheckPaddleHit:
   MOV AH, [PaddleY]            ; Check if ball is the left of the paddle.
   CMP [BallY], AH              ;
   JB .NoPaddleHit              ;
                                ;
   ADD AH, 0x08                 ; Check if ball is the right of the paddle.
   CMP [BallY], AH              ;
   JA .NoPaddleHit              ;
                                ;
   CMP BYTE [PaddleYD], 0x01    ; Check whether the paddle is moving to the right.
   JNE .EndIf01                 ;
      MOV BYTE [BallYD], 0x01   ;
   .EndIf01:                    ;
                                ;
   CMP BYTE [PaddleYD], 0xFF    ; Check whether the paddle is moving to the left.
   JNE .EndIf02                 ;
      MOV BYTE [BallYD], 0xFF   ;
   .EndIf02:                    ;
                                ;
   JMP .PaddleHit               ;
                                ;
   .NoPaddleHit:                ;
      DEC BYTE [Lives]          ; The paddle missed the ball.
      CALL DisplayStatus        ;
      RET                       ;
                                ;
   .PaddleHit:                  ;
      INC WORD [Score]          ; The paddle hit the ball.
      CALL DisplayStatus        ;
RET

ClearScreen:
   MOV AX, 0xB000    ; Clear the screen.
   MOV ES, AX        ;
   XOR DI, DI        ;
   XOR AX, AX        ;
   MOV CX, 0x07D0    ;
   REP STOSW         ;
RET

Delay:
   MOV CX, 0x0001          ; Delay.
   .DelayStart:            ;
       MOV DX, 0x6000      ;
       .SubDelayStart:     ;
          DEC DX           ;
          CMP DX, 0x0000   ;
       JA .SubDelayStart   ;
       DEC CX              ;   
       CMP CX, 0x0000      ;
   JA .DelayStart          ;
RET

DisplayBall:
   MOV AH, 0x02           ; Set the cursor position.
   MOV BH, 0x00           ;
   MOV DH, [BallX]        ;
   MOV DL, [BallY]        ;
   INT 0x10               ; Displays or erases the ball.
   MOV AH, 0x09           ;
   MOV BL, 0x02           ;
   MOV CX, 0x0001         ;
   INT 0x10               ;
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

DisplayPaddle:
   MOV AH, 0x02          ; Set the cursor position.
   MOV BH, 0x00          ;
   MOV DH, 0x17          ;
   MOV DL, [PaddleY]     ;
   INT 0x10              ;
   MOV AH, 0x09          ; Displays or erases the paddle.
   MOV BL, 0x02          ;
   MOV CX, 0x0008        ;
   INT 0x10              ;
RET

DisplayStatus:
   MOV AH, 0x02          ; Erase the status bar.
   MOV BH, 0x00          ;
   MOV DX, 0x1800        ;
   INT 0x10              ;
                         ;
   MOV AX, 0x0920        ; 
   MOV BL, 0x02          ;
   MOV CX, 0x4F          ;
   INT 0x10              ;
                         ;
   MOV AH, 0x02          ; Display the score.
   MOV BH, 0x00          ;
   MOV DX, 0x1803        ;
   INT 0x10              ;
                         ;
   MOV DX, [Score]       ;
   LEA DI, [ScoreAsText] ;
   CALL NumberToText     ;
   LEA SI, [ScoreAsText] ;
   CALL DisplayText      ;
                         ;
   MOV AH, 0x02          ; Display the number of lives. 
   MOV BH, 0x00          ;
   MOV DX, 0x184B        ;
   INT 0x10              ;
                         ;
   MOV AX, 0x0903        ;
   MOV BL, 0x02          ;
   XOR CH, CH            ;
   MOV CL, [Lives]       ;
   INT 0x10              ;
RET

HorizontalBallMove:
   MOV AL, [BallY]              ; Change the ball's direction if it has reached the left side of the field.
                                ;
   CMP AL, 0x00                 ;
   JA .EndIf01                  ;
      MOV BYTE [BallYD], 0x01   ;
   .EndIf01:                    ;
                                ;
   CMP AL, 0x4F                 ; Change the ball's direction if it has reached the right side of the field.
   JB .EndIf02                  ;
      MOV BYTE [BallYD], 0xFF   ;
   .EndIf02:                    ;
                                ;
   ADD AL, [BallYD]             ; Move the ball horizontally.
                                ;
   MOV [BallY], AL              ;
RET

Initialize:
   MOV AX ,0x0007       ; Set the screen mode.
   INT 0x10             ;
                        ;
   MOV AH, 0x01         ; Disable the cursor.
   MOV CX, 0x2000       ;
   INT 0x10             ;
RET

MovePaddle:
   MOV AL, " "                       ; Erase the paddle.
   CALL DisplayPaddle                ;
                                     ;
   MOV AL, [PaddleY]                 ;
                                     ;
   CMP BYTE [PaddleYD], 0xFF         ; Stop the paddle from moving if it has reached the left side of the field.
   JNE .EndIf01                      ;
      CMP AL, 0x00                   ;
      JNE .EndIf01                   ;
         MOV BYTE [PaddleYD], 0x00   ;
   .EndIf01:                         ;
                                     ;
   CMP BYTE [PaddleYD], 0x01         ; Stop the paddle from moving if it has reached the right side of the field.
   JNE .EndIf02                      ;
      CMP AL, 0x48                   ;
      JNE .EndIf02                   ;
         MOV BYTE [PaddleYD], 0x00   ;
   .EndIf02:                         ;
                                     ;
   ADD AL, [PaddleYD]                ; Move the paddle horizontally.
                                     ;
   MOV [PaddleY], AL                 ;
                                     ;
   MOV AL, "="                       ; Display the paddle.
   CALL DisplayPaddle                ;
RET

NumberToText:
   MOV AX, DS           ; Set the result to the default value.
   MOV ES, AX           ;
   MOV AL, " "          ;
   MOV CX, 0x0004       ;
   REP STOSB            ;
   ES                   ;
   MOV BYTE [DI], "0"   ;
                        ;
   MOV BX, 0x000A       ;
   MOV CX, DX           ; Retrieve the number to be converted to text.
   .NextDigit:          ;
      CMP CX, 0x0000    ;
      JE .Done          ;
         MOV AX, CX     ;
         XOR DX, DX     ;
         DIV BX         ;
         ADD DX, "0"    ;
         MOV [DI], DL   ;
         DEC DI         ;
         MOV CX, AX     ;
      JMP .NextDigit    ;
  .Done:                ;
RET

TitleScreen:
   CALL ClearScreen        ; Display the title screen.
                           ;
   MOV AH, 0x02            ;
   MOV DX, 0x0203          ;
   INT 0x10                ;
   LEA SI, [Title]         ;
   CALL DisplayText        ;
                           ;
   MOV AH, 0x02            ;
   MOV DX, 0x0406          ;
   INT 0x10                ;
   LEA SI, [Help]          ;
   CALL DisplayText        ;
                           ;
   MOV AH, 0x02            ;
   MOV DX, 0x0606          ;
   INT 0x10                ;
   LEA SI, [Continue]      ;
   CALL DisplayText        ;
                           ;
   WaitForEnter:           ;
      CALL CheckForKey     ;
      CMP AL, 0x0D         ;
   JNE WaitForEnter        ; 
RET

VerticalBallMove:
   MOV AL, [BallX]              ; Change the ball's direction if it has reached the top of the field.
                                ;
   CMP AL, 0x00                 ;
   JA .EndIf01                  ;
      MOV BYTE [BallXD], 0x01   ;
   .EndIf01:                    ;
                                ;
   CMP AL, 0x17                 ; Change the ball's direction if it has reached the bottom of the field.
   JB .EndIf02                  ;
      PUSH AX                   ;
      CALL CheckPaddleHit       ;
      POP AX                    ;
      MOV BYTE [BallXD], 0xFF   ;
   .EndIf02:                    ;
                                ;
   ADD AL, [BallXD]             ; Move the ball vertically.
                                ;
   MOV [BallX], AL              ;
RET

BallX DB 0x00                                             ; The ball's vertical position.
BallY DB 0x00                                             ; The ball's horizontal position.
BallXD DB 0x00                                            ; The ball's vertical direction.
BallYD DB 0x00                                            ; The ball's horizontal direction.
Lives DB 0x03                                             ; The number of lives.
PaddleY DB 0x27                                           ; The paddle's position.
PaddleYD DB 0x00                                          ; The paddle's direction.
Score DW 0x0000                                           ; The score.
ScoreAsText DB "    0$"                                   ; The score represented as text.
Title DB "Ball v1.01 - by: Peter Swinkels, ***2025***$"   ; The title text.
Help DB "4 = Left  6 = Right  2 = Stop  Esc = Quit$"      ; The help text.
Continue DB "Press the Enter to key continue...$"         ; The to continue text.
Stack TIMES 0x10 DB 0x00                                  ; The stack buffer.
