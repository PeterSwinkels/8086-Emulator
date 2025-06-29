org 100h 

section .data
    buffer      db 128 dup('$')         ; buffer voor lezen (met '$' voor DOS printen)
    bytesRead   dw 0

section .text
    global _start

_start:
    ; --- Lees bestand ---
    mov  ah, 3Fh
    mov  bx, 0                         ; 0 = in, 1 = aux
    mov  cx, 14
    mov  dx, buffer
    int  21h
    jc   read_error
    mov  [bytesRead], ax

    ; Voeg '$' toe aan einde voor DOS print (als ax < 128)
    mov  si, ax
    mov  byte [buffer + si], '$'

    ; --- Toon inhoud op scherm (functie 09h) ---
    mov  ah, 09h
    mov  dx, buffer
    int  21h

    jmp  done

open_error:
    mov  ah, 09h
    mov  dx, err_open
    int  21h
    jmp  done

read_error:
    mov  ah, 09h
    mov  dx, err_read
    int  21h

done:
    mov  ax, 4C00h
    int  21h

section .data
    err_open db "Fout bij openen van bestand.$"
    err_read db "Fout bij lezen van bestand.$"
