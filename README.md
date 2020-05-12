# 8086-Emulator
A program that immitates a 8086 CPU

This program allows you to load binary code and data into its memory and let CPU execute it. There are several options that allow examining and modifying the CPU's status and memory. By default there is no BIOS, OS or simulated hardware. In case I/O or an interrupt occur the program will stop executing to allow you to enter input which allows whatever program is being executed to think there actually is hardware present. Also present is a quick reference for commands available and an disassembler.

Features that might be implemented in the future are:
1. An assembler, right now you can only enter raw byte code into memory or a binary assembled using a third party assembler.
2. Support for math processor (8087) instructions.
3. Your requests and suggestions.


