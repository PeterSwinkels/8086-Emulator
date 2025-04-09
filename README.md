# 8086-Emulator
A program that emulates a 8086 CPU.

This program allows users to let a virtual 8086 execute programs. There are several options for examining and modifying the CPU's status and memory. By default there only a very rudimentary BIOS and simulated hardware. There is no OS except for some simulated MS-DOS functions. If I/O occurs to/from a port the emulator does not support, it will stop to allow the user to manually handle the I/O. If the emulator cannot execute an interrupt function it will stop the CPU to allow manual input from the user.  There is a quick reference for supported commands and a disassembler as well. The program also supports scripts on start up.

See the file Help.txt in this repository's documents folder for more information.

NOTE:

./Documents/Other/Fonts contains a font that is used by the emulator. Text might not display correctly if it isn't installed.
