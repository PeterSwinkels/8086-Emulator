# 8086-Emulator
A program that immitates a 8086 CPU

This program allows you to load binary code and data into its memory and let a virtual 8086 CPU execute it. There are several options that allow the examining and modifying of the CPU's status and memory. By default there is no BIOS, OS or simulated hardware. If I/O occurs to/from a port the emulator does not support it will stop to allow the user to manually handle the I/O. If an interrupt with a vector of [0x0000:0x0000] is executed the CPU will defer to the main program which will attempt to perform the required function. If the main program cannot peform the required function it will stop the CPU to allow manual input from the user.  There is a quick reference for supported commands and a disassembler as well. The program also supports scripts on start up.

See the file Help.txt in this repository's documents folder for more information.

NOTE:

./Documents/Other/Fonts contains a font that is used by the emulator. Text might not display correctly if it isn't installed.
