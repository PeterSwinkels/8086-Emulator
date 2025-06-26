# 8086 Emulator – README

This program emulates a basic 8086 CPU, allowing users to run simple 8086 assembly programs in a virtual environment.

Key features:

Execute and debug 8086 instructions with options to inspect and modify registers and memory.

Includes a minimal BIOS and simulated hardware.

Provides limited support for MS-DOS interrupts.

If the emulator encounters unsupported I/O ports or interrupt functions, it will pause execution and prompt the user for manual input.

Includes a built-in disassembler and command reference.

Supports startup scripts for automated setup or testing.

Limitations:

No full operating system — only a handful of MS-DOS functions are simulated.

Graphics, advanced interrupt handling, and many system calls are not supported.

Designed as an experimental tool; not all instructions or edge cases are handled.

Important: To ensure correct text display, install the custom font located in: ./Documents/Other/Fonts

For more details, see Help.txt in the Documents folder.
