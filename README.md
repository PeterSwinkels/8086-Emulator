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
Software that has the best chance of executing properly to any degree are simple text mode programs from the 1980's and early 1990's. (If they work in e.g. DOSBox with an emulated 8086 CPU and Hercules machine they might execute.)

Graphics, advanced interrupt handling, and many system calls are not supported.

Designed as an experimental tool; not all instructions or edge cases are handled.

For more details, see Help.txt in the Documents folder.

Important:

- To ensure correct text display, install the custom font located in: ./Documents/Other/Fonts
- 
- Copy ./Documents/Bin/SDL2.DLL to the folder containing the emulator's executable.
- 
- Use the Nuget Package Manager to install SDL2-CS.dll in the solution.


Also, a useful technical reference for old (1980's-1990's) x86 based hardware and software can be found at:

https://helppc.netcore2k.net/

A utility to deal with traces containing duplicate events generated while both executing and tracing. (A glitch which appears to be caused by asynchronous code.):

https://github.com/PeterSwinkels/8086-Emulator-Trace-Cleanup
