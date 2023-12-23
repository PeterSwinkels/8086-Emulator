'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System

'This module contains the default interrupt handler.
Public Module InterruptHandlerModule
   Private Const MS_DOS As Integer = &HFF00%           'Indicates that the operating is MS-DOS.
   Private Const MS_DOS_VERSION As Integer = &H1606%   'Defines the emulated MS-DOS version as 6.22.

   'This procedure handles the specified interrupt and returns success or failure.
   Public Function HandleInterrupt(Number As Integer, AH As Integer) As Boolean
      Try
         Dim Success As Boolean = False

         Select Case Number
            Case &H21%
               Select Case AH
                  Case &H30%
                     CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=MS_DOS_VERSION)
                     CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=MS_DOS)
                     CPU.ExecuteOpcode(CPU8086Class.OpcodesE.IRET)
                     Success = True
               End Select
         End Select

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function
End Module
