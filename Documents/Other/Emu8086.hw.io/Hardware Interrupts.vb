'This module's imports and settings.
'
'Note:
'A call needs to be made to the CheckForHardwareInterrupt procedure to check for hardware interrupts by the CPU class.
'
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.VisualBasic
Imports System
Imports System.Environment
Imports System.IO

'This module handles hardware interrupts triggered by external programs.
Public Module HardwareInterruptsModule
   Private ReadOnly INT_FILE As String = If(File.Exists("hw.txt"), File.ReadAllText("hw.txt").Trim({" "c, ControlChars.Cr, ControlChars.Lf, ControlChars.Tab}), "C:\Emu8086.hw")   'Contains the hardware interrupt file.

   'This procedure checks for an hardware interrupt and triggers it if necessary.
   Public Sub CheckForHardwareInterrupt()
      Try
         Dim Interrupts() As Byte = {}

         If CBool(CPU.Registers(CPU8086Class.FlagRegistersE.IF)) Then
            Using FileO As New FileStream(INT_FILE, FileMode.OpenOrCreate, FileAccess.Read, FileShare.ReadWrite)
               ReDim Interrupts(0 To CInt(FileO.Length - 1))

               FileO.Read(Interrupts, 0, Interrupts.Length)
            End Using

            For Number As Integer = &H0% To Interrupts.Length - &H1%
               If Not Interrupts(Number) = &H0% Then
                  SyncLock Synchronizer
                     CPUEvent.Append($"HW INT {Number:X}{NewLine}")
                  End SyncLock

                  CPU.ExecuteInterrupt(CPU8086Class.OpcodesE.INT, Number)
               End If
            Next Number

            ReDim Interrupts(&H0% To &HFF%)

            Using FileO As New FileStream(INT_FILE, FileMode.Open, FileAccess.Write, FileShare.ReadWrite)
               FileO.Write(Interrupts, 0, Interrupts.Length)
            End Using
         End If
      Catch ExceptionO As Exception
         CPUEvent.Append($"{ExceptionO.Message}{NewLine}")
      End Try
   End Sub

End Module
