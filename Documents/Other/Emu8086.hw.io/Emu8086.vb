'This module's imports and settings.
'
'Note:
'A call needs to be added to the CheckForHardwareInterrupt procedure in the CPU class.
'A call needs to be added to the DisplayEmu8086Information procedure in this program's start up procedure.

Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.VisualBasic
Imports System
Imports System.Environment
Imports System.IO

'This module handles hardware interrupts and port I/O from external programs Emu8086 style.
Public Module Emu8086Module
   Private ReadOnly INT_FILE As String = If(File.Exists("hw.txt"), File.ReadAllText("hw.txt").Trim({" "c, ControlChars.Cr, ControlChars.Lf, ControlChars.Tab}), "C:\Emu8086.hw")   'Contains the hardware interrupt file.
   Private ReadOnly IO_FILE As String = If(File.Exists("io.txt"), File.ReadAllText("io.txt").Trim({" "c, ControlChars.Cr, ControlChars.Lf, ControlChars.Tab}), "C:\Emu8086.io")   'Contains the I/O file.

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

   'This procedure displays information about the files used for I/O and interrupts.
   Public Sub DisplayEmu8086Information()
      Try
         SyncLock Synchronizer
            CPUEvent.Append($"I/O File: ""{IO_FILE}""{NewLine}")
            CPUEvent.Append($"HW INT File: ""{INT_FILE}""{NewLine}")
         End SyncLock
      Catch ExceptionO As Exception
         CPUEvent.Append($"{ExceptionO.Message}{NewLine}")
      End Try
   End Sub

   'This procedure attempts to read from the specified I/O port and returns the result.
   Public Function ReadIOPort(Port As Integer) As Integer?
      Try
         Dim Values() As Byte = {}

         Using FileO As New FileStream(IO_FILE, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            ReDim Values(0 To CInt(FileO.Length - 1))

            FileO.Read(Values, 0, Values.Length)
         End Using

         SyncLock Synchronizer
            CPUEvent.Append($"I/O: IN {Port:X}, {Values(Port):X}{NewLine}")
         End SyncLock

         Return Values(Port)
      Catch ExceptionO As Exception
         SyncLock Synchronizer
            CPUEvent.Append($"{ExceptionO.Message}{NewLine}")
         End SyncLock
      End Try

      Return Nothing
   End Function

   'This procedure attempts to write the specified I/O port and returns whether or not it succeeded.
   Public Function WriteIOPort(Port As Integer, Value As Integer) As Boolean
      Try
         Dim Values() As Byte = {}

         Using FileO As New FileStream(IO_FILE, FileMode.OpenOrCreate, FileAccess.Read, FileShare.ReadWrite)
            ReDim Values(0 To CInt(FileO.Length - 1))

            FileO.Read(Values, 0, Values.Length)
         End Using

         If Port >= Values.Length Then ReDim Preserve Values(0 To Port)
         Values(Port) = CByte(Value And &HFF%)
         If Value > &HFF% Then
            If Port + &H1% >= Values.Length Then ReDim Preserve Values(0 To Port + &H1%)
            Values(Port + &H1%) = CByte(Value >> &H8%)
         End If

         Using FileO As New FileStream(IO_FILE, FileMode.Open, FileAccess.Write, FileShare.ReadWrite)
            FileO.Write(Values, 0, Values.Length)
         End Using

         SyncLock Synchronizer
            CPUEvent.Append($"I/O: OUT {Port:X}, {Values(Port):X}{NewLine}")
            If Value > &HFF% Then
               CPUEvent.Append($"I/O: OUT {Port + &H1%:X}, {Values(Port + &H1%):X}{NewLine}")
            End If
         End SyncLock

         Return True
      Catch ExceptionO As Exception
         SyncLock Synchronizer
            CPUEvent.Append($"{ExceptionO.Message}{NewLine}")
         End SyncLock
      End Try

      Return False
   End Function
End Module
