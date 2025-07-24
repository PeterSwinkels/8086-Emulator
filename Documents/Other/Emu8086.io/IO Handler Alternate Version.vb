'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Environment
Imports System.IO

'This module contains the default I/O handler.
Public Module IOHandlerModule
   Private Const IO_FILE As String = "C:\Emu8086.io"   'Defines the default I/O file path.

   'This procedure attempts to read from the specified I/O port and returns the result.
   Public Function ReadIOPort(Port As Integer) As Integer?
      Try
         Dim Values() As Byte = {}

         Using FileO As New FileStream(IO_FILE, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            ReDim Values(0 To CInt(FileO.Length - 1))

            FileO.Read(Values, 0, Values.Length)
         End Using

         Return Values(Port)
      Catch ExceptionO As Exception
         CPUEvent.Append($"{ExceptionO.Message}{NewLine}")
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

         Return True
      Catch ExceptionO As Exception
         CPUEvent.Append($"{ExceptionO.Message}{NewLine}")
      End Try

      Return False
   End Function
End Module
