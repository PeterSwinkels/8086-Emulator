'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Environment
Imports System.IO
Imports System.Linq

'This module handles loading MS-DOS executables.
Public Module MSDOSEXEModule
   Private Const RELOCATION_ITEM_COUNT As Integer = &H6%    'Defines where the number of relocation items is stored.
   Private Const RELOCATION_ITEM_TABLE As Integer = &H18%   'Defines where the number of relocation item table offset is stored.
   Private Const INITIAL_CS As Integer = &H16%              'Defines where the initial code segment is stored.
   Private Const INITIAL_IP As Integer = &H14%              'Defines where the initial instruction pointer is stored.
   Private Const INITIAL_SP As Integer = &H10%              'Defines where the initial stack pointer is stored.
   Private Const INITIAL_SS As Integer = &HE%               'Defines where the initial stack segment is stored.

   Private ReadOnly MZ_SIGNATURE() As Byte = {&H4D%, &H5A%}  'Defines the signature of an MZ exectuable.

   'This procedure loads the MS-DOS executable file into the emulated CPU's memory after processing its header.
   Public Sub LoadMSDOSEXE(FileName As String)
      Try
         Dim Executable As New List(Of Byte)(File.ReadAllBytes(FileName))
         Dim Offset As New Integer

         If Executable.GetRange(0, MZ_SIGNATURE.Length).SequenceEqual(MZ_SIGNATURE) Then
            LoadMZExe(Executable, FileName)
         Else
            Offset = GetFlatCSIP()
            Output.AppendText($"Loading the compact binary executable ""{FileName}"" at address {Offset:X8}.{NewLine}")
            Executable.CopyTo(CPU.Memory, Offset)
            CPU.Registers(CPU8086Class.SegmentRegistersE.DS, NewValue:=CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.CS)) - &H10%)
            CPU.Registers(CPU8086Class.Registers16BitE.SP, NewValue:=&HFFFF%)
            CPU.Registers(CPU8086Class.SegmentRegistersE.SS, NewValue:=CPU.Registers(CPU8086Class.SegmentRegistersE.CS))
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   Private Sub LoadMZExe(Executable As List(Of Byte), FileName As String)

      Try
         Dim RelocationTableSize As Integer = (BitConverter.ToUInt16(Executable.ToArray(), RELOCATION_ITEM_COUNT) * &H4%)
         Dim HeaderSize As Integer = BitConverter.ToUInt16(Executable.ToArray(), RELOCATION_ITEM_TABLE) + RelocationTableSize
         Dim Offset As Integer = CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H4%
         Dim ProcessedExecutable As New List(Of Byte)(From ByteO In Executable Skip HeaderSize)
         Dim RelocatedCS As Integer = (CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) + (HeaderSize >> &H4%) + &H1%) + Executable(INITIAL_CS)
         Dim RelocationItemFlatAddress As New Integer
         Dim RelocationItemOffset As New Integer
         Dim RelocationItemSegment As New Integer

         If Offset + ProcessedExecutable.Count <= CPU.Memory.Length Then
            Output.AppendText($"Loading the MZ-executable ""{FileName}"" at address {Offset:X8}.{NewLine}")

            CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=&H0%)
            CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=ProcessedExecutable.Count >> &H10%)
            CPU.Registers(CPU8086Class.Registers16BitE.CX, NewValue:=ProcessedExecutable.Count And &HFFFF%)
            CPU.Registers(CPU8086Class.Registers16BitE.DX, NewValue:=&H0%)
            CPU.Registers(CPU8086Class.Registers16BitE.SP, NewValue:=BitConverter.ToUInt16(Executable.ToArray(), INITIAL_SP))
            CPU.Registers(CPU8086Class.SegmentRegistersE.CS, NewValue:=RelocatedCS)
            CPU.Registers(CPU8086Class.SegmentRegistersE.ES, NewValue:=CPU.Registers(CPU8086Class.SegmentRegistersE.DS))
            CPU.Registers(CPU8086Class.SegmentRegistersE.SS, NewValue:=(RelocatedCS + BitConverter.ToUInt16(Executable.ToArray(), INITIAL_SS)))

            For Position As Integer = BitConverter.ToUInt16(Executable.ToArray(), RELOCATION_ITEM_TABLE) To BitConverter.ToUInt16(Executable.ToArray(), RELOCATION_ITEM_TABLE) + RelocationTableSize Step &H4%
               RelocationItemOffset = BitConverter.ToInt16(Executable.ToArray(), Position)
               RelocationItemSegment = BitConverter.ToInt16(Executable.ToArray(), Position + &H2%)
               RelocationItemFlatAddress = (RelocationItemSegment << &H4%) Or RelocationItemOffset

               ProcessedExecutable(RelocationItemFlatAddress) = ToByte((CInt(ProcessedExecutable(RelocationItemFlatAddress)) + (RelocatedCS >> &H8%)) And &HFF%)
               ProcessedExecutable(RelocationItemFlatAddress + &H1%) = ToByte((CInt(ProcessedExecutable(RelocationItemFlatAddress + &H1%)) + (RelocatedCS And &HFF%)) And &HFF%)
            Next Position

            Executable.CopyTo(CPU.Memory, Offset)
         Else
            Output.AppendText($"""{FileName}"" does not fit inside the emulated memory.{NewLine}")
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

   End Sub
End Module
