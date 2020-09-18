'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO

'This module contains procedures that emulate MS-DOS specific functions.
Public Module MSDOSModule
   'This enumeration lists the MS-DOS executable header's items.
   Private Enum MSDOSHeaderE As Integer
      Signature = &H0%                'Defines the MS-DOS executable's signature.
      ImageSizeModulo = &H2%          'Defines the executable's image size modulo of 0x200 bytes.
      ImageSize = &H4%                'Defines the executable's image size in pages of 0x200 bytes.
      RelocationCount = &H6%          'Defines the executable's number or relocation items.
      HeaderSize = &H8%               'Defines the executable's header size in paragraphs of 0x10 bytes.
      MinimumParagraphs = &HA%        'Defines the executable's minimum memory requirement in paragraphs of 0x10 bytes.
      MaximumParagraphs = &HC%        'Defines the executable's maximum memory requirement in paragraphs of 0x10 bytes.
      StackSegment = &HE%             'Defines the stack segment (SS) register's initial value.
      StackPointer = &H10%            'Defines the stack pointer (SP) register's initial value.
      Checksum = &H12%                'Defines the executable's negative pgm checksum.
      InstructionPointer = &H14%      'Defines the instruction pointer (IP) register's initial value.
      CodeSegment = &H16%             'Defines the code segment (CS) register's initial value.
      RelocationTableOffset = &H18%   'Defines the relocation table's offset.
      OverlayNumber = &H1A%           'Defines the overlay number.
   End Enum

   'This structure defines the relocation items table.
   Private Structure RelocationItemsStr
      Public Count As Integer                         'Defines the number of relocation items.
      Public Ending As Integer                        'Defines the end of the relocation item.
      Public Positions As List(Of SegmentOffsetStr)   'Defines the list of relocation item positions.
      Public Start As Integer                         'Defines the start of the relocation items.
   End Structure

   'This structure defines a segment and offset.
   Public Structure SegmentOffsetStr
      Public Segment As Integer  'Defines a segment.
      Public Offset As Integer   'Defines an offset.
   End Structure

   Private Const SIGNATURE As Integer = &H5A4D%   'Defines the MS-DOS executable's signature.

   'This procedure returns the relocation item positions for the specified executable.
   Private Function GetRelocationItemPositions(Executable As List(Of Byte)) As RelocationItemsStr
      Try
         Dim RelocationItems As New RelocationItemsStr With {.Count = &H0%, .Ending = &H0%, .Positions = New List(Of SegmentOffsetStr), .Start = &H0%}
         Dim Position As New Integer

         With RelocationItems
            .Start = BitConverter.ToUInt16(Executable.ToArray(), MSDOSHeaderE.RelocationTableOffset)
            .Count = BitConverter.ToUInt16(Executable.ToArray(), MSDOSHeaderE.RelocationCount)
            .Ending = BitConverter.ToUInt16(Executable.ToArray(), MSDOSHeaderE.HeaderSize) << &H4%

            Position = .Start
            For RelocationItem As Integer = &H0% To .Count - &H1%
               .Positions.Add(New SegmentOffsetStr With {.Offset = BitConverter.ToUInt16(Executable.ToArray(), Position), .Segment = BitConverter.ToUInt16(Executable.ToArray(), Position + &H2%)})
               Position += &H4%
            Next RelocationItem
         End With

         Return RelocationItems
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns the word at the specified address.
   Private Function GetWord(Data As List(Of Byte), Address As Integer) As Integer
      Try
         Return Data(Address) Or (CInt(Data(Address + &H1%)) << &H8%)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure loads the specified MS-DOS executable and processes its header.
   Public Function LoadExecutable(FileName As String) As Boolean
      Try
         Dim Address As New Integer
         Dim CS As Integer = CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.CS))
         Dim Executable As New List(Of Byte)(File.ReadAllBytes(FileName))
         Dim IP As Integer = CInt(CPU.Registers(CPU8086Class.Registers16BitE.IP))
         Dim Offset As Integer = GetFlatCSIP()

         If BitConverter.ToUInt16(Executable.ToArray(), MSDOSHeaderE.Signature) = SIGNATURE Then
            CPU.Registers(CPU8086Class.SegmentRegistersE.CS, NewValue:=CS + BitConverter.ToUInt16(Executable.ToArray(), MSDOSHeaderE.CodeSegment))
            CPU.Registers(CPU8086Class.Registers16BitE.IP, NewValue:=BitConverter.ToUInt16(Executable.ToArray(), MSDOSHeaderE.InstructionPointer))
            CPU.Registers(CPU8086Class.SegmentRegistersE.ES, NewValue:=CPU.Registers(CPU8086Class.SegmentRegistersE.DS))
            CPU.Registers(CPU8086Class.SegmentRegistersE.SS, NewValue:=CS + BitConverter.ToUInt16(Executable.ToArray(), MSDOSHeaderE.StackSegment))
            CPU.Registers(CPU8086Class.Registers16BitE.SP, NewValue:=BitConverter.ToUInt16(Executable.ToArray(), MSDOSHeaderE.StackPointer))

            CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=&H0%)
            CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=Executable.Count >> &H10%)
            CPU.Registers(CPU8086Class.Registers16BitE.CX, NewValue:=Executable.Count And &HFFFF%)
            CPU.Registers(CPU8086Class.Registers16BitE.DX, NewValue:=&H0%)

            With GetRelocationItemPositions(Executable)
               Executable.RemoveRange(&H0%, .Ending + &H1%)

               For Each RelocationItemPosition As SegmentOffsetStr In .Positions
                  Address = (RelocationItemPosition.Segment << &H4%) Or RelocationItemPosition.Offset
                  PutWord(Executable, Address, GetWord(Executable, Address) + CS)
               Next RelocationItemPosition
            End With

            Executable.CopyTo(CPU.Memory, Offset)

            Return True
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function

   'This procedure writes the specified word to the specified address in the specified data.
   Private Sub PutWord(Data As List(Of Byte), Address As Integer, Word As Integer)
      Try
         Data(Address) = CByte(Word And &HFF%)
         Data(Address + &H1%) = CByte((Word And &HFF00%) >> &H8%)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Module
