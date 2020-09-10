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
   'This enumeration list the MS-DOS executable header's items.
   Private Enum MSDOSHeaderE As Integer
      Signature = &H0%                'The MS-DOS executable's signature.
      ImageSizeModulo = &H2%          'The executable's image size modulo of 0x200 bytes.
      ImageSize = &H4%                'The executable's image size in pages of 0x200 bytes.
      RelocationCount = &H6%          'The executable's number or relocation items.
      HeaderSize = &H8%               'The executable's header size in paragraphs of 0x10 bytes.
      MinimumParagraphs = &HA%        'The executable's minimum memory requirement in paragraphs of 0x10 bytes.
      MaximumParagraphs = &HC%        'The executable's maximum memory requirement in paragraphs of 0x10 bytes.
      StackSegment = &HE%             'The stack segment (SS) register's initial value.
      StackPointer = &H10%            'The stack pointer (SP) register's initial value.
      Checksum = &H12%                'The executable's negative pgm checksum.
      InstructionPointer = &H14%      'The instruction pointer (IP) register's initial value.
      CodeSegment = &H16%             'The code segment (CS) register's initial value.
      RelocationTableOffset = &H18%   'The relocation table's offset.
      OverlayNumber = &H1A%           'The overlay number.
   End Enum

   'This structure defines the relocation items table.
   Private Structure RelocationItemsStr
      Public Count As Integer                'Defines the number of relocation items.
      Public Ending As Integer               'Defines the end of the relocation item.
      Public Positions As List(Of Integer)   'Defines the list of relocation item positions.
      Public Start As Integer                'Defines the start of the relocation item.
   End Structure

   Private Const SIGNATURE As Integer = &H5A4D%   'Defines the MS-DOS executable's signature.

   'This procedure returns the relocation item positions for the specified executable.
   Private Function GetRelocationItemPositions(Executable As List(Of Byte)) As RelocationItemsStr
      Try
         Dim RelocationItems As New RelocationItemsStr With {.Count = &H0%, .Ending = &H0%, .Positions = New List(Of Integer), .Start = &H0%}
         Dim Position As New Integer

         With RelocationItems
            .Start = BitConverter.ToUInt16(Executable.ToArray(), MSDOSHeaderE.RelocationTableOffset)
            .Count = BitConverter.ToUInt16(Executable.ToArray(), MSDOSHeaderE.RelocationCount)
            .Ending = BitConverter.ToUInt16(Executable.ToArray(), MSDOSHeaderE.HeaderSize) << &H4%

            Position = .Start
            For RelocationItem As Integer = &H0% To .Count - &H1%
               .Positions.Add(.Ending + BitConverter.ToUInt16(Executable.ToArray(), Position) + (BitConverter.ToUInt16(Executable.ToArray(), Position + &H2%) << &H4%))
               Position += &H4%
            Next RelocationItem
         End With

         Return RelocationItems
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure loads the specified MS-DOS executable and processes its header.
   Public Function LoadExecutable(FileName As String) As Boolean
      Try
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
               For Each RelocationItemPosition As Integer In .Positions
                  PutWord(Executable, RelocationItemPosition, BitConverter.ToUInt16(Executable.ToArray(), RelocationItemPosition) + IP)
                  PutWord(Executable, RelocationItemPosition + &H2%, BitConverter.ToUInt16(Executable.ToArray(), RelocationItemPosition + &H2%) + CS)
               Next RelocationItemPosition

               Executable.RemoveRange(&H0%, .Ending + &H1%)
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
