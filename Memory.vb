'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Linq

'This class contains the emulated memory related procedures.
Public Class MemoryClass
   Public Const ADDRESS_MASK As Integer = &HFFFFF%   'Defines the 20 bits used to address memory.

   Private ReadOnly Memory() As Byte = {}   'Contains the memory.

   'This procedure initializes this class.
   Public Sub New()
      ReDim Memory(&H0% To &HFFFFF%)
   End Sub

   'This procedure returns the memory as an array.
   Public Function AsArray() As Byte()
      Return Memory
   End Function

   'This procedure returns the memory as a list of bytes.
   Public Function AsList() As List(Of Byte)
      Return Memory.ToList()
   End Function

   'This procedure returns the specified range of bytes from the memory.
   Public Function GetRange(Address As Integer, Length As Integer) As Byte()
      Return Memory.ToList().GetRange(Address, Length).ToArray()
   End Function

   'This procedure manages the byte at the specified address.
   Default Public Property Item(Address As Integer) As Byte
      Get
         Return Memory(Address And ADDRESS_MASK)
      End Get
      Set(NewValue As Byte)
         Memory(Address And ADDRESS_MASK) = NewValue
      End Set
   End Property

   'This procedure returns the word value at the specified address.
   Public Function GetWord(Optional FlatAddress As Integer? = Nothing, Optional Segment As Integer = Nothing, Optional Offset As Integer = Nothing) As Integer
      If FlatAddress IsNot Nothing Then
         Offset = FlatAddress.Value And &HFFFF%
         Segment = FlatAddress.Value And &HF0000%
      Else
         Segment = Segment << &H4%
      End If

      Return Memory((Segment + Offset) And ADDRESS_MASK) Or (CInt(Memory((Segment + ((Offset + &H1%) And &HFFFF%) And ADDRESS_MASK))) << &H8%)
   End Function

   'This procedure returns the memory's length.
   Public ReadOnly Property Length As Integer
      Get
         Return Memory.Length
      End Get
   End Property

   'This procedure puts the specified range at the specified address in memory. 
   Public Sub PutRange(Address As Integer, Bytes() As Byte)
      Array.Copy(Bytes, &H0%, Memory, Address, Bytes.Length)
   End Sub

   'This procedure sets the specified word value at the specified address.
   Public Sub PutWord(Optional FlatAddress As Integer? = Nothing, Optional Segment As Integer = Nothing, Optional Offset As Integer = Nothing, Optional Word As Integer = Nothing)
      If FlatAddress IsNot Nothing Then
         Offset = FlatAddress.Value And &HFFFF%
         Segment = FlatAddress.Value And &HF0000%
      Else
         Segment = Segment << &H4%
      End If

      Memory((Segment + Offset) And ADDRESS_MASK) = CByte(Word And &HFF%)
      Memory((Segment + ((Offset + &H1%) And &HFFFF%)) And ADDRESS_MASK) = CByte((Word And &HFF00%) >> &H8%)
   End Sub
End Class
