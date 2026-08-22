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
   Public Function GetWord(Address As Integer) As Integer
      Return Memory(Address And ADDRESS_MASK) Or (CInt(Memory(((Address And ADDRESS_MASK) + &H1%) And &HFFFF%)) << &H8%)
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
   Public Sub PutWord(Address As Integer, NewValue As Integer)
      Memory(Address And ADDRESS_MASK) = CByte(NewValue And &HFF%)
      Memory(((Address And ADDRESS_MASK) + &H1%) And &HFFFF%) = CByte((NewValue >> &H8%) And &HFF%)
   End Sub
End Class
