'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Emulator8086Program.CPU8086Class
Imports System
Imports System.Drawing

'This class emulates VGA 320x200.
Public Class VGA320x200Class
   Implements VideoAdapterClass

   Private Const HEIGHT As Integer = 200     'Defines the graphic's mode height in pixels.
   Private Const SCALING As Integer = &H2%   'Defines the scale factor.
   Private Const WIDTH As Integer = 320      'Defines the graphic's mode width in pixels.

   'This procedure clears video adapter's buffer.
   Public Sub ClearBuffer() Implements VideoAdapterClass.ClearBuffer
      Dim Count As Integer = VGA_320_x_200_BUFFER_SIZE \ &H2%
      Dim Position As Integer = AddressesE.VGA320x200

      Do While Count > &H0%
         CPU.PutWord(Position, &H0%)
         Count -= &H1%
         Position += &H2%
      Loop
   End Sub

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte, CodePage() As Integer) Implements VideoAdapterClass.Display
      Dim GraphicsO As Graphics = Graphics.FromImage(Screen)

      With GraphicsO
         For y As Integer = 0 To HEIGHT - 1
            For x As Integer = 0 To WIDTH - 1
               .FillRectangle(VGA.VGAPaintBrushes(Memory(AddressesE.VGA320x200 + ((y * WIDTH) + x))), x * SCALING, y * SCALING, SCALING, SCALING)
            Next x
         Next y
      End With
   End Sub

   'This procedure draws the specified character.
   Public Sub DrawCharacter(Index As Integer, Attribute As Integer) Implements VideoAdapterClass.DrawCharacter
      Dim BitSet(&H0% To &H7%) As Boolean
      Dim Character(&H0% To &H7%) As Byte
      Dim Position As New Integer
      Dim RemainingBits As New Integer
      Dim y As Integer = Cursor.Y * &H8%

      Array.Copy(CPU.Memory, If(Index < &H80%, AddressesE.Characters, AddressesE.ExtendedCharacters) + (Index * &H8%), Character, &H0%, Character.Length)

      For Each ScanLine As Byte In Character
         RemainingBits = ScanLine
         For Bit As Integer = &H7% To &H0% Step -&H1%
            BitSet(Bit) = CBool(RemainingBits And &H1%)
            RemainingBits = RemainingBits >> &H1%
         Next Bit

         Position = AddressesE.VGA320x200 + (y * VGA_320_X_200_BYTES_PER_ROW) + (Cursor.X * &H8%)

         For Bit As Integer = &H0% To &H7%
            CPU.Memory(Position) = CByte(Attribute * Math.Abs(CInt(BitSet(Bit))))
            Position += &H1%
         Next Bit
         y += 1
      Next ScanLine
   End Sub

   'This procedure initializes the video adapter.
   Public Sub Initialize() Implements VideoAdapterClass.Initialize
      ClearBuffer()

      For Index As Integer = VGA.VGA_DEFAULT_PALETTE.GetLowerBound(0) To VGA.VGA_DEFAULT_PALETTE.GetUpperBound(0)
         VGA.VGAPaintBrushes(Index) = New SolidBrush(Color.FromArgb(VGA.VGA_DEFAULT_PALETTE(Index) Or &HFF000000%))
      Next Index

      CPU.Memory(AddressesE.VideoPage) = &H0%
      ResetCursor()
      CursorBlink.Enabled = False
   End Sub

   'This procedure returns the screen size used by a video adapter.
   Public Function Resolution() As Size Implements VideoAdapterClass.Resolution
      Return New Size(WIDTH * SCALING, HEIGHT * SCALING)
   End Function

   'This procedure scrolls the video adapter's buffer.
   Public Sub ScrollBuffer(Up As Boolean, ScrollArea As VideoAdapterClass.ScreenAreaStr, Count As Integer) Implements VideoAdapterClass.ScrollBuffer
      Dim Address As New Integer
      Dim Attribute As Byte = CByte(CPU.Registers(SubRegisters8BitE.BH))
      Dim CharacterByte As New Byte
      Dim NewAddress As New Integer
      Dim Position As New Integer

      If Count = &H0% OrElse Count > MCC.RowCount() Then
         VideoAdapter.ClearBuffer()
      Else
         For Scroll As Integer = &H1% To Count * VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE
            Select Case Up
               Case True
                  For Row As Integer = ScrollArea.ULCRow * VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE To (ScrollArea.LRCRow * VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE) + (VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE - &H1%)
                     For Column As Integer = ScrollArea.ULCColumn * VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE To (ScrollArea.LRCColumn * VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE) + VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE
                        Address = AddressesE.VGA320x200 + ((Row * VGA_320_X_200_BYTES_PER_ROW) + Column)
                        CharacterByte = CPU.Memory(Address)
                        CPU.Memory(Address) = Attribute
                        If Row > ScrollArea.ULCRow * VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE Then
                           NewAddress = AddressesE.VGA320x200 + ((Row - &H1%) * VGA_320_X_200_BYTES_PER_ROW) + Column
                           CPU.Memory(NewAddress) = CharacterByte
                        End If
                     Next Column
                  Next Row
               Case False
                  For Row As Integer = ScrollArea.LRCRow * VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE To ScrollArea.ULCRow * VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE Step -&H1%
                     For Column As Integer = ScrollArea.ULCColumn * VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE To (ScrollArea.LRCColumn * VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE) + VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE
                        Address = AddressesE.VGA320x200 + ((Row * VGA_320_X_200_BYTES_PER_ROW) + Column)
                        CharacterByte = CPU.Memory(Address)
                        CPU.Memory(Address) = Attribute
                        If Row < ScrollArea.LRCRow * VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE Then
                           NewAddress = AddressesE.VGA320x200 + ((Row + &H1%) * VGA_320_X_200_BYTES_PER_ROW) + Column
                           CPU.Memory(NewAddress) = CharacterByte
                        End If
                     Next Column
                  Next Row
            End Select
         Next Scroll
      End If
   End Sub
End Class
