'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Drawing

'This class emulates the CGA 320x200 video mode.
Public Class CGA320x200Class
   Implements VideoAdapterClass

   Private Const HEIGHT As Integer = 200             'Defines the graphic's mode height in pixels.
   Private Const PIXELS_PER_BYTE As Integer = &H4%   'Defines the number of pixels per byte.
   Private Const SCALING As Integer = &H2%           'Defines the scale factor.
   Private Const WIDTH As Integer = 320              'Defines the graphic's mode width in pixels.

   'This procedure clears video adapter's buffer.
   Public Sub ClearBuffer() Implements VideoAdapterClass.ClearBuffer
      Dim Count As Integer = (CGA_320_x_200_BUFFER_SIZE \ &H2%)
      Dim Position As Integer = &H0%

      Do While Count > &H0%
         CPU.PutWord(AddressesE.CGA320x200 + Position, &H0%)
         Count -= &H1%
         Position += &H2%
      Loop
   End Sub

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte, CodePage() As Integer) Implements VideoAdapterClass.Display
      Dim Index As New Integer
      Dim Position As New Integer

      With Graphics.FromImage(Screen)
         For y2 As Integer = 0 To 1
            Position = AddressesE.CGA320x200 + If(y2 = 0, &H0%, CGA_320_x_200_BUFFER_SIZE \ &H2%)
            For y1 As Integer = 0 To HEIGHT - 1 Step 2
               For x As Integer = 0 To WIDTH - 1 Step PIXELS_PER_BYTE
                  For Pixel As Integer = &H0% To PIXELS_PER_BYTE - &H1%
                     Index = ((CPU.Memory(Position) And (&H3% << (Pixel * &H2%))) >> (Pixel * &H2%))
                     .FillRectangle(PaletteBrushes(Index), (x + (PIXELS_PER_BYTE - &H1%) - Pixel) * SCALING, (y1 + y2) * SCALING, SCALING, SCALING)
                  Next Pixel
                  Position += &H1%
               Next x
            Next y1
         Next y2
      End With
   End Sub

   'This procedure draws the specified character.
   Public Sub DrawCharacter(Index As Integer, Attribute As Integer) Implements VideoAdapterClass.DrawCharacter
      Dim Bits(&H0% To &H7%) As Boolean
      Dim Character(&H0% To &H7%) As Byte
      Dim Position As New Integer
      Dim Shift As New Integer
      Dim x As New Integer
      Dim y As Integer = Cursor.Y * &H8%
      Dim Value As New Integer

      Array.Copy(CPU.Memory, If(Index < &H80%, AddressesE.Characters, AddressesE.ExtendedCharacters) + (Index * &H8%), Character, &H0%, Character.Length)

      For Each [Byte] As Byte In Character
         Position = AddressesE.CGA320x200 + If((y And &H1%) = &H0%, &H0%, CGA_320_x_200_BUFFER_SIZE \ &H2%) + ((y \ 2) * CGA_320_X_200_BYTES_PER_ROW)

         Value = [Byte]
         For Bit As Integer = &H7% To &H0% Step -&H1%
            Bits(Bit) = CBool(Value And &H1%)
            Value = Value >> &H1%
         Next Bit

         x = Cursor.X * &H8%
         Position += (x \ CGA_320_X_200_PIXELS_PER_BYTE)

         Shift = &H6%
         For Bit As Integer = &H0% To &H7%
            If Bits(Bit) Then
               If Not (Attribute And &H80%) = &H0% Then
                  CPU.Memory(Position) = CByte((CPU.Memory(Position) Xor ((Attribute And &H7F%) << Shift)))
               Else
                  CPU.Memory(Position) = CByte((CPU.Memory(Position) And Not (&H3% << Shift)) Or ((Attribute << Shift) And &HFF%))
               End If
            End If

            Shift -= &H2%

            If Bit = &H3% Then
               Position += &H1%
               Shift = &H6%
               x += 1
            End If
         Next Bit
         y += 1
      Next [Byte]
   End Sub

   'This procedure initializes the video adapter.
   Public Sub Initialize() Implements VideoAdapterClass.Initialize
      ClearBuffer()

      CPU.Memory(AddressesE.VideoPage) = &H0%
      CPU.PutWord(AddressesE.CursorPositions, Word:=&H0%)
      CPU.PutWord(AddressesE.CursorScanLines, Word:=CURSOR_DEFAULT)
      CursorBlink.Enabled = False

      MCC.ColorSelect(&H20%)
   End Sub

   'This procedure returns the screen size used by a video adapter.
   Public Function Resolution() As Size Implements VideoAdapterClass.Resolution
      Return New Size(WIDTH * SCALING, HEIGHT * SCALING)
   End Function

   'This procedure scrolls the video adapter's buffer.
   Public Sub ScrollBuffer(Up As Boolean, ScrollArea As VideoAdapterClass.ScreenAreaStr, Count As Integer) Implements VideoAdapterClass.ScrollBuffer
      Dim Address As New Integer
      Dim Attribute As Byte = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.BH))
      Dim CharacterByte1 As New Byte
      Dim CharacterByte2 As New Byte
      Dim NewAddress As New Integer
      Dim Position As New Integer

      If Count = &H0% OrElse Count > CGA_320_X_200_LINE_COUNT Then
         VideoAdapter.ClearBuffer()
      Else
         For Scroll As Integer = &H1% To Count * (CGA_320_X_200_LINES_PER_CHARACTER \ &H2%)
            Select Case Up
               Case True
                  For Row As Integer = ScrollArea.ULCRow * (CGA_320_X_200_LINES_PER_CHARACTER \ &H2%) To (ScrollArea.LRCRow * (CGA_320_X_200_LINES_PER_CHARACTER \ &H2%)) + ((CGA_320_X_200_LINES_PER_CHARACTER \ &H2%) - &H1%)
                     For Column As Integer = ScrollArea.ULCColumn * &H2% To (ScrollArea.LRCColumn * &H2%) + &H1%
                        Address = AddressesE.CGA320x200 + ((Row * CGA_320_X_200_BYTES_PER_ROW) + Column)
                        CharacterByte1 = CPU.Memory(Address)
                        CharacterByte2 = CPU.Memory((CGA_320_x_200_BUFFER_SIZE \ &H2%) + Address)
                        CPU.Memory(Address) = Attribute
                        CPU.Memory((CGA_320_x_200_BUFFER_SIZE \ &H2%) + Address) = Attribute
                        If Row > ScrollArea.ULCRow * (CGA_320_X_200_LINES_PER_CHARACTER \ &H2%) Then
                           NewAddress = AddressesE.CGA320x200 + ((Row - &H1%) * CGA_320_X_200_BYTES_PER_ROW) + Column
                           CPU.Memory(NewAddress) = CharacterByte1
                           CPU.Memory((CGA_320_x_200_BUFFER_SIZE \ &H2%) + NewAddress) = CharacterByte2
                        End If
                     Next Column
                  Next Row
               Case False
                  For Row As Integer = (ScrollArea.LRCRow * (CGA_320_X_200_LINES_PER_CHARACTER \ &H2%)) + ((CGA_320_X_200_LINES_PER_CHARACTER \ &H2%) - &H1%) To ScrollArea.ULCRow * (CGA_320_X_200_LINES_PER_CHARACTER \ &H2%) Step -&H1%
                     For Column As Integer = ScrollArea.ULCColumn * &H2% To (ScrollArea.LRCColumn * &H2%) + &H1%
                        Address = AddressesE.CGA320x200 + ((Row * CGA_320_X_200_BYTES_PER_ROW) + Column)
                        CharacterByte1 = CPU.Memory(Address)
                        CharacterByte2 = CPU.Memory((CGA_320_x_200_BUFFER_SIZE \ &H2%) + Address)
                        CPU.Memory(Address) = Attribute
                        CPU.Memory((CGA_320_x_200_BUFFER_SIZE \ &H2%) + Address) = Attribute
                        If Row < ScrollArea.LRCRow * (CGA_320_X_200_LINES_PER_CHARACTER \ &H2%) Then
                           NewAddress = AddressesE.CGA320x200 + ((Row + &H1%) * CGA_320_X_200_BYTES_PER_ROW) + Column
                           CPU.Memory(NewAddress) = CharacterByte1
                           CPU.Memory((CGA_320_x_200_BUFFER_SIZE \ &H2%) + NewAddress) = CharacterByte2
                        End If
                     Next Column
                  Next Row
            End Select
         Next Scroll
      End If
   End Sub
End Class
