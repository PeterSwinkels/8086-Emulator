'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Emulator8086Program.CPU8086Class
Imports System
Imports System.Drawing

'This class emulates the CGA 320x200 video mode.
Public Class CGA320x200Class
   Implements VideoAdapterClass

   Private Const HEIGHT As Integer = 200             'Defines the graphic mode's height in pixels.
   Private Const INVERT_BIT As Integer = &H80%       'Defines the inverted character color bit.
   Private Const PIXELS_PER_BYTE As Integer = &H4%   'Defines the number of pixels per byte.
   Private Const SCALING As Integer = &H2%           'Defines the scale factor.
   Private Const WIDTH As Integer = 320              'Defines the graphic mode's width in pixels.

   'This procedure clears video adapter's buffer.
   Public Sub ClearBuffer() Implements VideoAdapterClass.ClearBuffer
      Dim Count As Integer = (VideoPageSizesE.CGA320x200A \ &H2%)
      Dim Position As Integer = &H0%

      Do While Count > &H0%
         CPU.PutWord(AddressesE.CGABuffer + Position, &H0%)
         Count -= &H1%
         Position += &H2%
      Loop
   End Sub

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte, ByRef CodePage() As Integer) Implements VideoAdapterClass.Display
      Dim BaseX As New Integer
      Dim GraphicsO As Graphics = Nothing
      Dim Index As New Integer
      Dim Position As New Integer
      Dim Shift As New Integer

      Try
         GraphicsO = Graphics.FromImage(Screen)

         With GraphicsO
            For y2 As Integer = 0 To 1
               Position = AddressesE.CGABuffer + If(y2 = 0, &H0%, VideoPageSizesE.CGA320x200A \ &H2%)
               For y1 As Integer = 0 To HEIGHT - 1 Step 2
                  For x As Integer = 0 To WIDTH - 1 Step PIXELS_PER_BYTE
                     BaseX = x + (PIXELS_PER_BYTE - &H1%)
                     For Pixel As Integer = &H0% To PIXELS_PER_BYTE - &H1%
                        Shift = Pixel * &H2%
                        Index = ((Memory(Position) And (&H3% << Shift)) >> Shift)
                        .FillRectangle(MCC.PaintBrushes(Index), (BaseX - Pixel) * SCALING, (y1 + y2) * SCALING, SCALING, SCALING)
                     Next Pixel
                     Position += &H1%
                  Next x
               Next y1
            Next y2
         End With
      Catch
      Finally
         If GraphicsO IsNot Nothing Then GraphicsO.Dispose()
      End Try
   End Sub

   'This procedure draws the specified character.
   Public Sub DrawCharacter(Index As Integer, Attribute As Integer) Implements VideoAdapterClass.DrawCharacter
      Dim Background As New Byte
      Dim BitSet(&H0% To &H7%) As Boolean
      Dim Character(&H0% To &H7%) As Byte
      Dim Invert As Boolean = (Attribute And INVERT_BIT) = INVERT_BIT
      Dim Position As New Integer
      Dim RemainingBits As New Integer
      Dim Shift As New Integer
      Dim x As New Integer
      Dim y As Integer = Cursor.Y * &H8%

      Array.Copy(CPU.Memory, If(Index < &H80%, AddressesE.Characters + (Index * &H8%), AddressesE.ExtendedCharacters + ((Index - &H80%) * &H8%)), Character, &H0%, Character.Length)

      Attribute = Attribute And &H3%

      For Each ScanLine As Byte In Character
         Position = AddressesE.CGABuffer + If((y And &H1%) = &H0%, &H0%, VideoPageSizesE.CGA320x200A \ &H2%) + ((y \ &H2%) * CGA_320_X_200_BYTES_PER_ROW)

         RemainingBits = ScanLine
         For Bit As Integer = &H7% To &H0% Step -&H1%
            BitSet(Bit) = CBool(RemainingBits And &H1%)
            RemainingBits >>= &H1%
         Next Bit

         x = Cursor.X * &H8%
         Position += (x \ CGA_320_X_200_PIXELS_PER_BYTE)

         Background = CPU.Memory(Position)
         CPU.Memory(Position) = &H0%
         Shift = &H6%
         For Bit As Integer = &H0% To &H7%
            CPU.Memory(Position) = CByte(CPU.Memory(Position) Or If(BitSet(Bit), Attribute << Shift, &H0%))
            Shift -= &H2%

            If Bit = &H3% Then
               Position += &H1%
               CPU.Memory(Position) = &H0%
               Shift = &H6%
               x += 1
            End If
         Next Bit
         y += 1

         If Invert AndAlso Background > &H0% Then
            CPU.PutWord(Position - &H1%, CPU.GetWord(Position - &H1%) Xor &HFFFF%)
         End If
      Next ScanLine
   End Sub

   'This procedure initializes the video adapter.
   Public Sub Initialize() Implements VideoAdapterClass.Initialize
      ClearBuffer()

      CPU.Memory(AddressesE.VideoPage) = &H0%
      ResetCursor()
      MCC.SelectActivePalette(&H1%)
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

      If Count = &H0% OrElse Count > MCC.RowCount() Then
         For Row As Integer = ScrollArea.ULCRow * CGA_320_X_200_LINES_PER_CHARACTER To (ScrollArea.LRCRow + &H1%) * CGA_320_X_200_LINES_PER_CHARACTER
            For Column As Integer = ScrollArea.ULCColumn * &H2% To (ScrollArea.LRCColumn + &H1%) * &H2%
               Address = AddressesE.CGABuffer + If((Row And &H1%) = &H0%, &H0%, VideoPageSizesE.CGA320x200A \ &H2%) + ((Row \ &H2%) * CGA_320_X_200_BYTES_PER_ROW) + Column
               CPU.Memory(Address) = Attribute
            Next Column
         Next Row
      Else
         For Scroll As Integer = &H1% To Count * CGA_320_X_200_LINES_PER_CHARACTER
            Select Case Up
               Case True
                  For Row As Integer = ScrollArea.ULCRow * CGA_320_X_200_LINES_PER_CHARACTER To ScrollArea.LRCRow * CGA_320_X_200_LINES_PER_CHARACTER
                     For Column As Integer = ScrollArea.ULCColumn * &H2% To (ScrollArea.LRCColumn * &H2%) + &H1%
                        Address = AddressesE.CGABuffer + If((Row And &H1%) = &H0%, &H0%, VideoPageSizesE.CGA320x200A \ &H2%) + ((Row \ &H2%) * CGA_320_X_200_BYTES_PER_ROW) + Column
                        CharacterByte = CPU.Memory(Address)
                        CPU.Memory(Address) = Attribute
                        If Row > ScrollArea.ULCRow * CGA_320_X_200_LINES_PER_CHARACTER Then
                           NewAddress = AddressesE.CGABuffer + If(((Row - &H1%) And &H1%) = &H0%, &H0%, VideoPageSizesE.CGA320x200A \ &H2%) + (((Row - &H1%) \ &H2%) * CGA_320_X_200_BYTES_PER_ROW) + Column
                           CPU.Memory(NewAddress) = CharacterByte
                        End If
                     Next Column
                  Next Row
               Case False
                  For Row As Integer = ScrollArea.LRCRow * CGA_320_X_200_LINES_PER_CHARACTER To ScrollArea.ULCRow * CGA_320_X_200_LINES_PER_CHARACTER Step -&H1%
                     For Column As Integer = ScrollArea.ULCColumn * &H2% To (ScrollArea.LRCColumn * &H2%) + &H1%
                        Address = AddressesE.CGABuffer + If((Row And &H1%) = &H0%, &H0%, VideoPageSizesE.CGA320x200A \ &H2%) + ((Row \ &H2%) * CGA_320_X_200_BYTES_PER_ROW) + Column
                        CharacterByte = CPU.Memory(Address)
                        CPU.Memory(Address) = Attribute
                        If Row > ScrollArea.ULCRow * CGA_320_X_200_LINES_PER_CHARACTER Then
                           NewAddress = AddressesE.CGABuffer + If(((Row + &H1%) And &H1%) = &H0%, &H0%, VideoPageSizesE.CGA320x200A \ &H2%) + (((Row + &H1%) \ &H2%) * CGA_320_X_200_BYTES_PER_ROW) + Column
                           CPU.Memory(NewAddress) = CharacterByte
                        End If
                     Next Column
                  Next Row
            End Select
         Next Scroll
      End If
   End Sub
End Class
