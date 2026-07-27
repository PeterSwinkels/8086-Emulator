'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Emulator8086Program.CPU8086Class
Imports System
Imports System.Drawing

'This class emulates the CGA 640x200 video mode.
Public Class CGA640x200Class
   Implements VideoAdapterClass

   Private Const HEIGHT As Integer = 200                'Defines the graphics mode's height in pixels.
   Private Const PIXELS_PER_BYTE As Integer = &H8%      'Defines the number of pixels per byte.
   Private Const HORIZONTAL_SCALING As Integer = &H1%   'Defines the horizontal scale factor.
   Private Const VERTICAL_SCALING As Integer = &H2%     'Defines the vertical scale factor.
   Private Const WIDTH As Integer = 640                 'Defines the graphics mode's width in pixels.

   Private ReadOnly BLACK_BRUSH As New SolidBrush(Color.Black)   'Defines a black brush.
   Private ReadOnly WHITE_BRUSH As New SolidBrush(Color.White)   'Defines a white brush.

   'This procedure clears video adapter's buffer.
   Public Sub ClearBuffer() Implements VideoAdapterClass.ClearBuffer
      Dim Count As Integer = (VideoPageSizesE.CGA640x200 \ &H2%)
      Dim Position As Integer = &H0%

      Do While Count > &H0%
         CPU.PutWord(AddressesE.CGABuffer + Position, &H0%)
         Count -= &H1%
         Position += &H2%
      Loop
   End Sub

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte, CodePage() As Integer) Implements VideoAdapterClass.Display
      Dim BaseX As New Integer
      Dim GraphicsO As Graphics = Nothing
      Dim PixelOff As New Boolean
      Dim Position As New Integer

      Try
         GraphicsO = Graphics.FromImage(Screen)

         With GraphicsO
            For y2 As Integer = 0 To 1
               Position = AddressesE.CGABuffer + If(y2 = 0, &H0%, VideoPageSizesE.CGA640x200 \ &H2%)
               For y1 As Integer = 0 To HEIGHT - 1 Step 2
                  For x As Integer = 0 To WIDTH - 1 Step PIXELS_PER_BYTE
                     BaseX = x + (PIXELS_PER_BYTE - &H1%)
                     For Pixel As Integer = &H0% To PIXELS_PER_BYTE - &H1%
                        PixelOff = (((Memory(Position) And (&H1% << Pixel)) >> Pixel) = &H0%)
                        .FillRectangle(If(PixelOff, BLACK_BRUSH, WHITE_BRUSH), (BaseX - Pixel) * HORIZONTAL_SCALING, (y1 + y2) * VERTICAL_SCALING, HORIZONTAL_SCALING, VERTICAL_SCALING)
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
      Dim Position As New Integer
      Dim RemainingBits As New Integer
      Dim Shift As New Integer
      Dim x As New Integer
      Dim y As Integer = Cursor.Y * &H8%

      Array.Copy(CPU.Memory, If(Index < &H80%, AddressesE.Characters, AddressesE.ExtendedCharacters) + (Index * &H8%), Character, &H0%, Character.Length)

      Attribute = Attribute And &H1%

      For Each ScanLine As Byte In Character
         Position = AddressesE.CGABuffer + If((y And &H1%) = &H0%, &H0%, VideoPageSizesE.CGA640x200 \ &H2%) + ((y \ 2) * CGA_640_X_200_BYTES_PER_ROW)

         RemainingBits = ScanLine
         For Bit As Integer = &H7% To &H0% Step -&H1%
            BitSet(Bit) = CBool(RemainingBits And &H1%)
            RemainingBits = RemainingBits >> &H1%
         Next Bit

         x = Cursor.X * &H8%
         Position += (x \ CGA_640_X_200_PIXELS_PER_BYTE)

         Background = CPU.Memory(Position)
         CPU.Memory(Position) = &H0%
         Shift = &H7%
         For Bit As Integer = &H0% To &H7%
            CPU.Memory(Position) = CByte(CPU.Memory(Position) Or If(BitSet(Bit), Attribute << Shift, &H0%))
            Shift -= &H1%
         Next Bit
         y += 1
      Next ScanLine
   End Sub

   'This procedure initializes the video adapter.
   Public Sub Initialize() Implements VideoAdapterClass.Initialize
      ClearBuffer()

      CPU.Memory(AddressesE.VideoPage) = &H0%
      ResetCursor()
   End Sub

   'This procedure returns the screen size used by a video adapter.
   Public Function Resolution() As Size Implements VideoAdapterClass.Resolution
      Return New Size(WIDTH * HORIZONTAL_SCALING, HEIGHT * VERTICAL_SCALING)
   End Function

   'This procedure scrolls the video adapter's buffer.
   Public Sub ScrollBuffer(Up As Boolean, ScrollArea As VideoAdapterClass.ScreenAreaStr, Count As Integer) Implements VideoAdapterClass.ScrollBuffer
      Dim Address As New Integer
      Dim Attribute As Byte = CByte(CPU.Registers(SubRegisters8BitE.BH) * &HFF%)
      Dim CharacterByte As New Byte
      Dim NewAddress As New Integer

      If Count = &H0% OrElse Count > MCC.RowCount() Then
         VideoAdapter.ClearBuffer()
      Else
         For Scroll As Integer = &H1% To Count * CGA_640_X_200_LINES_PER_CHARACTER
            Select Case Up
               Case True
                  For Row As Integer = ScrollArea.ULCRow * CGA_640_X_200_LINES_PER_CHARACTER To ScrollArea.LRCRow * CGA_640_X_200_LINES_PER_CHARACTER
                     For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
                        Address = AddressesE.CGABuffer + If((Row And &H1%) = &H0%, &H0%, VideoPageSizesE.CGA640x200 \ &H2%) + ((Row \ &H2%) * CGA_640_X_200_BYTES_PER_ROW) + Column
                        CharacterByte = CPU.Memory(Address)
                        CPU.Memory(Address) = Attribute
                        If Row > ScrollArea.ULCRow * CGA_640_X_200_LINES_PER_CHARACTER Then
                           NewAddress = AddressesE.CGABuffer + If(((Row - &H1%) And &H1%) = &H0%, &H0%, VideoPageSizesE.CGA640x200 \ &H2%) + (((Row - &H1%) \ &H2%) * CGA_640_X_200_BYTES_PER_ROW) + Column
                           CPU.Memory(NewAddress) = CharacterByte
                        End If
                     Next Column
                  Next Row
               Case False
                  For Row As Integer = ScrollArea.LRCRow * CGA_640_X_200_LINES_PER_CHARACTER To ScrollArea.ULCRow * CGA_640_X_200_LINES_PER_CHARACTER Step -&H1%
                     For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
                        Address = AddressesE.CGABuffer + If((Row And &H1%) = &H0%, &H0%, VideoPageSizesE.CGA640x200 \ &H2%) + ((Row \ &H2%) * CGA_640_X_200_BYTES_PER_ROW) + Column
                        CharacterByte = CPU.Memory(Address)
                        CPU.Memory(Address) = Attribute
                        If Row > ScrollArea.ULCRow * CGA_640_X_200_LINES_PER_CHARACTER Then
                           NewAddress = AddressesE.CGABuffer + If(((Row + &H1%) And &H1%) = &H0%, &H0%, VideoPageSizesE.CGA640x200 \ &H2%) + (((Row + &H1%) \ &H2%) * CGA_640_X_200_BYTES_PER_ROW) + Column
                           CPU.Memory(NewAddress) = CharacterByte
                        End If
                     Next Column
                  Next Row
            End Select
         Next Scroll
      End If
   End Sub
End Class
