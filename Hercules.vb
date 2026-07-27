'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Emulator8086Program.CPU8086Class
Imports System
Imports System.Drawing

'This class emulates the Hercules video mode.
Public Class HerculesClass
   Implements VideoAdapterClass

   Private Const BANK_COUNT As Integer = &H4%           'Defines the number of memory banks used.
   Private Const BANK_SIZE As Integer = &H2000%         'Defines the graphics mode's memory bank size.
   Private Const HEIGHT As Integer = 348                'Defines the graphics mode's height in pixels.
   Private Const PIXELS_PER_BYTE As Integer = &H8%      'Defines the number of pixels per byte.
   Private Const HORIZONTAL_SCALING As Integer = &H2%   'Defines the horizontal scale factor.
   Private Const VERTICAL_SCALING As Integer = &H2%     'Defines the vertical scale factor.
   Private Const WIDTH As Integer = 720                 'Defines the graphics mode's width in pixels.

   Private ReadOnly BLACK_BRUSH As New SolidBrush(Color.Black)   'Defines a black brush.
   Private ReadOnly WHITE_BRUSH As New SolidBrush(Color.White)   'Defines a white brush.

   'This procedure clears video adapter's buffer.
   Public Sub ClearBuffer() Implements VideoAdapterClass.ClearBuffer
      Dim Count As Integer = (VideoPageSizesE.Hercules720x348 \ &H2%)
      Dim Position As Integer = &H0%

      Do While Count > &H0%
         CPU.PutWord(AddressesE.HerculesBuffer + Position, &H0%)
         Count -= &H1%
         Position += &H2%
      Loop
   End Sub

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte, CodePage() As Integer) Implements VideoAdapterClass.Display
      Dim Bank As New Integer
      Dim Bit As New Integer
      Dim ByteOffset As New Integer
      Dim GraphicsO As Graphics = Nothing
      Dim PixelOff As New Boolean
      Dim Position As New Integer
      Dim RowOffset As New Integer

      Try
         GraphicsO = Graphics.FromImage(Screen)

         With GraphicsO
            For y As Integer = 0 To HEIGHT - 1
               Bank = y Mod BANK_COUNT
               RowOffset = (y \ BANK_COUNT) * HERCULES_720_348_BYTES_PER_ROW
               For x As Integer = 0 To WIDTH - 1
                  ByteOffset = x \ PIXELS_PER_BYTE
                  Position = AddressesE.HerculesBuffer + (Bank * BANK_SIZE) + RowOffset + ByteOffset
                  Bit = &H7% - (x Mod PIXELS_PER_BYTE)
                  PixelOff = ((CPU.Memory(Position) And (&H1% << Bit)) = &H0%)
                  .FillRectangle(If(PixelOff, BLACK_BRUSH, WHITE_BRUSH), x * HORIZONTAL_SCALING, y * VERTICAL_SCALING, HORIZONTAL_SCALING, VERTICAL_SCALING)
               Next x
            Next y
         End With
      Catch
      Finally
         If GraphicsO IsNot Nothing Then GraphicsO.Dispose()
      End Try
   End Sub

   'This procedure is ignored.
   Public Sub DrawCharacter(Index As Integer, Attribute As Integer) Implements VideoAdapterClass.DrawCharacter
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

   'This procedure is ignored.
   Public Sub ScrollBuffer(Up As Boolean, ScrollArea As VideoAdapterClass.ScreenAreaStr, Count As Integer) Implements VideoAdapterClass.ScrollBuffer
   End Sub
End Class
