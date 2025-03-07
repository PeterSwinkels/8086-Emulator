'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Drawing

'This class emulates CGA 320x200.
Public Class CGA320x200Class
   Implements VideoAdapterClass

   Private Const BYTES_PER_ROW As Integer = &H50%    'Defines the number of bytes per pixel row.
   Private Const PIXELS_PER_BYTE As Integer = &H4%   'Defines the number of pixels per byte.

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte, CodePage() As Integer) Implements VideoAdapterClass.Display
      Dim Index As New Integer
      Dim Position As New Integer

      With DirectCast(Screen, Bitmap)
         For y2 As Integer = 0 To 1
            Position = AddressesE.CGA320x200 + (y2 * BYTES_PER_ROW)
            For y1 As Integer = 0 To ScreenSize.Height - 1 Step 2
               For x As Integer = 0 To ScreenSize.Width - 1 Step PIXELS_PER_BYTE
                  For Pixel As Integer = &H0% To PIXELS_PER_BYTE - &H1%
                     Index = ((CPU.Memory(Position) And (&H3% << (Pixel * &H2%))) >> (Pixel * &H2%))
                     .SetPixel(x + (PIXELS_PER_BYTE - &H1%) - Pixel, y1 + y2, If(Index = 0, BackgroundColor, Palette(Index)))
                  Next Pixel
                  Position += &H1%
               Next x
            Next y1
         Next y2
      End With
   End Sub

   'This procedure initializes the video adapter.
   Public Sub Initialize() Implements VideoAdapterClass.Initialize
      Dim Count As New Integer
      Dim Position As New Integer

      Count = (CGA_320_x_200_BUFFER_SIZE \ &H2%)
      Position = AddressesE.CGA320x200
      Do While Count > &H0%
         CPU.PutWord(Position, &H0%)
         Count -= &H1%
         Position += &H2%
      Loop

      CPU.Memory(AddressesE.VideoPage) = &H0%
      CPU.PutWord(AddressesE.CursorPositions, Word:=&H0%)
      CPU.PutWord(AddressesE.CursorScanLines, Word:=CURSOR_DEFAULT)
      CursorBlink.Enabled = False

      Palette = PALETTE1
   End Sub

   'This procedure returns the screen size used by a video adapter.
   Public Function ScreenSize() As Size Implements VideoAdapterClass.ScreenSize
      Return New Size(320, 200)
   End Function
End Class
