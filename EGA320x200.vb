'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Drawing

'This class emulates EGA 320x200.
Public Class EGA320x200Class
   Implements VideoAdapterClass

   Private Const HEIGHT As Integer = 200     'Defines the graphic mode's height in pixels.
   Private Const WIDTH As Integer = 320      'Defines the graphic mode's width in pixels.

   'This procedure clears video adapter's buffer.
   Public Sub ClearBuffer() Implements VideoAdapterClass.ClearBuffer
      Dim Count As Integer = VideoPageSizesE.EGA320x200 \ &H2%
      Dim Position As Integer = AddressesE.EGABuffer

      Do While Count > &H0%
         Memory.PutWord(Position, &H0%)
         Count -= &H1%
         Position += &H2%
      Loop
   End Sub

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte) Implements VideoAdapterClass.Display
      Dim GraphicsO As Graphics = Nothing

      Try
         GraphicsO = Graphics.FromImage(Screen)

         With GraphicsO
         End With
      Catch
      Finally
         If GraphicsO IsNot Nothing Then GraphicsO.Dispose()
      End Try
   End Sub

   'This procedure draws the specified character.
   Public Sub DrawCharacter(Index As Integer, Attribute As Integer) Implements VideoAdapterClass.DrawCharacter
   End Sub

   'This procedure initializes the video adapter.
   Public Sub Initialize() Implements VideoAdapterClass.Initialize
      ClearBuffer()

      Memory(AddressesE.VideoPage) = &H0%
      ResetCursor()
      CursorBlink.Enabled = False
   End Sub

   'This procedure returns the screen size used by a video adapter.
   Public Function Resolution() As Size Implements VideoAdapterClass.Resolution
      Return New Size(WIDTH * MCC.Scaling, HEIGHT * MCC.Scaling)
   End Function

   'This procedure scrolls the video adapter's buffer.
   Public Sub ScrollBuffer(Up As Boolean, ScrollArea As VideoAdapterClass.ScreenAreaStr, Count As Integer) Implements VideoAdapterClass.ScrollBuffer
   End Sub
End Class
