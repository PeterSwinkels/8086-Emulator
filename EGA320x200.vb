'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq

'This class emulates EGA 320x200.
Public Class EGA320x200Class
   Implements VideoAdapterClass

   Private Const BIT_PLANE_SIZE As Integer = &H1F40%   'Defines the graphic mode's bit-plane size.
   Private Const HEIGHT As Integer = 200               'Defines the graphic mode's height in pixels.
   Private Const WIDTH As Integer = 320                'Defines the graphic mode's width in pixels.

   'This procedure clears video adapter's buffer.
   Public Sub ClearBuffer() Implements VideoAdapterClass.ClearBuffer
      VGA.BitPlanes = New List(Of Byte())

      For Each BitPlane As Integer In [Enum].GetValues(GetType(VGAClass.BitPlanesE))
         VGA.BitPlanes.Add(Enumerable.Repeat(CByte(&H0%), BIT_PLANE_SIZE).ToArray())
      Next BitPlane
   End Sub

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte) Implements VideoAdapterClass.Display
      Dim Color As New Integer
      Dim GraphicsO As Graphics = Nothing
      Dim Planes As Integer = VGA.SequencerData(VGAClass.SequencerRegistersE.SelectedOutputBitPlanes)
      Dim Position As New Integer
      Dim SelectionBit As New Integer

      Try
         GraphicsO = Graphics.FromImage(Screen)

         For Each BitPlane As Integer In [Enum].GetValues(GetType(VGAClass.BitPlanesE))
            SelectionBit = &H1% << BitPlane
            If (Planes And SelectionBit) = SelectionBit Then
               VGA.BitPlanes(BitPlane) = Emulator8086Program.Memory.GetRange(AddressesE.EGABuffer, BIT_PLANE_SIZE)
            End If
         Next BitPlane

         With GraphicsO
            Position = &H0%
            For y As Integer = 0 To HEIGHT - 1
               For x As Integer = 0 To WIDTH - 1 Step 8
                  For Bit As Integer = &H0% To &H7%
                     Color = &H0%
                     For Each BitPlane As Integer In [Enum].GetValues(GetType(VGAClass.BitPlanesE))
                        If (VGA.BitPlanes(BitPlane)(Position) And (&H1% << Bit)) = (&H1% << Bit) Then
                           Color = Color Or (&H1% << BitPlane)
                        End If
                     Next BitPlane

                     .FillRectangle(EGA.Brushes(Color), (x + Bit) * MCC.Scaling, y * MCC.Scaling, MCC.Scaling, MCC.Scaling)
                  Next Bit
                  Position += &H1%
               Next x
            Next y
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
