'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Generic
Imports System.Drawing

'This class contains the EGA related procedures.
Public Class EGAClass
   'This enumeration lists the supported EGA video memory sizes.
   Public Enum MemorySizesE As Byte
      k64   '64 kb.
      k128  '128 kb.
      k192  '192 kb.
      k256  '256 kb.
   End Enum

   Public Const EGA_FEATURE_SWITCH_BITS As Integer = &HC%            'Defines the feature and switch bits.
   Public Const EGA_MEMORY_SIZE As MemorySizesE = MemorySizesE.k256  'Defines the reported video memory size.
   Public Const MONO_MODE As Boolean = False                         'Indicates whether or not mono mode is enabled.

   Public EGABrushes As New List(Of SolidBrush)   'Contains the brushes created using the current palette.

   'This procedure returns the 6-bit RGB code as a 24-bit RGB color.
   Private Function EGA6RGBBitTo24RGB(RGB6Bit As Byte) As Color
      Dim B1 As Integer = (RGB6Bit And &H1%)
      Dim B2 As Integer = (RGB6Bit And &H8%) >> &H3%
      Dim Blue As Integer = (B2 * &H40%) + (B1 * &HBF%)
      Dim G1 As Integer = (RGB6Bit And &H2%) >> &H1%
      Dim G2 As Integer = (RGB6Bit And &H10%) >> &H4%
      Dim Green As Integer = (G2 * &H40%) + (G1 * &HBF%)
      Dim R1 As Integer = (RGB6Bit And &H4%) >> &H2%
      Dim R2 As Integer = (RGB6Bit And &H20%) >> &H5%
      Dim Red As Integer = (R2 * &H40%) + (R1 * &HBF%)

      Return Color.FromArgb(Red, Green, Blue)
   End Function

   'This procedure updates the current EGA palette using the specified 6-bit RGB codes.
   Public Sub SetEntirePalette(Palette() As Byte)
      EGA.EGABrushes = New List(Of SolidBrush)
      For Each RGB6Bit As Byte In Palette
         EGA.EGABrushes.Add(New SolidBrush(EGA6RGBBitTo24RGB(RGB6Bit)))
      Next RGB6Bit
   End Sub
End Class
