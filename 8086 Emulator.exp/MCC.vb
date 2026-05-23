'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Diagnostics
Imports System.Drawing

'This class contains the 6845 Motorola CRT Controller's related procedures.
Public Class MCCClass
   'This enumeration lists the MCC's registers.
   Public Enum RegistersE As Integer
      HorizontalTotalCharacters     'Horizontal total characters.
      HorizontalCharactersPerLine   'Horizontal displayed characters per line.
      HorizontalSyncPosition        'Horizontal synchronization position.
      HorizontalSyncWidth           'Horizontal synchronization width.
      VerticalTotalLines            'Vertical total lines.
      VerticalTotalAdjust           'Vertical total adjust.
      VerticalDisplayedRows         'Vertical displayed rows.
      VerticalSyncPosition          'Vertical synchronization position.
      InterlaceMode                 'Interlace mode.
      MaximumScanLineAddress        'Maximum scan line address.
      CursorStart                   'Cursor start.
      CursorEnd                     'Cursor end.
      StartAddressMSB               'Start address MSB.
      StartAddressLSB               'Start address LSB.
      CursorAddressMSB              'Cursor address MSB.
      CursorAddressLSB              'Cursor address LSB.
      LightPenMSB                   'Light pen MSB.
      LightPenLSB                   'Light pen LSB.
   End Enum

   'This enumeration lists the scanline counts.
   Public Enum RasterScanLinesE As Byte
      SC200   '200 scanlines.
      SC350   '350 scanlines.
      SC400   '400 scanlines.
      SC480   '480 scanlines.
   End Enum

   Public Const INTENSITY_BIT As Integer = &H10%              'Defines the color intensity bit.
   Public Const VIDEO_RAM_256KB As Byte = &H3%                 'Defines the value for 256 kb of video RAM.
   Private Const CGA_RETRACE_DURATION As Integer = 2           'Defines the CGA retrace duration in milliseconds.
   Private Const CHARACTER_SCANLINE_COUNT As Byte = &HE%       'Defines the number of scanlines per character.
   Private Const DEFAULT_CHARACTER_ROW_COUNT As Byte = &H19%   'Defines the default number of character rows.
   Private Const FRAME_DURATION As Integer = 20                'Defines the frame duration in milliseconds.
   Private Const MDA_RETRACE_DURATION As Integer = 20          'Defines the MDA retrace duration in milliseconds.

   Public ReadOnly BACKGROUND_COLORS() As Color = {Color.Black, Color.DarkBlue, Color.DarkGreen, Color.DarkCyan, Color.DarkRed, Color.Purple, Color.Brown, Color.White, Color.Gray, Color.Blue, Color.Green, Color.Cyan, Color.Red, Color.Pink, Color.Yellow, Color.White}   'Contains the background colors.
   Public ReadOnly GET_SELECTED_REGISTER As Func(Of RegistersE) = Function() SelectedRegister       'Returns the selected register.
   Private ReadOnly CYCLE_CLOCK As Stopwatch = Stopwatch.StartNew()                                  'Defines the clock used to control refresh cycles.
   Private ReadOnly PALETTE0L() As Color = {Color.Black, Color.Green, Color.DarkRed, Color.Brown}    'Contains palette #0 low intensity colors.
   Private ReadOnly PALETTE1L() As Color = {Color.Black, Color.DarkCyan, Color.Purple, Color.Gray}   'Contains palette #1 low intensity colors.
   Private ReadOnly PALETTE0H() As Color = {Color.Black, Color.LightGreen, Color.Red, Color.Yellow}  'Contains palette #0 high intensity colors.
   Private ReadOnly PALETTE1H() As Color = {Color.Black, Color.Cyan, Color.Magenta, Color.White}     'Contains palette #1 high intensity colors.

   Public ActivePalette(&H0% To &H3%) As Color                                                    'Contains the active palette.
   Public BlinkingOn As Boolean = True                                                            'Incidates whether or not blinking colors are enabled.
   Public IsMDA As Boolean = True                                                                 'Indicates whether or not an MDA will be emulated.
   Public PaintBrushes(&H0% To &H3%) As Brush                                                     'Contains the paint brushes created using the active palette.
   Public SelectedRegister As New RegistersE                                                      'Contains the selected register.
   Private IntenseColors As Boolean = True                                                         'Indicates whether or not intense colors are turned on.
   Private RegisterValue(RegistersE.HorizontalTotalCharacters To RegistersE.LightPenLSB) As Byte   'Contains the register values.
   Private SelectedPalette As New Integer                                                          'Contains the palette number used by the active palette.

   'This function returns the number of scanlines per character.
   Public Function CharacterScanLineCount() As Byte
      Return CHARACTER_SCANLINE_COUNT
   End Function

   'This procedure returns the number of colors for the current video mode.
   Public Function ColorCount() As Integer
      Dim Count As New Integer

      Select Case CurrentVideoMode
         Case VideoModesE.CGA640x200,
              VideoModesE.EGA640x350Mono,
              VideoModesE.Text40x25Mono,
              VideoModesE.Text80x25Mono,
              VideoModesE.VGA640x480Mono
            Count = &H2%
         Case VideoModesE.CGA320x200A,
              VideoModesE.CGA320x200B,
              VideoModesE.PCjr640x200
            Count = &H4%
         Case VideoModesE.EGA320x200,
              VideoModesE.EGA640x200,
              VideoModesE.EGA640x350,
              VideoModesE.PCjr160x200,
              VideoModesE.PCjr320x200,
              VideoModesE.Text40x25Color,
              VideoModesE.Text80x25Color,
              VideoModesE.Text80x25Gray,
              VideoModesE.VGA640x480
            Count = &H10%
         Case VideoModesE.VGA320x200
            Count = &H100%
      End Select

      Return Count
   End Function

   'This procedure selects the active the palette.
   Public Sub ColorSelect(Value As Integer)
      Select Case Value
         Case &H0%
            SelectedPalette = Value
            ActivePalette = If(IntenseColors, PALETTE0H, PALETTE0L)
            CreateBrushes()
         Case &H1%
            SelectedPalette = Value
            ActivePalette = If(IntenseColors, PALETTE1H, PALETTE1L)
            CreateBrushes()
      End Select
   End Sub

   'This procedure returns the number of character columns for the current video mode.
   Public Function ColumnCount() As Byte
      Dim Count As Byte = Nothing

      Select Case CurrentVideoMode
         Case VideoModesE.CGA320x200A,
              VideoModesE.CGA320x200B,
              VideoModesE.EGA320x200,
              VideoModesE.PCjr160x200,
              VideoModesE.PCjr320x200,
              VideoModesE.Text40x25Color,
              VideoModesE.Text40x25Mono,
              VideoModesE.VGA320x200
            Count = &H28%
         Case VideoModesE.CGA640x200,
              VideoModesE.EGA640x200,
              VideoModesE.EGA640x350,
              VideoModesE.EGA640x350Mono,
              VideoModesE.PCjr640x200,
              VideoModesE.Text80x25Color,
              VideoModesE.Text80x25Gray,
              VideoModesE.Text80x25Mono,
              VideoModesE.VGA640x480,
              VideoModesE.VGA640x480Mono
            Count = &H50%
      End Select

      Return Count
   End Function

   'This procedure creates the brushes based on the active palette.
   Private Sub CreateBrushes()
      For Index As Integer = ActivePalette.GetLowerBound(0) To ActivePalette.GetUpperBound(0)
         PaintBrushes(Index) = New SolidBrush(ActivePalette(Index))
      Next Index
   End Sub

   'This procedure returns the monochrome display adapter status port's current value.
   Public Function MDAStatus() As Byte
      Return CByte(If(CYCLE_CLOCK.ElapsedMilliseconds Mod FRAME_DURATION >= (FRAME_DURATION - MDA_RETRACE_DURATION), &H80%, &H2%))
   End Function

   'This procedure returns the number of raster scanlines for the current video mode.
   Public Function RasterScanLineCount() As RasterScanLinesE
      Dim Count As New RasterScanLinesE

      Select Case CurrentVideoMode
         Case VideoModesE.CGA320x200A,
              VideoModesE.CGA320x200B,
              VideoModesE.EGA320x200,
              VideoModesE.EGA640x200,
              VideoModesE.PCjr160x200,
              VideoModesE.PCjr320x200,
              VideoModesE.PCjr640x200,
              VideoModesE.VGA320x200
            Count = RasterScanLinesE.SC200
         Case VideoModesE.EGA640x350,
              VideoModesE.Text40x25Color,
              VideoModesE.Text40x25Mono,
              VideoModesE.Text80x25Color,
              VideoModesE.Text80x25Gray,
              VideoModesE.Text80x25Mono
            Count = RasterScanLinesE.SC350
         Case VideoModesE.VGA640x480,
              VideoModesE.VGA640x480Mono
            Count = RasterScanLinesE.SC480
         Case Else
            Count = RasterScanLinesE.SC400
      End Select

      Return Count
   End Function

   'This procedure reads/writes to a register and returns the result.
   Public Function Register(Optional NewValue As Byte? = Nothing) As Byte
      If NewValue IsNot Nothing Then
         RegisterValue(SelectedRegister) = NewValue.Value
      End If

      Return RegisterValue(SelectedRegister)
   End Function

   'This procedure returns the number of character rows.
   Public Function RowCount() As Byte
      Return DEFAULT_CHARACTER_ROW_COUNT
   End Function

   'This procedure selects the specified color intensity.
   Public Sub SelectIntensity(NewIntensity As Boolean)
      IntenseColors = NewIntensity
      ColorSelect(SelectedPalette)
   End Sub

   'This procedure selects the specified register and returns whether or not the selection succeeded.
   Public Function SelectRegister(Register As RegistersE) As Boolean
      Dim Success As Boolean = True

      Select Case Register
         Case RegistersE.CursorAddressMSB, RegistersE.CursorAddressLSB
            SelectedRegister = Register
         Case Else
            Success = False
      End Select

      Return Success
   End Function

   'This procedure returns the color graphics display adapter status port's current value.
   Public Function CGAStatus() As Byte
      Return CByte(If(CYCLE_CLOCK.ElapsedMilliseconds Mod FRAME_DURATION >= (FRAME_DURATION - CGA_RETRACE_DURATION), &H8%, &H1%))
   End Function

   'This procedure returns the number of video pages for the current video mode.
   Public Function VideoPageCount() As Byte
      Dim Count As Byte = Nothing

      Select Case CurrentVideoMode
         Case VideoModesE.EGA320x200,
              VideoModesE.Text40x25Color,
              VideoModesE.Text40x25Mono,
              VideoModesE.Text80x25Color
            Count = &H8%
         Case VideoModesE.EGA640x200,
              VideoModesE.Text80x25Gray
            Count = &H4%
         Case VideoModesE.EGA640x350Mono,
              VideoModesE.VGA640x480Mono
            Count = &H2%
         Case Else
            Count = &H1%
      End Select

      Return Count
   End Function

End Class
