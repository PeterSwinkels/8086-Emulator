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

   Public Const INTENSITY_BIT As Integer = &H10%   'Defines the color intensity bit.
   Private Const FRAME_DURATION As Integer = 20     'Defines the frame duration in milliseconds.
   Private Const RETRACE_DURATION As Integer = 2    'Defines the retrace duration in milliseconds.

   Public ReadOnly GET_SELECTED_REGISTER As Func(Of RegistersE) = Function() SelectedRegister       'Returns the selected register.
   Private ReadOnly CYCLE_CLOCK As Stopwatch = Stopwatch.StartNew()                                  'Defines the clock used to control refresh cycles.
   Private ReadOnly PALETTE0L() As Color = {Color.Black, Color.Green, Color.DarkRed, Color.Brown}    'Contains palette #0 low intensity colors.
   Private ReadOnly PALETTE1L() As Color = {Color.Black, Color.DarkCyan, Color.Purple, Color.Gray}   'Contains palette #1 low intensity colors.
   Private ReadOnly PALETTE0H() As Color = {Color.Black, Color.LightGreen, Color.Red, Color.Yellow}  'Contains palette #0 high intensity colors.
   Private ReadOnly PALETTE1H() As Color = {Color.Black, Color.Cyan, Color.Magenta, Color.White}     'Contains palette #1 high intensity colors.

   Public ActivePalette(&H0% To &H3%) As Color                                                    'Contains the active palette.
   Public IsMDA As Boolean = True                                                                 'Indicates whether or not an MDA will be emulated.
   Public PaintBrushes(&H0% To &H3%) As Brush                                                     'Contains the paint brushes created using the active palette.
   Public SelectedRegister As New RegistersE                                                      'Contains the selected register.
   Private IntenseColors As Boolean = True                                                         'Indicates whether or not intense colors are turned on.
   Private RegisterValue(RegistersE.HorizontalTotalCharacters To RegistersE.LightPenLSB) As Byte   'Contains the register values.
   Private SelectedPalette As New Integer                                                          'Contains the palette number used by the active palette.

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

   'This procedure creates the brushes based on the active palette.
   Private Sub CreateBrushes()
      For Index As Integer = ActivePalette.GetLowerBound(0) To ActivePalette.GetUpperBound(0)
         PaintBrushes(Index) = New SolidBrush(ActivePalette(Index))
      Next Index
   End Sub

   'This procedure returns the monochrome display adapter status port's current value.
   Public Function MDAStatus() As Byte
      Return CByte(If(CYCLE_CLOCK.ElapsedMilliseconds Mod FRAME_DURATION >= (FRAME_DURATION - RETRACE_DURATION), &H80%, &H2%))
   End Function

   'This procedure reads/writes to a register and returns the result.
   Public Function Register(Optional NewValue As Byte? = Nothing) As Byte
      If NewValue IsNot Nothing Then
         RegisterValue(SelectedRegister) = NewValue.Value
      End If

      Return RegisterValue(SelectedRegister)
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
End Class
