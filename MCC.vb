'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Diagnostics

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

   Private Const FRAME_DURATION As Integer = 20    'Defines the frame duration in milliseconds.
   Private Const RETRACE_DURATION As Integer = 2   'Defines the retrace duration in milliseconds.
   Public Const IS_MDA As Boolean = True            'Indicates whether or not an MDA will be emulated.

   Private ReadOnly CYCLE_CLOCK As Stopwatch = Stopwatch.StartNew()   'Defines the clock used to control refresh cycles.   

   Public ReadOnly GET_SELECTED_REGISTER As Func(Of RegistersE) = Function() SelectedRegister     'Returns the selected register.
   Private RegisterValue(RegistersE.HorizontalTotalCharacters To RegistersE.LightPenLSB) As Byte   'Contains the register values.
   Private SelectedRegister As New RegistersE                                                      'Contains the selected register.


   'This procedure selects the specified color palette.
   Public Sub ColorSelect(Value As Integer)
      Select Case Value
         Case &H0%
            Palette = PALETTE0
         Case &H20%
            Palette = PALETTE1
      End Select
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
