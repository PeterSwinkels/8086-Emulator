'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Diagnostics


'This class contains the 6845 Motorola CRT Controller's related procedures.
Public Class MCCClass
   Private Const FRAME_DURATION As Integer = 20    'Defines the frame duration in milliseconds.
   Private Const RETRACE_DURATION As Integer = 2   'Defines the retrace duration in milliseconds.

   Private ReadOnly CYCLE_CLOCK As Stopwatch = Stopwatch.StartNew()   'Defines the clock used to control refresh cycles.   

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
   Public Function Status() As Byte
      Return CByte(If(CYCLE_CLOCK.ElapsedMilliseconds Mod FRAME_DURATION >= (FRAME_DURATION - RETRACE_DURATION), &H80%, &H2%))
   End Function
End Class
