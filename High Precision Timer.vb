'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks

'This class contains the high precision timer.
Public Class HighPrecisionTimerClass
   <DllImport("kernel32.dll")> Private Shared Function QueryPerformanceCounter(ByRef lpPerformanceCount As Long) As Boolean
   End Function
   <DllImport("Kernel32.dll")> Private Shared Function QueryPerformanceFrequency(ByRef lpPerformanceFreq As Long) As Boolean
   End Function

   Public Clock As New Task(AddressOf Tick)           'Contains the timer's clock.
   Public ClockToken As New CancellationTokenSource   'Indicates whether or not to stop the timer.
   Private IntervalTickCount As New Double             'Contains the number of ticks between intervals.
   Private LastIntervalCount As New Long               'Contains the number of ticks performed at the time of the last interval.

   Public Event IntervalElapsed()   'Contains the interval elapsed event.

   'This procedure initializes this timer.
   Public Sub New(IntervalDuration As Double)
      Dim Frequency As New Long

      QueryPerformanceFrequency(Frequency)
      IntervalTickCount = Frequency * IntervalDuration
   End Sub

   'This procedure is executed each time the timer ticks.
   Public Sub Tick()
      Dim Count As New Long

      QueryPerformanceCounter(LastIntervalCount)

      Do Until ClockToken.Token.IsCancellationRequested
         QueryPerformanceCounter(Count)
         If Count - LastIntervalCount >= IntervalTickCount Then
            LastIntervalCount = Count
            RaiseEvent IntervalElapsed()
         End If
      Loop
   End Sub
End Class
