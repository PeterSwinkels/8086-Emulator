'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Timers

'This class contains the 8253 PIT.
Public Class PITClass
   'This enumeration lists the read/write format bits.
   Private Enum FormatsE As Integer
      None     'Latches a counter.
      MSB      'Read/write of MSB only.
      LSB      'Read/write of LSB only.
      LSBMSB   'Read/write of LSB followed by MSB.
   End Enum

   Public Enum ModesE As Integer
      Mode0   'Interrupt on terminal count. (one-shot)
      Mode1   'Hardware re-triggerable. (one-shot)
      Mode2   'Rate generator. (periodic interrupts, used for timer)
      Mode3   'Square wave generator.
      Mode4   'Software triggered strobe.
      Mode5   'Hardware triggered strobe.
   End Enum

   'This structure defines a counter and its settings.
   Private Structure CounterStr
      Public BCD As Boolean       'Defines whether or not a counter is in BCD mode.
      Public Format As FormatsE   'Defines a counter's format.
      Public Mode As ModesE       'Defines a counter's mode.
      Public MSB As Byte?         'Defines a counter's most significant byte.
      Public Latched As Boolean   'Indicates whether or not a counter is latched.
      Public LSB As Byte?         'Defines a counter's least significant byte.
      Public LSBRead As Boolean   'Indicates whether or not a counter's lsb has been read.
      Public Reload As Integer    'Defines a counter's reload value.
      Public [Timer] As Timer     'Defines a counter's timer.
      Public Value As Integer     'Defines a counter's value.
   End Structure

   Private Const BINARY_BCD_MASK As Integer = &H1%  'Defines the binary/BCD bit.
   Private Const COUNTER_MASK As Integer = &HC0%    'Defines the counter bits.
   Private Const FORMAT_MASK As Integer = &H30%     'Defines the format bits.
   Private Const FREQUENCY As Double = 1193182.0    'Defines the PIT's clock frequency.
   Private Const MODE_MASK As Integer = &HE%        'Defines the mode bits.

   Private ReadOnly Counters(&H0% To &H2%) As CounterStr  'Contains the counters.

   Public Event PITEvent(Counter As Integer, Mode As ModesE)  'Defines an event triggered by the 8253 PIT.

   'This procedure initializes this class.
   Public Sub New()
      For Counter As Integer = Counters.GetLowerBound(0) To Counters.GetUpperBound(0)
         Counters(Counter).[Timer] = New Timer() With {.AutoReset = True, .Enabled = False, .Interval = FREQUENCY}
         AddHandler Counters(Counter).[Timer].Elapsed, AddressOf UpdateCounters
      Next Counter
   End Sub

   'This procedure caps the specified value's octeds, if necessary, to ensure the value is in BCD and returns the result.
   Private Function BCDCap(Value As Integer) As Integer
      Dim BCD As Integer = 0
      Dim Octed As New Integer

      Do
         Octed = Value And &HF%
         BCD = BCD Or If(Octed < &HA%, Octed, &H9%)
         Value = Value >> &H4%
         If Value = &H0% Then Exit Do
         BCD = BCD << &H4%
      Loop

      Return BCD
   End Function

   'This procedure initializes the specified counter and starts it depending on its mode.
   Private Sub InitializeCounter(Counter As Integer, NewValue As Integer)
      With Counters(Counter)
         If .BCD Then NewValue = CByte(BCDCap(NewValue))
         If .Mode = ModesE.Mode3 Then NewValue = NewValue And &HFFFE%

         If .[Timer].Enabled Then
            Select Case .Mode
               Case ModesE.Mode0
                  .[Timer].Stop()
                  .Value = NewValue
                  .Reload = .Value
                  StartCounter(Counter)
               Case Else
                  .Reload = .Value
            End Select
         Else
            .Value = NewValue
            .Reload = .Value

            Select Case .Mode
               Case ModesE.Mode0, ModesE.Mode2, ModesE.Mode3, ModesE.Mode4
                  StartCounter(Counter)
            End Select
         End If
      End With
   End Sub

   'This procedure emulates the mode control register.
   Public Sub ModeControl(NewValue As Integer)
      Dim Format As FormatsE = DirectCast((NewValue And FORMAT_MASK) >> &H4%, FormatsE)

      With Counters((NewValue And COUNTER_MASK) >> &H6%)
         .BCD = ((NewValue And BINARY_BCD_MASK) = &H1%)
         .Mode = DirectCast((NewValue And MODE_MASK) >> &H1%, ModesE)

         Select Case Format
            Case FormatsE.MSB To FormatsE.LSBMSB
               .Format = Format
            Case FormatsE.None
               .Latched = True
               .LSB = CByte(.Value And &HFF%)
               .MSB = CByte(.Value >> &H8%)
         End Select
      End With
   End Sub

   'This procedure reads from the specified counter and returns the result.
   Public Function ReadCounter(Counter As Integer) As Byte
      Dim Value As Byte = Nothing

      With Counters(Counter)
         If .Latched Then
            .Latched = False
            Select Case .Format
               Case FormatsE.LSB
                  .LSBRead = False
                  Value = CByte(.LSB)
               Case FormatsE.LSBMSB
                  If .LSBRead Then
                     .LSBRead = False
                     Value = CByte(.LSB)
                  Else
                     .LSBRead = True
                     Value = CByte(.MSB)
                  End If
               Case FormatsE.MSB
                  .LSBRead = False
                  Value = CByte(.MSB)
            End Select
         Else
            Select Case .Format
               Case FormatsE.LSB
                  .LSBRead = False
                  Value = CByte(.Value And &HFF%)
               Case FormatsE.LSBMSB
                  If .LSBRead Then
                     .LSBRead = False
                     Value = CByte(.Value >> &H8%)
                  Else
                     .LSBRead = True
                     Value = CByte(.Value And &HFF%)
                  End If
               Case FormatsE.MSB
                  .LSBRead = False
                  Value = CByte(.Value >> &H8%)
            End Select
         End If
      End With

      Return Value
   End Function

   'This procedure starts the specified counter.
   Public Sub StartCounter(Counter As Integer)
      With Counters(Counter)
         If .Value > &H0% Then
            .[Timer].Interval = (.Value / FREQUENCY) * 1000
            .[Timer].Enabled = True
         End If
      End With
   End Sub

   'This procedure updates the counters.
   Private Sub UpdateCounters(sender As Object, e As ElapsedEventArgs)
      For Counter As Integer = Counters.GetLowerBound(0) To Counters.GetUpperBound(0)
         With Counters(Counter)
            Select Case .Value
               Case &H0%
                  RaiseEvent PITEvent(Counter, .Mode)
                  Select Case .Mode
                     Case ModesE.Mode2, ModesE.Mode3
                        .Value = .Reload
                     Case Else
                        .[Timer].Stop()
                  End Select
               Case Else
                  If .BCD Then
                     .Value -= If(.Mode = ModesE.Mode3, If((.Value And &HF%) = &H0%, &H8%, &H2%), If((.Value And &HF%) = &H0%, &H7%, &H1%))
                  Else
                     .Value -= If(.Mode = ModesE.Mode3, &H2%, &H1%)
                  End If
            End Select
         End With
      Next Counter
   End Sub

   'This procedure writes to the specified counter.
   Public Sub WriteCounter(Counter As Integer, NewValue As Byte)
      With Counters(Counter)
         Select Case .Format
            Case FormatsE.LSB
               .LSB = NewValue
               .MSB = Nothing
               InitializeCounter(Counter, NewValue:=CInt(.LSB))
            Case FormatsE.LSBMSB
               If .LSB Is Nothing Then
                  .LSB = NewValue
               Else
                  .MSB = NewValue
                  InitializeCounter(Counter, NewValue:=(CInt(.MSB) << &H8%) Or CInt(.LSB))
               End If
            Case FormatsE.MSB
               .LSB = Nothing
               .MSB = NewValue
               InitializeCounter(Counter, NewValue:=(CInt(.MSB) << &H8%) Or CInt(.MSB) << &H8%)
         End Select
      End With
   End Sub
End Class
