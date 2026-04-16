'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Threading.Tasks

'This class contains the 8253 PIT.
Public Class PITClass
   'This enumeration lists the counters.
   Public Enum CountersE As Integer
      TimeOfDay            'Time of day clock.
      RAMRefresh           'RAM refresh.
      CassetteAndSpeaker   'Cassette and speaker.
   End Enum

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
      Public BCD As Boolean              'Defines whether or not a counter is in BCD mode.
      Public Format As FormatsE          'Defines a counter's format.
      Public Mode As ModesE              'Defines a counter's mode.
      Public Mode3Half As Boolean        'Defines ... ???
      Public MSB As Byte?                'Defines a counter's most significant byte.
      Public Latched As Boolean          'Indicates whether or not a counter is latched.
      Public LatchedValue As Integer     'Contains the latched counter.
      Public LatchedLSBRead As Boolean   'Indicates LSB/MSB order for latch.
      Public LSB As Byte?                'Defines a counter's least significant byte.
      Public LSBRead As Boolean          'Indicates whether or not a counter's lsb has been read.
      Public Reload As Integer           'Defines a counter's reload value.
      Public Value As Integer            'Defines a counter's value.
   End Structure

   Private Const BINARY_BCD_MASK As Integer = &H1%   'Defines the binary/BCD bit.
   Private Const COUNTER_MASK As Integer = &HC0%     'Defines the counter bits.
   Private Const FORMAT_MASK As Integer = &H30%      'Defines the format bits.
   Private Const FREQUENCY As Double = 1.19318       'Defines the PIT's clock frequency in MHz.
   Private Const MODE_MASK As Integer = &HE%         'Defines the mode bits.

   Private ReadOnly Counters(&H0% To &H2%) As CounterStr  'Contains the counters.

   Public WithEvents HighPrecisionTimer As HighPrecisionTimerClass = Nothing   'Contains the PIT's timer.

   'This procedure initializes this class.
   Public Sub New()
      HighPrecisionTimer = New HighPrecisionTimerClass(IntervalDuration:=FREQUENCY / 1000000)

      With Counters(CountersE.TimeOfDay)
         .BCD = False
         .Format = FormatsE.LSBMSB
         .Latched = False
         .Mode = ModesE.Mode3
         .Reload = &HFFFF%
         .Value = .Reload
      End With

      HighPrecisionTimer.Clock.Start()
   End Sub

   'This procedure caps the specified value's octets, if necessary, to ensure the value is in BCD and returns the result.
   Private Function BCDCap(Value As Integer) As Integer
      Dim BCD As Integer = 0
      Dim Octet As New Integer

      Do
         Octet = Value And &HF%
         BCD = BCD Or If(Octet < &HA%, Octet, &H9%)
         Value = Value >> &H4%
         If Value = &H0% Then Exit Do
         BCD = BCD << &H4%
      Loop

      Return BCD
   End Function

   'This procedure initializes the specified counter and starts it depending on its mode.
   Private Sub InitializeCounter(Counter As CountersE, NewValue As Integer)
      With Counters(Counter)
         If .BCD Then NewValue = CByte(BCDCap(NewValue))

         Select Case .Mode
            Case ModesE.Mode0
               .Value = NewValue
               .Reload = .Value
            Case ModesE.Mode3
               NewValue = NewValue And &HFFFE

               .Reload = NewValue
               .Value = NewValue
               .Mode3Half = False
            Case Else
               .Reload = .Value
         End Select
      End With
   End Sub

   'This procedure emulates the mode control register.
   Public Sub ModeControl(NewValue As Integer)
      Dim CounterIndex As Integer = (NewValue And COUNTER_MASK) >> &H6%
      Dim Format As FormatsE = DirectCast((NewValue And FORMAT_MASK) >> &H4%, FormatsE)

      With Counters(CounterIndex)
         .BCD = ((NewValue And BINARY_BCD_MASK) = &H1%)

         If Format = FormatsE.None Then
            .Latched = True
            .LatchedValue = .Value
            .LatchedLSBRead = False
         Else
            .Mode = DirectCast((NewValue And MODE_MASK) >> &H1%, ModesE)
            .Format = Format
            .LSB = Nothing
            .MSB = Nothing
            .LSBRead = False
         End If
      End With
   End Sub

   'This procedure reads from the specified counter and returns the result.
   Public Function ReadCounter(Counter As CountersE) As Byte
      Dim Value As Byte = Nothing

      With Counters(Counter)
         If .Latched Then
            Select Case .Format
               Case FormatsE.LSB
                  Value = CByte(.LatchedValue And &HFF%)
                  .Latched = False
               Case FormatsE.MSB
                  Value = CByte(.LatchedValue >> &H8%)
                  .Latched = False
               Case FormatsE.LSBMSB
                  If Not .LatchedLSBRead Then
                     Value = CByte(.LatchedValue And &HFF%)
                     .LatchedLSBRead = True
                  Else
                     Value = CByte(.LatchedValue >> &H8%)
                     .LatchedLSBRead = False
                     .Latched = False
                  End If
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

   'This procedure updates the counters.
   Private Sub UpdateCounters() Handles HighPrecisionTimer.IntervalElapsed
      For Each Counter As CountersE In [Enum].GetValues(GetType(CountersE))
         With Counters(Counter)
            If .Mode = ModesE.Mode3 Then
               .Value -= &H2%

               If .Value <= &H0% Then
                  If Not .Mode3Half Then
                     .Value = .Reload
                     .Mode3Half = True
                  Else
                     If CPU.Clock.Status = TaskStatus.Running AndAlso Counter = PITClass.CountersE.TimeOfDay Then
                        PIC.RaiseIRQ(&H0%)
                     End If
                     .Value = .Reload
                     .Mode3Half = False
                  End If
               End If

               Continue For
            End If

            If .Mode = ModesE.Mode2 Then
               .Value -= &H1%
               If .Value = &H1% Then
                  If CPU.Clock.Status = TaskStatus.Running AndAlso Counter = PITClass.CountersE.TimeOfDay Then
                     PIC.RaiseIRQ(&H0%)
                  End If
               End If

               If .Value <= &H0% Then
                  .Value = .Reload
               End If

               Continue For
            End If

            If .Mode = ModesE.Mode0 Then
               If .Value > &H0% Then
                  .Value -= &H1%
                  If .Value = &H0% Then
                     If CPU.Clock.Status = TaskStatus.Running AndAlso Counter = PITClass.CountersE.TimeOfDay Then
                        PIC.RaiseIRQ(&H0%)
                     End If
                  End If
               End If

               Continue For
            End If

            .Value -= &H1%
            If .Value <= &H0% Then
               If CPU.Clock.Status = TaskStatus.Running AndAlso Counter = PITClass.CountersE.TimeOfDay Then
                  PIC.RaiseIRQ(&H0%)
               End If
               If .Mode = ModesE.Mode2 OrElse .Mode = ModesE.Mode3 Then
                  .Value = .Reload
               Else
                  .Value = &H0%
               End If
            End If

         End With
      Next Counter
   End Sub

   'This procedure writes to the specified counter.
   Public Sub WriteCounter(Counter As CountersE, NewValue As Byte)
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
                  If Counter = CountersE.CassetteAndSpeaker Then
                     PC_SPEAKER.SetFrequency((CInt(.MSB.Value) << 8) Or .LSB.Value)
                  End If
               End If
            Case FormatsE.MSB
               .LSB = Nothing
               .MSB = NewValue
               InitializeCounter(Counter, NewValue:=(CInt(.MSB) << &H8%))
         End Select
      End With
   End Sub
End Class
