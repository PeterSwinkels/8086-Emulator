'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

'This class contains the Real Time Clock.
Public Class RTCClass

   Private SelectedRegister As Integer = &H0%   'Contains the selected register.

   'This enumeration lists the supported RTC registers.
   Private Enum RegistersE As Integer
      LSBOfExtendedMemorSize = &H30%   'LSB of extended memory size found above 1 megabyte during POST.
   End Enum

   'This procedure selects the specified register.
   Public Sub SelectRegister(NewRegister As Integer)
      SelectedRegister = NewRegister
   End Sub

   'This procedure the selected register's contents.
   Public Function ReadRegister() As Byte
      Dim Value As New Byte

      Select Case SelectedRegister
         Case RegistersE.LSBOfExtendedMemorSize
            Value = &H0%
      End Select

      Return Value
   End Function
End Class
