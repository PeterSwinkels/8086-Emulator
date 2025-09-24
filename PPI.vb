'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

'This class contains the 8255 Programmable Peripheral Interface's related procedures.
Public Class PPIClass
   'This procedure manages I/O port B and returns the last value written.
   Public Function PortB(Optional NewValue As Integer? = Nothing) As Integer
      Static CurrentValue As New Integer

      If NewValue IsNot Nothing Then CurrentValue = CInt(NewValue)

      Return CurrentValue
   End Function
End Class
