
'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

'This class contains the 8237 DMA controller.
Public Class DMAClass
   Private Addresses(&H0% To &H3%) As Integer   'Contains the channel addresses.
   Public WriteMSB As Boolean = False          'Indicates whether the LSB or MSB is selected to be written.

   'This procedure updates the specified channel with the specified address.
   Public Sub UpdateChannelAddress(Channel As Integer, Value As Byte)
      If WriteMSB Then
         Addresses(Channel) = (Addresses(Channel) And &HFF%) Or (CInt(Value) << &H8%)
      Else
         Addresses(Channel) = (Addresses(Channel) And &HFF00%) Or Value
      End If

      WriteMSB = Not WriteMSB
   End Sub
End Class
