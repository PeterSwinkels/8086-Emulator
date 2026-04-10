'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Convert

'This class contains the 8259 PIC.
Public Class PICClass
   Private Const NON_SPECIFIC_EOI As Byte = &H20%   'Defines a non-specific EOI.

   Private InterruptInProgress As New Boolean   'Indicates whether or not an interrupt is in progress.
   Private IRQMask As Byte = &HFF%              'Contains the IRQ mask.
   Private PendingIRQs As Byte = &H0%           'Contains the active IRQs.

   'This procedure returns the first pending unmasked IRQ.
   Private Function FindPendingUnmaskedIRQ() As Integer?
      Dim IRQFound As Integer? = Nothing
      Dim MaskBit As New Boolean
      Dim PendingBit As New Boolean

      For IRQ As Integer = &H0% To &H7%
         MaskBit = (IRQMask And (&H1% << IRQ)) <> &H0%
         PendingBit = (PendingIRQs And (&H1% << IRQ)) <> &H0%

         If PendingBit AndAlso Not MaskBit Then
            IRQFound = IRQ
            Exit For
         End If
      Next IRQ

      Return IRQFound
   End Function

   'This procedure returns an interrupt vector.
   Public Function GetInterruptVector() As Byte
      Dim IRQ As New Integer?
      Dim Vector As Byte = &HFF%

      If Not InterruptInProgress Then
         IRQ = FindPendingUnmaskedIRQ()
         If IRQ IsNot Nothing Then
            PendingIRQs = ToByte(PendingIRQs And Not (&H1% << IRQ))
            '' InterruptInProgress = True ''<<<--- May cause freezing. Also appears to be not really necessary.
            Vector = CByte(&H8% + IRQ)
         End If
      End If

      Return Vector
   End Function

   'This procedure returns the mask register value.
   Public Function ReadMask() As Byte
      Return IRQMask
   End Function

   'This procedure handles raised IRQs.
   Public Sub RaiseIRQ(IRQ As Integer)
      If IRQ >= &H0% AndAlso IRQ <= &H7% Then
         PendingIRQs = ToByte(PendingIRQs Or CByte(&H1% << IRQ))
      End If
   End Sub

   'This procedure handles commands sent to the ICW 1.
   Public Sub WriteCommand(Value As Byte)
      If Value = NON_SPECIFIC_EOI Then
         InterruptInProgress = False
      End If
   End Sub

   'This procedure handles commands sent to the ICW 2.
   Public Sub WriteMask(Value As Byte)
      IRQMask = Value
   End Sub
End Class
