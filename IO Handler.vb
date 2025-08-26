'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Environment

'This module contains the default I/O handler.
Public Module IOHandlerModule

   'This enumeration lists the I/O ports recognized by the CPU.
   Public Enum IOPortsE As Integer
      Reserved1 = &H10%                 'Reserved.
      Reserved2 = &H1F%                 'Reserved.
      PITCounter0 = &H40%               'Time of day clock.
      PITCounter1 = &H41%               'RAM refresh counter.
      PITCounter2 = &H42%               'Cassette and speaker.
      PITModeControl = &H43%            'Mode control register.
      PPIPortB = &H61%                  'Port B output.
      Reserved3 = &HE0%                 'Reserved.
      Reserved4 = &HEF%                 'Reserved.
      Reserved5 = &HF8%                 'Reserved.
      Reserved6 = &HFF%                 'Reserved.
      Reserved7 = &H21F%                'Reserved.
      Reserved8 = &H220%                'Reserved.
      Reserved9 = &H26F%                'Reserved.
      Reserved10 = &H280%               'Reserved.
      Reserved11 = &H2AF%               'Reserved.
      Reserved12 = &H2E8%               'Reserved.
      Reserved13 = &H2EF%               'Reserved.
      Reserved14 = &H2F0%               'Reserved.
      Reserved15 = &H2F7%               'Reserved.
      Reserved16 = &H330%               'Reserved.
      Reserved17 = &H33F%               'Reserved.
      Reserved18 = &H340%               'Reserved.
      Reserved19 = &H35F%               'Reserved.
      MDA3B0 = &H3B0%                   '6845 MDA.
      MDA3B1 = &H3B1%                   '6845 MDA.
      MDA3B2 = &H3B2%                   '6845 MDA.
      MDA3B3 = &H3B3%                   '6845 MDA.
      MDAIndex = &H3B4%                 'Index register.
      MDAData = &H3B5%                  'Data register.
      MDA3B6 = &H3B6%                   'Decodes to 0x3B4.
      MDA3B7 = &H3B7%                   'Decodes to 0x3B5.
      MDAMode = &H3B8%                  'Mode control register.
      MDAColor = &H3B9%                 'Color select register.
      MDAStatus = &H3BA%                'Status register.
      MDALightPenStrobeReset = &H3BB%   'Light pen strobe reset.
      CGA3D0 = &H3D0%                   '6845 CGA.
      CGA3D1 = &H3D1%                   '6845 CGA.
      CGA3D2 = &H3D2%                   '6845 CGA.
      CGA3D3 = &H3D3%                   '6845 CGA.
      CGAIndex = &H3D4%                 'Index register.
      CGAData = &H3D5%                  'Data register.
      CGA3D6 = &H3D6%                   'Decodes to 0x3D4.
      CGA3D7 = &H3D7%                   'Decodes to 0x3D5.
      CGAMode = &H3D8%                  'Mode control register.
      CGAColor = &H3D9%                 'Color select register.
      CGAStatus = &H3DA%                'Status register.
      CGALightPenStrobeReset = &H3DB%   'Light pen strobe reset.
      CGAPresetLightPenLatch = &H3DC%   'Preset light pen latch.
   End Enum

   Private Const PIT_IO_PORT_MASK As Integer = &H3%   'Defines the PIT I/O port number bits.

   Private ReadOnly MCC As New MCCClass   'Contains the 6845 Motorola CRT Controller.
   Private ReadOnly PPI As New PPIClass   'Contains the 8255 Programmable Peripheral Interface .

   'This procedure attempts to read from the specified I/O port and returns the result.
   Public Function ReadIOPort(Port As Integer) As Integer?
      Try
         Dim Value As Integer? = Nothing

         Select Case Port
            Case IOPortsE.MDAStatus
               Value = MCC.Status()
            Case IOPortsE.PITCounter0 To IOPortsE.PITCounter2
               Value = PIT.ReadCounter(DirectCast(Port And PIT_IO_PORT_MASK, PITClass.CountersE))
            Case IOPortsE.PITModeControl
               Value = &H0%
            Case IOPortsE.PPIPortB
               Value = PPI.PortB()
            Case IOPortsE.Reserved1 To IOPortsE.Reserved2, IOPortsE.Reserved3 To IOPortsE.Reserved4, IOPortsE.Reserved5 To IOPortsE.Reserved6, IOPortsE.Reserved7, IOPortsE.Reserved8 To IOPortsE.Reserved9, IOPortsE.Reserved10 To IOPortsE.Reserved11, IOPortsE.Reserved12 To IOPortsE.Reserved13, IOPortsE.Reserved14 To IOPortsE.Reserved15, IOPortsE.Reserved16 To IOPortsE.Reserved17, IOPortsE.Reserved18 To IOPortsE.Reserved19
               Value = &HFF%
         End Select

         Return Value
      Catch ExceptionO As Exception
         SyncLock Synchronizer
            CPUEvent.Append($"{ExceptionO.Message}{NewLine}")
         End SyncLock
      End Try

      Return Nothing
   End Function

   'This procedure attempts to write the specified I/O port and returns whether or not it succeeded.
   Public Function WriteIOPort(Port As Integer, Value As Integer) As Boolean
      Try
         Dim Success As Boolean = True

         Select Case Port
            Case IOPortsE.CGAColor
               MCC.ColorSelect(Value)
            Case IOPortsE.PPIPortB
               PPI.PortB(Value)
            Case IOPortsE.PITCounter0 To IOPortsE.PITCounter2
               PIT.WriteCounter(DirectCast(Port And PIT_IO_PORT_MASK, PITClass.CountersE), NewValue:=CByte(Value))
            Case IOPortsE.PITModeControl
               PIT.ModeControl(NewValue:=Value)
            Case IOPortsE.Reserved1 To IOPortsE.Reserved2, IOPortsE.Reserved3 To IOPortsE.Reserved4, IOPortsE.Reserved5 To IOPortsE.Reserved6, IOPortsE.Reserved7, IOPortsE.Reserved8 To IOPortsE.Reserved9, IOPortsE.Reserved10 To IOPortsE.Reserved11, IOPortsE.Reserved12 To IOPortsE.Reserved13, IOPortsE.Reserved14 To IOPortsE.Reserved15, IOPortsE.Reserved16 To IOPortsE.Reserved17, IOPortsE.Reserved18 To IOPortsE.Reserved19
               Success = True
            Case Else
               Success = False
         End Select

         Return Success
      Catch ExceptionO As Exception
         SyncLock Synchronizer
            CPUEvent.Append($"{ExceptionO.Message}{NewLine}")
         End SyncLock
      End Try

      Return False
   End Function
End Module
