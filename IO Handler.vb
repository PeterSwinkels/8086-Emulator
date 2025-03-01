﻿'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System

'This module contains the default I/O handler.
Public Module IO_Handler

   'This enumeration lists the I/O ports recognized by the CPU.
   Private Enum IOPortsE As Integer
      PITCounter0 = &H40%               'Time of day clock.
      PITCounter1 = &H41%               'RAM refresh counter.
      PITCounter2 = &H42%               'Cassette and speaker.
      PITModeControl = &H43%            'Mode control register.
      MMA3B0 = &H3B0%                   '6845 MMA.
      MMA3B1 = &H3B1%                   '6845 MMA.
      MMA3B2 = &H3B2%                   '6845 MMA.
      MMA3B3 = &H3B3%                   '6845 MMA.
      MMAIndex = &H3B4%                 'Index register.
      MMAData = &H3B5%                  'Data register.
      MMA3B6 = &H3B6%                   'Decodes to 0x3B4.
      MMA3B7 = &H3B7%                   'Decodes to 0x3B5.
      MMAMode = &H3B8%                  'Mode control register.
      MMAColor = &H3B9%                 'Color select register.
      MMAStatus = &H3BA%                'Status register.
      MMALightPenStrobeReset = &H3BB%   'Light pen strobe reset.
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
   Private ReadOnly PIT As New PITClass               'Contains the 8253 PIT.

   'This procedure attempts to read from the specified I/O port and returns the result.
   Public Function ReadIOPort(Port As Integer) As Integer?
      Try
         Dim Value As Integer? = Nothing

         Select Case Port
            Case IOPortsE.PITCounter0 To IOPortsE.PITCounter2
               Value = PIT.ReadCounter(Port And PIT_IO_PORT_MASK)
            Case IOPortsE.PITModeControl
               Value = &H0%
         End Select

         Return Value
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure attempts to write the specified I/O port and returns whether or not it succeeded.
   Public Function WriteIOPort(Port As Integer, Value As Integer) As Boolean
      Try
         Dim Success As Boolean = True

         Select Case Port
            Case IOPortsE.PITCounter0 To IOPortsE.PITCounter2
               PIT.WriteCounter(Port And PIT_IO_PORT_MASK, NewValue:=CByte(Value))
            Case IOPortsE.PITModeControl
               PIT.ModeControl(NewValue:=Value)
            Case Else
               Success = False
         End Select

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function
End Module
