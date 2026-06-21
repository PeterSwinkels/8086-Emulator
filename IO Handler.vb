'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Convert
Imports System.Environment

'This module contains the default I/O handler.
Public Module IOHandlerModule

   'This enumeration lists the I/O ports recognized by the CPU.
   Public Enum IOPortsE As Integer
      DMAChannel0Address = &H0%                'DMA channel 0 address register.
      DMAChannel0WordCount = &H1%              'DMA channel 0 address word count.
      DMAChannel1Address = &H2%                'DMA channel 1 address register.
      DMAChannel1WordCount = &H3%              'DMA channel 1 address word count.
      DMAChannel2Address = &H4%                'DMA channel 2 address register.
      DMAChannel2WordCount = &H5%              'DMA channel 2 address word count.
      DMAChannel3Address = &H6%                'DMA channel 3 address register.
      DMAChannel3WordCount = &H7%              'DMA channel 3 address word count.
      DMAStatusCommandRegister = &H8%          'DMA status/command register.
      DMARequestRegister = &H9%                'DMA request register.
      DMAMaskRegister = &HA%                   'DMA mask register.
      DMAModeRegister = &HB%                   'DMA mode register.
      DMAClearMSBLSBFlipFlop = &HC%            'DMA clear MSB/LSB flip flop.
      DMAMasterClearTemporaryRegister = &HD%   'DMA master clear temporary register.
      DMAClearMaskRegister = &HE%              'DMA clear mask register.
      DMAMultipleMaskRegister = &HF%           'DMA multiple mask register.
      Reserved1 = &H10%                        'Reserved.
      Reserved2 = &H1F%                        'Reserved.
      PICICW1 = &H20%                          'Initialization Command Word 1.
      PICICW2 = &H21%                          'Initialization Command Word 2.
      PITCounter0 = &H40%                      'Time of day clock.
      PITCounter1 = &H41%                      'RAM refresh counter.
      PITCounter2 = &H42%                      'Cassette and speaker.
      PITModeControl = &H43%                   'Mode control register.
      Port49h = &H49%                          'Port 49h.
      SN76496 = &HC0%                          'TI SN76496 Programmable Tone/Noise Generator. (PCjr)
      Port216h% = &H216%                       'Port 216h.
      Port21Ah% = &H21A%                       'Port 21Ah.
      Port21Eh% = &H21E%                       'Port 21Eh.
      Port2C0h = &H2C0%                        'Port 2C0h.
      Port2C1h = &H2C1%                        'Port 2C1h.
      Port2C3h = &H2C3%                        'Port 2C3h.
      Port388h = &H388%                        'Port 388h.
      Port389h = &H389%                        'Port 389h.
      KeyboardIO = &H60%                       'Keyboard I/O register.
      PPIPortB = &H61%                         'Port B output.
      PPIPortC = &H62%                         'Port C input.
      Reserved3 = &HE0%                        'Reserved.
      Reserved4 = &HEF%                        'Reserved.
      Reserved5 = &HF8%                        'Reserved.
      Reserved6 = &HFF%                        'Reserved.
      Joystick = &H201%                        'Joystick.
      Reserved7 = &H21F%                       'Reserved.
      Reserved8 = &H220%                       'Reserved.
      Reserved9 = &H26F%                       'Reserved.
      Reserved10 = &H280%                      'Reserved.
      Reserved11 = &H2AF%                      'Reserved.
      Reserved12 = &H2E8%                      'Reserved.
      Reserved13 = &H2EF%                      'Reserved.
      Reserved14 = &H2F0%                      'Reserved.
      Reserved15 = &H2F7%                      'Reserved.
      Reserved16 = &H330%                      'Reserved.
      Reserved17 = &H33F%                      'Reserved.
      Reserved18 = &H340%                      'Reserved.
      Reserved19 = &H35F%                      'Reserved.
      MDA3B0 = &H3B0%                          '6845 MDA.
      MDA3B1 = &H3B1%                          '6845 MDA.
      MDA3B2 = &H3B2%                          '6845 MDA.
      MDA3B3 = &H3B3%                          '6845 MDA.
      MDAIndex = &H3B4%                        'Index register.
      MDAData = &H3B5%                         'Data register.
      MDA3B6 = &H3B6%                          'Decodes to 0x3B4.
      MDA3B7 = &H3B7%                          'Decodes to 0x3B5.
      MDAMode = &H3B8%                         'Mode control register.
      MDAColor = &H3B9%                        'Color select register.
      MDAStatus = &H3BA%                       'Status register.
      MDALightPenStrobeReset = &H3BB%          'Light pen strobe reset.
      HGCConfigurationSwitch = &H3BF%          'Hercules Graphics Card (HGC) "Configuration Switch".
      VGAVideoDACPELAddress = &H3C8%           'VGA video DAC PEL address.
      VGAVideoDAC = &H3C9%                     'VGA video DAC.
      CGA3D0 = &H3D0%                          '6845 CGA.
      CGA3D1 = &H3D1%                          '6845 CGA.
      CGA3D2 = &H3D2%                          '6845 CGA.
      CGA3D3 = &H3D3%                          '6845 CGA.
      CGAIndex = &H3D4%                        'Index register.
      CGAData = &H3D5%                         'Data register.
      CGA3D6 = &H3D6%                          'Decodes to 0x3D4.
      CGA3D7 = &H3D7%                          'Decodes to 0x3D5.
      CGAMode = &H3D8%                         'Mode control register.
      CGAColor = &H3D9%                        'Color select register.
      CGAStatus = &H3DA%                       'Status register.
      CGALightPenStrobeReset = &H3DB%          'Light pen strobe reset.
      CGAPresetLightPenLatch = &H3DC%          'Preset light pen latch.
   End Enum

   Private Const PIT_IO_PORT_MASK As Integer = &H3%   'Defines the PIT I/O port number bits.

   'This procedure attempts to read from the specified I/O port and returns the result.
   Public Function ReadIOPort(Port As Integer) As Integer?
      Try
         Dim Value As Integer? = Nothing

         Select Case (Port And &HFFFF%)
            Case IOPortsE.DMAStatusCommandRegister
               Value = &H0%
            Case IOPortsE.DMAChannel0Address To IOPortsE.DMAChannel3WordCount, IOPortsE.DMARequestRegister To IOPortsE.DMAMultipleMaskRegister
               Value = &HFF%
            Case IOPortsE.CGA3D0, IOPortsE.CGA3D6, IOPortsE.CGAIndex
               Value = If(MCC.IsMDA, &HFF%, MCC.SelectedRegister)
            Case IOPortsE.CGA3D1, IOPortsE.CGA3D7, IOPortsE.CGAData
               Value = If(MCC.IsMDA, &HFF%, MCC.Register())
            Case IOPortsE.CGA3D2, IOPortsE.CGA3D3, IOPortsE.CGAColor, IOPortsE.CGAMode, IOPortsE.CGALightPenStrobeReset, IOPortsE.CGAPresetLightPenLatch
               Value = &HFF%
            Case IOPortsE.CGAStatus
               Value = If(MCC.IsMDA, &HFF%, MCC.CGAStatus())
            Case IOPortsE.Joystick
               Value = &HFF%
            Case IOPortsE.KeyboardIO
               Value = KeyScancode
            Case IOPortsE.MDA3B0, IOPortsE.MDA3B6, IOPortsE.MDAIndex
               Value = MCC.SelectedRegister
            Case IOPortsE.MDA3B1, IOPortsE.MDAData
               Value = MCC.Register()
            Case IOPortsE.MDAStatus
               Value = MCC.MDAStatus()
            Case IOPortsE.PICICW2
               Value = PIC.ReadMask()
            Case IOPortsE.PITCounter0 To IOPortsE.PITCounter2
               Value = PIT.ReadCounter(DirectCast(Port And PIT_IO_PORT_MASK, PITClass.CountersE))
            Case IOPortsE.PITModeControl
               Value = &H0%
            Case IOPortsE.Port49h
               Value = &HFF%
            Case IOPortsE.Port216h, IOPortsE.Port21Ah%, IOPortsE.Port21Eh
               Value = &HFF%
            Case IOPortsE.Port2C0h To IOPortsE.Port2C1h, IOPortsE.Port388h
               Value = &HFF%
            Case IOPortsE.PPIPortB
               Value = PPI.PortB()
            Case IOPortsE.PPIPortC
               Value = &HFF%
            Case IOPortsE.Reserved1 To IOPortsE.Reserved2,
                 IOPortsE.Reserved3 To IOPortsE.Reserved4,
                 IOPortsE.Reserved5 To IOPortsE.Reserved6,
                 IOPortsE.Reserved7,
                 IOPortsE.Reserved8 To IOPortsE.Reserved9,
                 IOPortsE.Reserved10 To IOPortsE.Reserved11,
                 IOPortsE.Reserved12 To IOPortsE.Reserved13,
                 IOPortsE.Reserved14 To IOPortsE.Reserved15,
                 IOPortsE.Reserved16 To IOPortsE.Reserved17,
                 IOPortsE.Reserved18 To IOPortsE.Reserved19
               Value = &HFF%
            Case IOPortsE.SN76496
               Value = &HFF%
         End Select

         Return Value
      Catch ExceptionO As Exception
         SyncLock SYNCHRONIZER
            CPU_EVENT.Append($"{ExceptionO.Message}{NewLine}")
         End SyncLock
      End Try

      Return Nothing
   End Function

   'This procedure attempts to write the specified I/O port and returns whether or not it succeeded.
   Public Function WriteIOPort(Port As Integer, Value As Integer) As Boolean
      Try
         Dim Success As Boolean = True

         Select Case (Port And &HFFFF%)
            Case IOPortsE.DMAChannel0Address, IOPortsE.DMAChannel1Address, IOPortsE.DMAChannel2Address, IOPortsE.DMAChannel3Address
               DMA.UpdateChannelAddress(Port >> &H1%, ToByte(Value))
            Case IOPortsE.DMAClearMSBLSBFlipFlop
               DMA.WriteMSB = False
            Case IOPortsE.CGA3D0, IOPortsE.CGA3D6, IOPortsE.CGAIndex
               If Not MCC.IsMDA Then
                  MCC.SelectRegister(DirectCast(Value, MCCClass.RegistersE))
               End If
            Case IOPortsE.CGA3D1, IOPortsE.CGAData
               If Not MCC.IsMDA Then
                  MCC.Register(NewValue:=ToByte(Value))
               End If
            Case IOPortsE.CGAColor
               MCC.ColorSelect(Value)
            Case IOPortsE.CGAMode
               Success = True
            Case IOPortsE.HGCConfigurationSwitch
               Success = True
            Case IOPortsE.Joystick
               Success = True
            Case IOPortsE.MDA3B0, IOPortsE.MDAIndex, IOPortsE.MDA3B6
               MCC.SelectRegister(DirectCast(Value, MCCClass.RegistersE))
            Case IOPortsE.MDA3B1, IOPortsE.MDAData
               MCC.Register(NewValue:=ToByte(Value))
            Case IOPortsE.MDAColor, IOPortsE.MDAStatus
               Success = True
            Case IOPortsE.MDAMode
               If MCC.IsMDA Then
                  Select Case Value
                     Case &H3F%
                        MCC.BlinkingOn = True
                     Case &H40%
                        MCC.BlinkingOn = False
                  End Select
               End If

               CPU.Memory(AddressesE.CRTModeControlRegisterValue) = ToByte(Value)
            Case IOPortsE.PICICW1
               PIC.WriteCommand(ToByte(Value))
            Case IOPortsE.PICICW2
               PIC.WriteMask(ToByte(Value))
            Case IOPortsE.PPIPortB
               PPI.PortB(Value)
            Case IOPortsE.PITCounter0 To IOPortsE.PITCounter2
               PIT.WriteCounter(DirectCast(Port And PIT_IO_PORT_MASK, PITClass.CountersE), NewValue:=CByte(Value))
            Case IOPortsE.PITModeControl
               PIT.ModeControl(NewValue:=Value)
            Case IOPortsE.Port216h, IOPortsE.Port21Ah%, IOPortsE.Port21Eh
               Success = True
            Case IOPortsE.Port2C0h To IOPortsE.Port2C1h, IOPortsE.Port2C3h, IOPortsE.Port388h To IOPortsE.Port389h
               Success = True
            Case IOPortsE.SN76496
               Success = True
            Case IOPortsE.VGAVideoDAC
               VGA.WriteToDAC(Value)
            Case IOPortsE.VGAVideoDACPELAddress
               VGA.SelectAddress(Value)
            Case IOPortsE.Reserved1 To IOPortsE.Reserved2,
                 IOPortsE.Reserved3 To IOPortsE.Reserved4,
                 IOPortsE.Reserved5 To IOPortsE.Reserved6,
                 IOPortsE.Reserved7,
                 IOPortsE.Reserved8 To IOPortsE.Reserved9,
                 IOPortsE.Reserved10 To IOPortsE.Reserved11,
                 IOPortsE.Reserved12 To IOPortsE.Reserved13,
                 IOPortsE.Reserved14 To IOPortsE.Reserved15,
                 IOPortsE.Reserved16 To IOPortsE.Reserved17,
                 IOPortsE.Reserved18 To IOPortsE.Reserved19
               Success = True
            Case Else
               Success = False
         End Select

         Return Success
      Catch ExceptionO As Exception
         SyncLock SYNCHRONIZER
            CPU_EVENT.Append($"{ExceptionO.Message}{NewLine}")
         End SyncLock
      End Try

      Return False
   End Function
End Module
