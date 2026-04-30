'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Emulator8086Program.CPU8086Class
Imports System
Imports System.Convert
Imports System.Environment
Imports System.Threading.Tasks
Imports System.Windows.Forms

'This module contains the default interrupt handler.
Public Module InterruptHandlerModule
   Public Const CARRY_FLAG_INDEX As Integer = &H0%   'Defines the carry flag's bit index.
   Public Const ZERO_FLAG_INDEX As Integer = &H6%    'Defines the zero flag's bit index.
   Private Const CURSOR_MASK As Integer = &H1F1F%     'Defines the cursor end/start bits.
   Private Const VIDEO_MODE_MASK As Byte = &H7F%      'Defines the bits indicating a video mode.

   'This procedure handles any pending hardware interrupts.
   Public Sub ExecuteHardwareInterrupts()
      Try
         Dim Vector As Integer = PIC.GetInterruptVector()

         If Not Vector = &HFF% Then
            CPU.ExecuteInterrupt(OpcodesE.INT, Vector)
         End If

         If CPU.DoSystemTimerTick Then
            CPU.ExecuteInterrupt(OpcodesE.INT, SYSTEM_TIMER_TICK)
            CPU.DoSystemTimerTick = False
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles the specified interrupt and returns whether or not is succeeded.
   Public Function HandleInterrupt(Vector As Integer, Optional AH As Integer = Nothing, Optional IRET As Boolean = True) As Boolean
      Try
         Dim Address As New Integer
         Dim AL As New Integer
         Dim Attribute As New Byte
         Dim Character As New Byte
         Dim Count As New Integer
         Dim Flags As Integer = CPU.GetWord((CPU.Registers((SegmentRegistersE.SS)) << &H4%) + CPU.Registers(Registers16BitE.SP) + &H4%)
         Dim Mask As New Byte
         Dim Pixel As New Integer
         Dim PixelColor As New Byte
         Dim Position As New Integer
         Dim Shift As New Integer
         Dim Success As Boolean = False
         Dim Tracing As Boolean = CPU.Tracing
         Dim Value As New Integer?
         Dim VideoMode As New Byte
         Dim VideoModeBit7 As New Boolean
         Dim VideoModeValid As New Boolean
         Dim VideoPage As New Byte
         Dim VideoPageAddress As New Integer
         Dim x As New Integer
         Dim y As New Integer

         CPU.Tracing = False

         Select Case Vector
            Case &H8%
               UpdateClockCounter()
               PIC.WriteCommand(&H20%)
               Success = True
            Case &H9%
               CPU.Memory(AddressesE.KeyboardFlags) = ToByte(GetKeyboardFlags() And &HFF%)
               CPU.Memory(AddressesE.KeyboardFlags + &H1%) = ToByte(GetKeyboardFlags() >> &H8%)
               Success = True
            Case &H10%
               Select Case AH
                  Case &H0%
                     VideoMode = CByte(CPU.Registers(SubRegisters8BitE.AL))
                     VideoModeBit7 = CBool(VideoMode >> &H7%)
                     VideoMode = VideoMode And VIDEO_MODE_MASK

                     Select Case DirectCast(VideoMode, VideoModesE)
                        Case VideoModesE.Text80x25Mono
                           If MCC.IsMDA Then VideoModeValid = True
                        Case Else
                           If Not MCC.IsMDA Then VideoModeValid = [Enum].IsDefined(GetType(VideoModesE), VideoMode)
                     End Select

                     If VideoModeValid Then
                        CPU.Memory(AddressesE.VideoMode) = VideoMode
                        CPU.Memory(AddressesE.VideoModeOptions) = CByte(SET_BIT(CPU.Memory(AddressesE.VideoModeOptions), VideoModeBit7, &H7%))

                        CurrentVideoMode = DirectCast(VideoMode, VideoModesE)
                     End If

                     SwitchVideoAdapter()
                     Success = True
                  Case &H1%
                     If CPU.Registers(Registers16BitE.CX) = CURSOR_DISABLED Then
                        CPU.PutWord(AddressesE.CursorScanLines, Word:=CURSOR_DISABLED)
                     Else
                        CPU.PutWord(AddressesE.CursorScanLines, Word:=CPU.Registers(Registers16BitE.CX) And CURSOR_MASK)
                     End If
                     Success = True
                  Case &H2%
                     VideoPage = CByte(CPU.Registers(SubRegisters8BitE.BH))
                     If VideoPage < MAXIMUM_VIDEO_PAGE_COUNT Then
                        CPU.PutWord(AddressesE.CursorPositions + (VideoPage * &H2%), Word:=CPU.Registers(Registers16BitE.DX))
                        CursorPositionUpdate()
                     End If
                     Success = True
                  Case &H3%
                     VideoPage = CByte(CPU.Registers(SubRegisters8BitE.BH))
                     If VideoPage < MAXIMUM_VIDEO_PAGE_COUNT Then
                        CPU.Registers(Registers16BitE.CX, NewValue:=CPU.GetWord(AddressesE.CursorScanLines))
                        CPU.Registers(Registers16BitE.DX, NewValue:=CPU.GetWord(AddressesE.CursorPositions + (VideoPage * &H2%)))
                     End If
                     Success = True
                  Case &H5%
                     VideoPage = CByte(CPU.Registers(SubRegisters8BitE.AL))
                     If VideoPage < MAXIMUM_VIDEO_PAGE_COUNT Then
                        CPU.Memory(AddressesE.VideoPage) = VideoPage
                     End If
                     Success = True
                  Case &H6%
                     VideoAdapter.ScrollBuffer(Up:=True, ScrollArea:=New VideoAdapterClass.ScreenAreaStr With {.ULCRow = CPU.Registers(SubRegisters8BitE.CH), .ULCColumn = CPU.Registers(SubRegisters8BitE.CL), .LRCRow = CPU.Registers(SubRegisters8BitE.DH), .LRCColumn = CPU.Registers(SubRegisters8BitE.DL)}, Count:=CPU.Registers(SubRegisters8BitE.AL))
                     Success = True
                  Case &H7%
                     VideoAdapter.ScrollBuffer(Up:=False, ScrollArea:=New VideoAdapterClass.ScreenAreaStr With {.ULCRow = CPU.Registers(SubRegisters8BitE.CH), .ULCColumn = CPU.Registers(SubRegisters8BitE.CL), .LRCRow = CPU.Registers(SubRegisters8BitE.DH), .LRCColumn = CPU.Registers(SubRegisters8BitE.DL)}, Count:=CPU.Registers(SubRegisters8BitE.AL))
                     Success = True
                  Case &H8%
                     Select Case CurrentVideoMode
                        Case VideoModesE.Text80x25Color, VideoModesE.Text80x25Mono
                           CursorPositionUpdate()
                           CPU.Registers(Registers16BitE.AX, NewValue:=CPU.GetWord(AddressesE.Text80x25MonoPage0 + (Cursor.Y * &HA0%) + (Cursor.X * &H2%)))
                           Success = True
                     End Select
                  Case &H9%, &HA%
                     Select Case CurrentVideoMode
                        Case VideoModesE.CGA320x200A, VideoModesE.CGA320x200B, VideoModesE.VGA320x200
                           Character = CByte(CPU.Registers(SubRegisters8BitE.AL))
                           Attribute = CByte(CPU.Registers(SubRegisters8BitE.BL))
                           Count = CPU.Registers(Registers16BitE.CX)
                           CursorPositionUpdate()
                           Do While Count > &H0%
                              VideoAdapter.DrawCharacter(Character, Attribute)
                              If Cursor.X < MCC.ColumnCount() Then
                                 Cursor.X += 1
                              Else
                                 Cursor.X = 0
                                 Cursor.Y += 1
                              End If
                              Count -= &H1%
                           Loop
                        Case VideoModesE.Text80x25Color, VideoModesE.Text80x25Mono
                           Select Case CurrentVideoMode
                              Case VideoModesE.Text80x25Color
                                 VideoPageAddress = AddressesE.Text80x25ColorPage0
                              Case VideoModesE.Text80x25Mono
                                 VideoPageAddress = AddressesE.Text80x25MonoPage0
                           End Select
                           Character = CByte(CPU.Registers(SubRegisters8BitE.AL))
                           Attribute = CByte(CPU.Registers(SubRegisters8BitE.BL))
                           Count = CPU.Registers(Registers16BitE.CX)
                           CursorPositionUpdate()
                           Position = VideoPageAddress + (Cursor.Y * &HA0%) + (Cursor.X * &H2%)
                           Do While Count > &H0%
                              CPU.Memory(Position) = Character
                              If AH = &H9% Then CPU.Memory(Position + &H1%) = Attribute
                              Count -= &H1%
                              Position += &H2%
                           Loop
                     End Select
                     Success = True
                  Case &HB%
                     Select Case CPU.Registers(SubRegisters8BitE.BH)
                        Case &H0%
                           MCC.ActivePalette(0) = MCC.BACKGROUND_COLORS(CPU.Registers(SubRegisters8BitE.BL) And &HF%)
                           MCC.SelectIntensity((CPU.Registers(SubRegisters8BitE.BL) And MCCClass.INTENSITY_BIT) = MCCClass.INTENSITY_BIT)
                        Case &H1%
                           MCC.ColorSelect(CPU.Registers(SubRegisters8BitE.BL) And &H1%)
                     End Select
                     Success = True
                  Case &HC%
                     Select Case CurrentVideoMode
                        Case VideoModesE.CGA320x200A, VideoModesE.CGA320x200B
                           x = CPU.Registers(Registers16BitE.CX)
                           y = CPU.Registers(Registers16BitE.DX)

                           AL = CByte(CPU.Registers(SubRegisters8BitE.AL))
                           PixelColor = CByte(AL And &H3%)
                           Position = AddressesE.CGA320x200 + If((y And 1) = 0, 0, CGA_320_x_200_BUFFER_SIZE \ 2) + (y \ 2) * 80 + (x \ 4)
                           Pixel = x And &H3%
                           Shift = (&H3% - Pixel) * &H2%
                           Mask = CByte(&H3% << Shift)
                           Value = CPU.Memory(Position)

                           If (AL And &H80%) = &H0% Then
                              Value = (Value And Not Mask) Or CByte(PixelColor << Shift)
                           Else
                              Value = Value Xor CByte(PixelColor << Shift)
                           End If

                           CPU.Memory(Position) = CByte(Value)
                        Case VideoModesE.VGA320x200
                           x = CPU.Registers(Registers16BitE.CX)
                           y = CPU.Registers(Registers16BitE.DX)
                           AL = CByte(CPU.Registers(SubRegisters8BitE.AL))
                           Position = AddressesE.VGA320x200 + ((y * 320) + x)
                           CPU.Memory(Position) = CByte(AL)
                     End Select

                     Success = True
                  Case &HE%
                     TeleType(CByte(CPU.Registers(SubRegisters8BitE.AL)))
                     Success = True
                  Case &HF%
                     VideoMode = CurrentVideoMode
                     CPU.Registers(SubRegisters8BitE.AH, NewValue:=MCC.ColumnCount())
                     VideoMode = VideoMode Or (CPU.Memory(AddressesE.VideoModeOptions) >> &H7%)
                     CPU.Registers(SubRegisters8BitE.AL, NewValue:=VideoMode)
                     CPU.Registers(SubRegisters8BitE.BH, NewValue:=CPU.Memory(AddressesE.VideoPage))
                     Success = True
                  Case &H10%
                     Select Case CurrentVideoMode
                        Case VideoModesE.Text80x25Mono
                           Success = True
                        Case VideoModesE.Text80x25Color
                           Select Case CByte(CPU.Registers(SubRegisters8BitE.AL))
                              Case &H3%
                                 MCC.BlinkingOn = CBool(CPU.Registers(SubRegisters8BitE.BL))
                                 Success = True
                           End Select
                     End Select
                  Case &H11%
                     Select Case CByte(CPU.Registers(SubRegisters8BitE.AL))
                        Case &H30%
                           Select Case CByte(CPU.Registers(SubRegisters8BitE.BL))
                              Case &H0%
                                 Address = &H1F% * &H4%
                                 CPU.Registers(SegmentRegistersE.ES, NewValue:=CPU.GetWord(Address + &H2%))
                                 CPU.Registers(Registers16BitE.BP, NewValue:=CPU.GetWord(Address))
                           End Select
                     End Select
                     Success = True
                  Case &H12%
                     Select Case CurrentVideoMode
                        Case VideoModesE.Text80x25Color
                           Success = True
                        Case VideoModesE.Text80x25Mono
                           Success = True
                     End Select
                  Case &H13%
                     WriteString()
                     Success = True
                  Case &H1A%
                     Success = True
                  Case &H1B%
                     Select Case CurrentVideoMode
                        Case VideoModesE.Text80x25Mono
                        Case Else
                           If CPU.Registers(Registers16BitE.BX) = &H0% Then
                              CPU.Registers(SubRegisters8BitE.AL, NewValue:=&H1B%)
                              Address = (CPU.Registers(SegmentRegistersE.ES) << &H4%)
                              Position = CPU.Registers(Registers16BitE.DI)
                              For Each [Byte] As Byte In VGA.GetDynamicFunctionality()
                                 CPU.Memory(Address + (Position And &HFFFF%)) = [Byte]
                              Next [Byte]
                           End If
                     End Select
                     Success = True
                  Case &H1C%
                     Select Case CurrentVideoMode
                        Case VideoModesE.Text80x25Mono
                           Success = True
                     End Select
                  Case &H4F%, &HEF%, &HFE%, &HFF%
                     Success = True
               End Select
            Case &H11%
               CPU.Registers(Registers16BitE.AX, NewValue:=CPU.GetWord(AddressesE.EquipmentFlags))
               Success = True
            Case &H12%
               CPU.Registers(Registers16BitE.AX, NewValue:=CPU.GetWord(AddressesE.BIOSMemorySize))
               Success = True
            Case &H15%
               Select Case AH
                  Case &H6%, &HC0%, &HC2%
                     Success = True
               End Select
            Case &H16%
               Select Case AH
                  Case &H0%
                     CPU.Registers(Registers16BitE.AX, NewValue:=&H0%)
                     Do
                        If CPU.Clock.Status = TaskStatus.Running Then
                           ExecuteHardwareInterrupts()
                        End If
                        Application.DoEvents()
                        Value = LastBIOSKeyCode()
                        If Value IsNot Nothing Then CPU.Registers(Registers16BitE.AX, NewValue:=Value)
                     Loop While (CPU.Registers(Registers16BitE.AX) = &H0%) AndAlso (Not CPU.ClockToken.IsCancellationRequested)

                     LastBIOSKeyCode(, Clear:=True)
                     Success = True
                  Case &H1%
                     Value = LastBIOSKeyCode()
                     If Value IsNot Nothing Then CPU.Registers(Registers16BitE.AX, NewValue:=Value)
                     Flags = SET_BIT(Flags, (Value Is Nothing), ZERO_FLAG_INDEX)
                     Success = True
                  Case &H2%
                     CPU.Registers(SubRegisters8BitE.AL, NewValue:=CPU.Memory(AddressesE.KeyboardFlags))
                     Success = True
               End Select
            Case &H17%
               Select Case AH
                  Case &H1%
                     Success = True
               End Select
            Case &H1A%
               Select Case AH
                  Case &H0%
                     CPU.Registers(SubRegisters8BitE.AL, NewValue:=CPU.Memory(AddressesE.ClockRollover))
                     CPU.Registers(Registers16BitE.CX, NewValue:=CPU.GetWord(AddressesE.Clock + &H2%))
                     CPU.Registers(Registers16BitE.DX, NewValue:=CPU.GetWord(AddressesE.Clock))
                     Success = True
                  Case &H1%
                     CPU.PutWord(AddressesE.Clock + &H2%, CPU.Registers(Registers16BitE.CX))
                     CPU.PutWord(AddressesE.Clock, CPU.Registers(Registers16BitE.DX))
                     Success = True
               End Select
            Case &H1C%
               Success = True
            Case &H20%, &H22%
               CPU.ClockToken.Cancel()
               TerminateProgram($"Program terminated.{NewLine}")
               Success = True
            Case &H21%
               Select Case AH
                  Case &H0%
                     CPU.ClockToken.Cancel()
                     TerminateProgram($"Program terminated.{NewLine}")
                     Success = True
                  Case &H31%
                     CPU.ClockToken.Cancel()
                     TerminateProgram($"Terminate and stay resident.{NewLine}")
                     Success = True
                  Case &H4C%
                     CPU.ClockToken.Cancel()
                     TerminateProgram($"Program terminated with return code: {CPU.Registers(SubRegisters8BitE.AL):X2}.{NewLine}")
                     Success = True
                  Case Else
                     Success = HandleMSDOSInterrupt(Vector, AH, Flags)
               End Select
            Case &H23%
               CPU.ClockToken.Cancel()
               TerminateProgram($"CTRL+Break.{NewLine}")
               Success = True
            Case &H24%
               CPU.ClockToken.Cancel()
               TerminateProgram($"INT 24h - Critical error.{NewLine}")
               Success = True
            Case &H27%
               CPU.ClockToken.Cancel()
               TerminateProgram($"Terminate and stay resident.{NewLine}")
               Success = True
            Case &H33%
               Select Case AH
                  Case &H0%
                     CPU.Registers(Registers16BitE.AX, NewValue:=&H0%)
                     Success = True
               End Select
            Case Else
               Success = HandleMSDOSInterrupt(Vector, AH, Flags)
         End Select

         If IRET Then
            CPU.PutWord((CPU.Registers((SegmentRegistersE.SS)) << &H4%) + CPU.Registers(Registers16BitE.SP) + &H4%, Flags)
            If Success Then CPU.ExecuteOpcode(OpcodesE.IRET)
         End If

         CPU.Tracing = Tracing

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function
End Module
