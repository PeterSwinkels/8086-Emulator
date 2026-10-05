'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Emulator8086Program.CPU8086Class
Imports System
Imports System.Convert
Imports System.Environment
Imports System.Linq
Imports System.Threading.Tasks
Imports System.Windows.Forms

'This module contains the default interrupt handler.
Public Module InterruptHandlerModule
   Public Const CARRY_FLAG_INDEX As Integer = &H0%           'Defines the carry flag's bit index.
   Public Const ZERO_FLAG_INDEX As Integer = &H6%            'Defines the zero flag's bit index.
   Private Const PRINTER_STATUS_NOT_BUSY As Integer = &H80%   'Defines the printer not busy status.
   Private Const VIDEO_MODE_MASK As Byte = &H7F%              'Defines the bits indicating a video mode.

   'This procedure handles the specified interrupt and returns whether or not is succeeded.
   Public Function HandleInterrupt(Vector As Integer, Optional AH As Integer = Nothing) As Boolean
      Try
         Dim Address As New Integer
         Dim AL As New Integer
         Dim Attribute As New Byte
         Dim Character As New Byte
         Dim Count As New Integer
         Dim Flags As Integer = Memory.GetWord((CPU.Registers((SegmentRegistersE.SS)) << &H4%) + CPU.Registers(Registers16BitE.SP) + &H4%)
         Dim Mask As New Byte
         Dim Pixel As New Integer
         Dim PixelColor As New Byte
         Dim Position As New Integer
         Dim RETF As Boolean = False
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
               Memory(AddressesE.KeyboardFlags) = ToByte(GetKeyboardFlags() And &HFF%)
               Memory(AddressesE.KeyboardFlags + &H1%) = ToByte(GetKeyboardFlags() >> &H8%)

               If LastBIOSKeyCode() IsNot Nothing Then
                  Memory.PutWord((BIOS_SEGMENT << &H4%) + Memory(AddressesE.KeyboardBufferHead), Word:=LastBIOSKeyCode().Value)

                  Memory(AddressesE.KeyboardBufferHead) = ToByte(Memory(AddressesE.KeyboardBufferHead) + &H2)

                  If Memory(AddressesE.KeyboardBufferHead) >= KEY_BUFFER_END Then
                     Memory(AddressesE.KeyboardBufferHead) = KEY_BUFFER_START
                  End If

                  If Memory(AddressesE.KeyboardBufferHead) >= Memory(AddressesE.KeyboardBufferTail) Then
                     Memory(AddressesE.KeyboardBufferTail) = ToByte(Memory(AddressesE.KeyboardBufferTail) + &H2)

                     If Memory(AddressesE.KeyboardBufferTail) >= KEY_BUFFER_END Then
                        Memory(AddressesE.KeyboardBufferTail) = KEY_BUFFER_START
                     End If
                  End If
               End If

               Success = True
            Case &HA%
               Success = True
            Case &H10%
               Select Case AH
                  Case &H0%
                     VideoMode = CByte(CPU.Registers(SubRegisters8BitE.AL))
                     VideoModeBit7 = CBool(VideoMode >> &H7%)
                     VideoMode = VideoMode And VIDEO_MODE_MASK

                     Select Case DirectCast(VideoMode, VideoModesE)
                        Case VideoModesE.Text80x25Mono_Hercules
                           If MCC.IsMDA Then VideoModeValid = True
                        Case Else
                           If Not MCC.IsMDA Then VideoModeValid = [Enum].IsDefined(GetType(VideoModesE), VideoMode)
                     End Select

                     If VideoModeValid Then
                        Memory(AddressesE.VideoMode) = VideoMode
                        Memory(AddressesE.VideoModeOptions) = CByte(SET_BIT(Memory(AddressesE.VideoModeOptions), VideoModeBit7, &H7%))

                        MCC.CurrentVideoMode = DirectCast(VideoMode, VideoModesE)
                     End If

                     SwitchVideoAdapter()
                     Success = True
                  Case &H1%
                     If CPU.Registers(Registers16BitE.CX) = CURSOR_DISABLED Then
                        Memory.PutWord(AddressesE.CursorScanLines, Word:=CURSOR_DISABLED)
                     Else
                        Memory.PutWord(AddressesE.CursorScanLines, Word:=CPU.Registers(Registers16BitE.CX))
                     End If
                     Success = True
                  Case &H2%
                     VideoPage = CByte(CPU.Registers(SubRegisters8BitE.BH))
                     If VideoPage < MAXIMUM_VIDEO_PAGE_COUNT Then
                        Memory.PutWord(AddressesE.CursorPositions + (VideoPage * &H2%), Word:=CPU.Registers(Registers16BitE.DX))
                        CursorPositionUpdate()
                     End If
                     Success = True
                  Case &H3%
                     VideoPage = CByte(CPU.Registers(SubRegisters8BitE.BH))
                     If VideoPage < MAXIMUM_VIDEO_PAGE_COUNT Then
                        CPU.Registers(Registers16BitE.CX, NewValue:=Memory.GetWord(AddressesE.CursorScanLines))
                        CPU.Registers(Registers16BitE.DX, NewValue:=Memory.GetWord(AddressesE.CursorPositions + (VideoPage * &H2%)))
                     End If
                     Success = True
                  Case &H5%
                     VideoPage = CByte(CPU.Registers(SubRegisters8BitE.AL))
                     If VideoPage < MCC.VideoPageCount() Then
                        Memory(AddressesE.VideoPage) = VideoPage
                     End If
                     Success = True
                  Case &H6%
                     VideoAdapter.ScrollBuffer(Up:=True, ScrollArea:=New VideoAdapterClass.ScreenAreaStr With {.ULCRow = CPU.Registers(SubRegisters8BitE.CH), .ULCColumn = CPU.Registers(SubRegisters8BitE.CL), .LRCRow = CPU.Registers(SubRegisters8BitE.DH), .LRCColumn = CPU.Registers(SubRegisters8BitE.DL)}, Count:=CPU.Registers(SubRegisters8BitE.AL))
                     Success = True
                  Case &H7%
                     VideoAdapter.ScrollBuffer(Up:=False, ScrollArea:=New VideoAdapterClass.ScreenAreaStr With {.ULCRow = CPU.Registers(SubRegisters8BitE.CH), .ULCColumn = CPU.Registers(SubRegisters8BitE.CL), .LRCRow = CPU.Registers(SubRegisters8BitE.DH), .LRCColumn = CPU.Registers(SubRegisters8BitE.DL)}, Count:=CPU.Registers(SubRegisters8BitE.AL))
                     Success = True
                  Case &H8%
                     Select Case MCC.CurrentVideoMode
                        Case VideoModesE.Text80x25Color, VideoModesE.Text80x25Gray, VideoModesE.Text80x25Mono_Hercules
                           CursorPositionUpdate()
                           CPU.Registers(Registers16BitE.AX, NewValue:=Memory.GetWord(AddressesE.Text80x25MonoBuffer + (Cursor.Y * &HA0%) + (Cursor.X * &H2%)))
                           Success = True
                     End Select
                  Case &H9%, &HA%
                     Select Case MCC.CurrentVideoMode
                        Case VideoModesE.CGA320x200A, VideoModesE.CGA320x200B, VideoModesE.CGA640x200, VideoModesE.VGA320x200
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
                        Case VideoModesE.Text80x25Color, VideoModesE.Text80x25Gray, VideoModesE.Text80x25Mono_Hercules
                           VideoPageAddress = MCC.VideoPageAddress()
                           Character = CByte(CPU.Registers(SubRegisters8BitE.AL))
                           Attribute = CByte(CPU.Registers(SubRegisters8BitE.BL))
                           Count = CPU.Registers(Registers16BitE.CX)
                           CursorPositionUpdate()
                           Position = VideoPageAddress + (Cursor.Y * &HA0%) + (Cursor.X * &H2%)
                           Do While Count > &H0%
                              Memory(Position) = Character
                              If AH = &H9% Then Memory(Position + &H1%) = Attribute
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
                           MCC.SelectActivePalette(CPU.Registers(SubRegisters8BitE.BL) And &H1%)
                     End Select
                     Success = True
                  Case &HC%
                     x = CPU.Registers(Registers16BitE.CX)
                     y = CPU.Registers(Registers16BitE.DX)
                     AL = CByte(CPU.Registers(SubRegisters8BitE.AL))
                     Select Case MCC.CurrentVideoMode
                        Case VideoModesE.CGA320x200A, VideoModesE.CGA320x200B
                           PixelColor = CByte(AL And &H3%)
                           Position = AddressesE.CGABuffer + If((y And 1) = 0, 0, VideoPageSizesE.CGA320x200A \ 2) + (y \ 2) * 80 + (x \ 4)
                           Pixel = x And &H3%
                           Shift = (&H3% - Pixel) * &H2%
                           Mask = CByte(&H3% << Shift)
                           Value = Memory(Position)

                           If (AL And &H80%) = &H0% Then
                              Value = (Value And Not Mask) Or CByte(PixelColor << Shift)
                           Else
                              Value = Value Xor CByte(PixelColor << Shift)
                           End If

                           Memory(Position) = CByte(Value)
                        Case VideoModesE.CGA640x200
                           PixelColor = CByte(AL And &H1%)
                           Position = AddressesE.CGABuffer + If((y And 1) = 0, 0, VideoPageSizesE.CGA640x200 \ 2) + (y \ 2) * 80 + (x \ 4)
                           Pixel = x And &H6%
                           Shift = (&H6% - Pixel) * &H4%
                           Mask = CByte(&H6% << Shift)
                           Value = Memory(Position)

                           If (AL And &H80%) = &H0% Then
                              Value = (Value And Not Mask) Or CByte(PixelColor << Shift)
                           Else
                              Value = Value Xor CByte(PixelColor << Shift)
                           End If

                           Memory(Position) = CByte(Value)
                        Case VideoModesE.VGA320x200
                           Position = AddressesE.VGABuffer + ((y * 320) + x)
                           Memory(Position) = CByte(AL)
                     End Select

                     Success = True
                  Case &HE%
                     Teletype(CByte(CPU.Registers(SubRegisters8BitE.AL)))
                     Success = True
                  Case &HF%
                     VideoMode = MCC.CurrentVideoMode
                     CPU.Registers(SubRegisters8BitE.AH, NewValue:=MCC.ColumnCount())
                     VideoMode = VideoMode Or (Memory(AddressesE.VideoModeOptions) >> &H7%)
                     CPU.Registers(SubRegisters8BitE.AL, NewValue:=VideoMode)
                     CPU.Registers(SubRegisters8BitE.BH, NewValue:=Memory(AddressesE.VideoPage))
                     Success = True
                  Case &H10%
                     Select Case MCC.CurrentVideoMode
                        Case VideoModesE.Text80x25Mono_Hercules
                           Success = True
                        Case VideoModesE.EGA320x200, VideoModesE.Text80x25Color, VideoModesE.Text80x25Gray, VideoModesE.VGA320x200
                           Select Case CPU.Registers(SubRegisters8BitE.AL)
                              Case &H1%
                                 Success = True
                              Case &H2%
                                 Address = (CPU.Registers(SegmentRegistersE.ES) << &H4%) + CPU.Registers(Registers16BitE.DX)
                                 EGA.SetEntirePalette(Memory.GetRange(Address, Length:=&H10%).ToArray())
                                 Success = True
                              Case &H3%
                                 MCC.BlinkingOn = CBool(CPU.Registers(SubRegisters8BitE.BL))
                                 Success = True
                              Case &H12%
                                 Select Case MCC.CurrentVideoMode
                                    Case VideoModesE.VGA320x200
                                       VGA.SetDACBlock()
                                       Success = True
                                 End Select
                              Case &H17%
                                 Select Case MCC.CurrentVideoMode
                                    Case VideoModesE.VGA320x200
                                       VGA.GetDACBlock()
                                       Success = True
                                 End Select
                           End Select
                     End Select
                  Case &H11%
                     Select Case CPU.Registers(SubRegisters8BitE.AL)
                        Case &H30%
                           Select Case CPU.Registers(SubRegisters8BitE.BL)
                              Case &H0%
                                 Address = &H1F% * &H4%
                                 CPU.Registers(SegmentRegistersE.ES, NewValue:=Memory.GetWord(Address + &H2%))
                                 CPU.Registers(Registers16BitE.BP, NewValue:=Memory.GetWord(Address))
                                 CPU.Registers(SubRegisters8BitE.DL, NewValue:=MCC.RowCount)
                           End Select
                     End Select
                     Success = True
                  Case &H12%
                     If MCC.IsMDA Then
                        Success = True
                     Else
                        Select Case CPU.Registers(SubRegisters8BitE.BL)
                           Case &H0%
                              Success = True
                           Case &H10%
                              CPU.Registers(SubRegisters8BitE.BH, NewValue:=If(EGAClass.MONO_MODE, &H1%, &H0%))
                              CPU.Registers(SubRegisters8BitE.BL, NewValue:=EGAClass.EGA_MEMORY_SIZE)
                              CPU.Registers(Registers16BitE.CX, NewValue:=EGAClass.EGA_FEATURE_SWITCH_BITS)
                              Success = True
                           Case &H20%
                              Success = True
                        End Select
                     End If
                  Case &H13%
                     WriteString()
                     Success = True
                  Case &H18%, &H19%
                     Success = True
                  Case &H1A%
                     If Not MCC.IsMDA Then
                        CPU.Registers(SubRegisters8BitE.AL, NewValue:=&H1A%)
                        CPU.Registers(Registers16BitE.BX, NewValue:=&H8%)
                     End If
                     Success = True
                  Case &H1B%
                     If Not MCC.IsMDA Then
                        If CPU.Registers(Registers16BitE.BX) = &H0% Then
                           CPU.Registers(SubRegisters8BitE.AL, NewValue:=&H1B%)
                           Address = (CPU.Registers(SegmentRegistersE.ES) << &H4%)
                           Position = CPU.Registers(Registers16BitE.DI)
                           Memory.PutRange(Address, VGA.GetDynamicFunctionality)
                        End If
                     End If
                     Success = True
                  Case &H1C%
                     Select Case MCC.CurrentVideoMode
                        Case VideoModesE.Text80x25Mono_Hercules
                           Success = True
                     End Select
                  Case &H30%
                     CPU.Registers(Registers16BitE.CX, NewValue:=&H0%)
                     CPU.Registers(Registers16BitE.DX, NewValue:=&H0%)
                     Success = True
                  Case &H4F%, &HBF%, &HEF%, &HFA%, &HFE%, &HFF%
                     Success = True
               End Select
            Case &H11%
               CPU.Registers(Registers16BitE.AX, NewValue:=Memory.GetWord(AddressesE.EquipmentFlags))
               Success = True
            Case &H12%
               CPU.Registers(Registers16BitE.AX, NewValue:=Memory.GetWord(AddressesE.BIOSMemorySize))
               Success = True
            Case &H15%
               Select Case AH
                  Case &H6%, &HC0%, &HC2%
                     Success = True
               End Select
            Case &H16%
               Select Case AH
                  Case &H0%, &H10%
                     CPU.Registers(Registers16BitE.AX, NewValue:=&H0%)
                     Do
                        If CPU.Clock.Status = TaskStatus.Running Then
                           CPU.ExecuteHardwareInterrupts()
                        End If
                        Application.DoEvents()
                        Value = LastBIOSKeyCode()

                        If Value IsNot Nothing Then
                           Select Case AH
                              Case &H0%
                                 If Array.IndexOf(EXTENDED_CODES, Value.Value) >= 0 Then
                                    Value = New Integer?
                                 End If
                           End Select
                        End If

                        If Value IsNot Nothing Then CPU.Registers(Registers16BitE.AX, NewValue:=Value)
                     Loop While (CPU.Registers(Registers16BitE.AX) = &H0%) AndAlso (Not CPU.ClockToken.IsCancellationRequested)

                     Memory(AddressesE.KeyboardBufferTail) = ToByte(Memory(AddressesE.KeyboardBufferTail) + &H2)
                     If Memory(AddressesE.KeyboardBufferTail) >= KEY_BUFFER_END Then
                        Memory(AddressesE.KeyboardBufferTail) = KEY_BUFFER_START
                     End If

                     LastBIOSKeyCode(, Clear:=True)
                     Success = True
                  Case &H1%
                     Value = LastBIOSKeyCode()
                     If Value IsNot Nothing Then CPU.Registers(Registers16BitE.AX, NewValue:=Value)
                     Flags = SET_BIT(Flags, (Value Is Nothing), ZERO_FLAG_INDEX)
                     Success = True
                  Case &H2%
                     CPU.Registers(SubRegisters8BitE.AL, NewValue:=Memory(AddressesE.KeyboardFlags))
                     Success = True
                  Case &H5%
                     WriteToKeyboardBuffer()
                     Success = True
               End Select
            Case &H17%
               Select Case AH
                  Case &H0%
                     PrintCharacter(CByte(CPU.Registers(SubRegisters8BitE.AL)), CPU.Registers(Registers16BitE.DX))
                     Success = True
                  Case &H1%, &H2%
                     CPU.Registers(SubRegisters8BitE.AH, NewValue:=PRINTER_STATUS_NOT_BUSY)
                     Success = True
               End Select
            Case &H1A%
               Select Case AH
                  Case &H0%
                     CPU.Registers(SubRegisters8BitE.AL, NewValue:=Memory(AddressesE.ClockRollover))
                     CPU.Registers(Registers16BitE.CX, NewValue:=Memory.GetWord(AddressesE.Clock + &H2%))
                     CPU.Registers(Registers16BitE.DX, NewValue:=Memory.GetWord(AddressesE.Clock))
                     Success = True
                  Case &H1%
                     Memory.PutWord(AddressesE.Clock + &H2%, Word:=CPU.Registers(Registers16BitE.CX))
                     Memory.PutWord(AddressesE.Clock, Word:=CPU.Registers(Registers16BitE.DX))
                     Success = True
               End Select
            Case &H1C%
               Success = True
            Case &H20%, &H22%
               CPU.ClockToken.Cancel()
               MSDOS.TerminateProgram($"Program terminated.{NewLine}")
               Success = True
            Case &H21%
               Select Case AH
                  Case &H0%
                     CPU.ClockToken.Cancel()
                     MSDOS.TerminateProgram($"Program terminated.{NewLine}")
                     Success = True
                  Case &H31%
                     CPU.ClockToken.Cancel()
                     MSDOS.TerminateProgram($"Terminate and stay resident.{NewLine}")
                     Success = True
                  Case &H4C%
                     CPU.ClockToken.Cancel()
                     MSDOS.TerminateProgram($"Program terminated with return code: {CPU.Registers(SubRegisters8BitE.AL):X2}.{NewLine}")
                     Success = True
                  Case Else
                     Success = MSDOS.HandleMSDOSInterrupt(Vector, AH, Flags:=Flags, RETF:=RETF)
               End Select
            Case &H23%
               CPU.ClockToken.Cancel()
               MSDOS.TerminateProgram($"CTRL+Break.{NewLine}")
               Success = True
            Case &H24%
               CPU.ClockToken.Cancel()
               MSDOS.TerminateProgram($"INT 24h - Critical error.{NewLine}")
               Success = True
            Case &H27%
               CPU.ClockToken.Cancel()
               MSDOS.TerminateProgram($"Terminate and stay resident.{NewLine}")
               Success = True
            Case &H33%
               Select Case AH
                  Case &H0%
                     CPU.Registers(Registers16BitE.AX, NewValue:=MOUSE_DRIVER_NOT_INSTALLED)
                     CPU.Registers(Registers16BitE.BX, NewValue:=MOUSE_BUTTON_COUNT)
                     Success = True
               End Select
            Case Else
               Success = MSDOS.HandleMSDOSInterrupt(Vector, AH, Flags:=Flags, RETF:=RETF)
         End Select

         Memory.PutWord((CPU.Registers((SegmentRegistersE.SS)) << &H4%) + CPU.Registers(Registers16BitE.SP) + &H4%, Word:=Flags)

         If Success Then CPU.ExecuteOpcode(If(RETF, OpcodesE.RETF, OpcodesE.IRET))

         CPU.Tracing = Tracing

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function
End Module
