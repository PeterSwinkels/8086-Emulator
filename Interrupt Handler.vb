'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Threading.Tasks
Imports System.Windows.Forms

'This module contains the default interrupt handler.
Public Module InterruptHandlerModule
   Public Const CARRY_FLAG_INDEX As Integer = &H0%   'Defines the carry flag's bit index.
   Private Const CURSOR_MASK As Integer = &H1F1F%    'Defines the cursor end/start bits.
   Private Const VIDEO_MODE_MASK As Byte = &H7F%     'Defines the bits indicating a video mode.
   Private Const ZERO_FLAG_INDEX As Integer = &H6%   'Defines the zero flag's bit index.

   'This procedure handles the specified interrupt and returns whether or not is succeeded.
   Public Function HandleInterrupt(Number As Integer, AH As Integer) As Boolean
      Try
         Dim Address As New Integer
         Dim Attribute As New Byte
         Dim Character As New Byte
         Dim Count As New Integer
         Dim Flags As Integer = CPU.GET_WORD((CInt(CPU.Registers((CPU8086Class.SegmentRegistersE.SS))) << &H4%) + CInt(CPU.Registers(CPU8086Class.Registers16BitE.SP)) + &H4%)
         Dim Position As New Integer
         Dim Success As Boolean = False
         Dim Tracing As Boolean = CPU.Tracing
         Dim VideoMode As New Byte
         Dim VideoModeBit7 As New Boolean
         Dim VideoPage As New Byte
         Dim VideoPageAddress As New Integer

         CPU.Tracing = False

         Select Case Number
            Case &H8%
               UpdateClockCounter()
               Success = True
            Case &H10%
               Select Case AH
                  Case &H0%
                     VideoMode = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL))
                     VideoModeBit7 = CBool(VideoMode >> &H7%)
                     VideoMode = VideoMode And VIDEO_MODE_MASK

                     If [Enum].IsDefined(GetType(VideoModesE), VideoMode) Then
                        CPU.Memory(AddressesE.VideoMode) = VideoMode
                        CPU.Memory(AddressesE.VideoModeOptions) = CByte(SET_BIT(CPU.Memory(AddressesE.VideoModeOptions), VideoModeBit7, &H7%))

                        CurrentVideoMode = DirectCast(VideoMode, VideoModesE)
                        SwitchVideoAdapter()
                     End If
                     Success = True
                  Case &H1%
                     If CInt(CPU.Registers(CPU8086Class.Registers16BitE.CX)) = CURSOR_DISABLED Then
                        CPU.PutWord(AddressesE.CursorScanLines, Word:=CURSOR_DISABLED)
                     Else
                        CPU.PutWord(AddressesE.CursorScanLines, Word:=CInt(CPU.Registers(CPU8086Class.Registers16BitE.CX)) And CURSOR_MASK)
                     End If
                     Success = True
                  Case &H2%
                     VideoPage = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.BH))
                     If VideoPage < MAXIMUM_VIDEO_PAGE_COUNT Then
                        CPU.PutWord(AddressesE.CursorPositions + (VideoPage * &H2%), Word:=CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX)))
                        CursorPositionUpdate()
                     End If
                     Success = True
                  Case &H3%
                     VideoPage = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.BH))
                     If VideoPage < MAXIMUM_VIDEO_PAGE_COUNT Then
                        CPU.Registers(CPU8086Class.Registers16BitE.CX, NewValue:=CPU.GET_WORD(AddressesE.CursorScanLines))
                        CPU.Registers(CPU8086Class.Registers16BitE.DX, NewValue:=CPU.GET_WORD(AddressesE.CursorPositions + (VideoPage * &H2%)))
                     End If
                     Success = True
                  Case &H5%
                     VideoPage = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL))
                     If VideoPage < MAXIMUM_VIDEO_PAGE_COUNT Then
                        CPU.Memory(AddressesE.VideoPage) = VideoPage
                     End If
                     Success = True
                  Case &H6%
                     VideoAdapter.ScrollBuffer(Up:=True, ScrollArea:=New VideoAdapterClass.ScreenAreaStr With {.ULCRow = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.CH)), .ULCColumn = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.CL)), .LRCRow = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.DH)), .LRCColumn = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.DL))}, Count:=CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)))
                     Success = True
                  Case &H7%
                     VideoAdapter.ScrollBuffer(Up:=False, ScrollArea:=New VideoAdapterClass.ScreenAreaStr With {.ULCRow = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.CH)), .ULCColumn = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.CL)), .LRCRow = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.DH)), .LRCColumn = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.DL))}, Count:=CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)))
                     Success = True
                  Case &H8%
                     Select Case DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
                        Case VideoModesE.Text80x25Mono
                           CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=CPU.GET_WORD(If(CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.BH)) = &H0%, AddressesE.Text80x25MonoPage0, AddressesE.Text80x25MonoPage1) + (Cursor.Y * &HA0%) + (Cursor.X * &H2%)))
                           Success = True
                     End Select
                  Case &H9%, &HA%
                     Select Case DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
                        Case VideoModesE.Text80x25Mono
                           VideoPageAddress = If(CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.BH)) = &H0%, AddressesE.Text80x25MonoPage0, AddressesE.Text80x25MonoPage1)
                           Character = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL))
                           Attribute = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.BL))
                           Count = CInt(CPU.Registers(CPU8086Class.Registers16BitE.CX))
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
                     Select Case CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.BH))
                        Case &H0%
                           BackgroundColor = BACKGROUND_COLORS(CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.BL)) And &HF%)
                        Case &H1%
                           Select Case CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.BL))
                              Case &H0%
                                 Palette = PALETTE0
                              Case &H1%
                                 Palette = PALETTE1
                           End Select
                     End Select
                     Success = True
                  Case &HE%
                     TeleType(CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)))
                     Success = True
                  Case &HF%
                     VideoMode = CPU.Memory(AddressesE.VideoMode)
                     CPU.Registers(CPU8086Class.SubRegisters8BitE.AH, NewValue:=If(VideoMode = VideoModesE.Text40x25Color OrElse VideoMode = VideoModesE.Text40x25Mono, &H28%, &H50%))
                     VideoMode = VideoMode Or (CPU.Memory(AddressesE.VideoModeOptions) >> &H7%)
                     CPU.Registers(CPU8086Class.SubRegisters8BitE.AL, NewValue:=VideoMode)
                     CPU.Registers(CPU8086Class.SubRegisters8BitE.BH, NewValue:=CPU.Memory(AddressesE.VideoPage))
                     Success = True
                  Case &H11%
                     Select Case CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL))
                        Case &H30%
                           Select Case CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.BL))
                              Case &H0%
                                 Address = &H1F% * &H4%
                                 CPU.Registers(CPU8086Class.SegmentRegistersE.ES, NewValue:=CPU.GET_WORD(Address + &H2%))
                                 CPU.Registers(CPU8086Class.Registers16BitE.BP, NewValue:=CPU.GET_WORD(Address))
                                 Success = True
                           End Select
                     End Select
                  Case &H12%
                     Select Case DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
                        Case VideoModesE.Text80x25Mono
                           Success = True
                     End Select
                  Case &HFE%
                     Success = True
               End Select
            Case &H11%
               CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=CPU.GET_WORD(AddressesE.EquipmentFlags))
               Success = True
            Case &H12%
               CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=CPU.GET_WORD(AddressesE.MemorySize))
               Success = True
            Case &H16%
               Select Case AH
                  Case &H0%
                     Do
                        Application.DoEvents()
                        CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=LastBIOSKeyCode())
                     Loop While (CInt(CPU.Registers(CPU8086Class.Registers16BitE.AX)) = &H0%) AndAlso (Not CPU.ClockToken.IsCancellationRequested)

                     LastBIOSKeyCode(, Clear:=True)
                     Success = True
                  Case &H1%
                     CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=LastBIOSKeyCode())
                     Flags = SET_BIT(Flags, (CInt(CPU.Registers(CPU8086Class.Registers16BitE.AX)) = &H0%), ZERO_FLAG_INDEX)
                     Success = True
               End Select
            Case &H1A%
               Select Case AH
                  Case &H0%
                     CPU.Registers(CPU8086Class.Registers16BitE.CX, NewValue:=CPU.Memory(AddressesE.Clock + &H2%))
                     CPU.Registers(CPU8086Class.Registers16BitE.DX, NewValue:=CPU.Memory(AddressesE.Clock))
                     Success = True
               End Select
            Case &H1C%
               Success = True
            Case Else
               Success = HandleMSDOSInterrupt(Number, AH, Flags)
         End Select

         CPU.PutWord((CInt(CPU.Registers((CPU8086Class.SegmentRegistersE.SS))) << &H4%) + CInt(CPU.Registers(CPU8086Class.Registers16BitE.SP)) + &H4%, Flags)

         If Success Then CPU.ExecuteOpcode(CPU8086Class.OpcodesE.IRET)

         CPU.Tracing = Tracing

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function

   'This procedure emulates character output in Teletype mode.
   Public Sub TeleType(Character As Byte)
      Try
         Dim ScrollArea As New VideoAdapterClass.ScreenAreaStr With {.ULCColumn = &H0%, .ULCRow = &H0%, .LRCColumn = TEXT_80_X_25_COLUMN_COUNT - &H1%, .LRCRow = TEXT_80_X_25_LINE_COUNT}
         Dim VideoPage As Integer = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.BH))
         Dim VideoPageAddress As Integer = If(VideoPage = &H0%, AddressesE.Text80x25MonoPage0, AddressesE.Text80x25MonoPage1)
         Dim Attribute As Integer = CPU.Memory(VideoPageAddress + ((Cursor.Y * TEXT_80_X_25_BYTES_PER_ROW) + (Cursor.X * &H2%)) + &H1%)

         CPU.Registers(CPU8086Class.SubRegisters8BitE.BH, NewValue:=Attribute)

         Select Case DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
            Case VideoModesE.Text80x25Mono
               CursorBlink.Enabled = False

               Select Case DirectCast(Character, TeletypeE)
                  Case TeletypeE.BEL
                     Task.Run(Sub() Console.Beep())
                  Case TeletypeE.BS
                     If Cursor.X > &H0% Then Cursor.X -= &H1%
                  Case TeletypeE.CR
                     Cursor.X = &H0%
                  Case TeletypeE.LF
                     If Cursor.Y < TEXT_80_X_25_LINE_COUNT - &H1% Then
                        Cursor.Y += &H1%
                     Else
                        VideoAdapter.ScrollBuffer(Up:=True, ScrollArea, Count:=&H1%)
                     End If
                  Case TeletypeE.TAB
                     Cursor.X = ((((Cursor.X + &H1%) \ &H8%) + &H1%) * &H8%) - &H1%
                     If Cursor.X >= TEXT_80_X_25_COLUMN_COUNT - &H1% Then
                        Cursor.X = &H0%
                        If Cursor.Y < TEXT_80_X_25_LINE_COUNT - &H1% Then
                           Cursor.Y += &H1%
                        Else
                           VideoAdapter.ScrollBuffer(Up:=True, ScrollArea, Count:=&H1%)
                        End If
                     End If
                  Case Else
                     CPU.Memory(VideoPageAddress + (Cursor.Y * TEXT_80_X_25_BYTES_PER_ROW) + (Cursor.X * &H2%)) = Character

                     If Cursor.X < TEXT_80_X_25_COLUMN_COUNT - &H1% Then
                        Cursor.X += &H1%
                     Else
                        Cursor.X = &H0%
                        If Cursor.Y < TEXT_80_X_25_LINE_COUNT - &H1% Then
                           Cursor.Y += &H1%
                        Else
                           VideoAdapter.ScrollBuffer(Up:=True, ScrollArea, Count:=&H1%)
                        End If
                     End If
               End Select

               CPU.Memory(AddressesE.CursorPositions) = CByte(Cursor.X)
               CPU.Memory(AddressesE.CursorPositions + &H1%) = CByte(Cursor.Y)
               CPU.Registers(CPU8086Class.SubRegisters8BitE.BH, NewValue:=VideoPage)

               CursorBlink.Enabled = True
         End Select
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Module
