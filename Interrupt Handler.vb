'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Windows.Forms

'This module contains the default interrupt handler.
Public Module InterruptHandlerModule
   Private Const CURSOR_MASK As Integer = &H1F1F%    'Defines the cursor end/start bits.
   Private Const VIDEO_MODE_MASK As Byte = &H7F%     'Defines the bits indicating a video mode.
   Private Const ZERO_FLAG_INDEX As Integer = &H6%   'Defines the zero flag's bit index.

   'This structure defines an area on the screen in text modes.
   Private Structure ScreenAreaStr
      Public ULCRow As Integer      'Defines the upper left corner's row.
      Public ULCColumn As Integer   'Defines the upper left corner's column.
      Public LRCRow As Integer      'Defines the lower right corner's row.
      Public LRCColumn As Integer   'Defines the lower right corner's column.
   End Structure

   'This procedure clears the emulated screen buffer.
   Private Sub ClearScreen()
      Try
         Dim Count As New Integer
         Dim Position As New Integer

         Select Case DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
            Case VideoModesE.Text80x25Mono
               Count = (TEXT_80_X_25_MONO_BUFFER_SIZE \ &H2%)
               Position = AddressesE.Text80x25Mono
               Do While Count > &H0%
                  CPU.PutWord(Position, &H200%)
                  Count -= &H1%
                  Position += &H2%
               Loop
         End Select
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles the specified interrupt and returns whether or not is succeeded.
   Public Function HandleInterrupt(Number As Integer, AH As Integer) As Boolean
      Try
         Dim Attribute As New Byte
         Dim Character As New Byte
         Dim Count As New Integer
         Dim Flags As Integer = CPU.GetWord((CInt(CPU.Registers((CPU8086Class.SegmentRegistersE.SS))) << &H4%) + CInt(CPU.Registers(CPU8086Class.Registers16BitE.SP)) + &H4%)
         Dim Position As New Integer
         Dim Success As Boolean = False
         Dim VideoMode As New Byte
         Dim VideoModeBit7 As New Boolean
         Dim VideoPage As New Byte

         Select Case Number
            Case &H10%
               Select Case AH
                  Case &H0%
                     VideoMode = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL))
                     VideoModeBit7 = CBool(VideoMode >> &H7%)
                     VideoMode = VideoMode And VIDEO_MODE_MASK

                     If [Enum].IsDefined(GetType(VideoModesE), VideoMode) Then
                        CPU.Memory(AddressesE.VideoMode) = VideoMode
                        CPU.Memory(AddressesE.VideoModeOptions) = CByte(SetBit(CPU.Memory(AddressesE.VideoModeOptions), VideoModeBit7, Index:=&H7%))
                        CPU.PutWord(AddressesE.CursorScanLines, Word:=CURSOR_DEFAULT)
                        CPU.Memory(AddressesE.VideoPage) = &H0%
                        For Page As Integer = &H0% To MAXIMUM_VIDEO_PAGE_COUNT - &H1%
                           CPU.PutWord(AddressesE.CursorPositions + (Page * &H2%), Word:=&H0%)
                        Next Page
                        ClearScreen()
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
                        CPU.Registers(CPU8086Class.Registers16BitE.CX, NewValue:=CPU.GetWord(AddressesE.CursorScanLines))
                        CPU.Registers(CPU8086Class.Registers16BitE.DX, NewValue:=CPU.GetWord(AddressesE.CursorPositions + (VideoPage * &H2%)))
                     End If
                     Success = True
                  Case &H5%
                     VideoPage = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL))
                     If VideoPage < MAXIMUM_VIDEO_PAGE_COUNT Then
                        CPU.Memory(AddressesE.VideoPage) = VideoPage
                     End If
                     Success = True
                  Case &H6%
                     ScrollScreen(Up:=True, ScrollArea:=New ScreenAreaStr With {.ULCRow = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.CH)), .ULCColumn = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.CL)), .LRCRow = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.DH)), .LRCColumn = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.DL))}, Count:=CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)))
                     Success = True
                  Case &H7%
                     ScrollScreen(Up:=False, ScrollArea:=New ScreenAreaStr With {.ULCRow = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.CH)), .ULCColumn = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.CL)), .LRCRow = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.DH)), .LRCColumn = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.DL))}, Count:=CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)))
                     Success = True
                  Case &H9%
                     Select Case DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
                        Case VideoModesE.Text80x25Mono
                           Character = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL))
                           Attribute = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.BL))
                           Count = CInt(CPU.Registers(CPU8086Class.Registers16BitE.CX))
                           Position = AddressesE.Text80x25Mono + (Cursor.Y * &HA0%) + (Cursor.X * &H2%)
                           Do While Count > &H0%
                              CPU.Memory(Position) = Character
                              CPU.Memory(Position + &H1%) = Attribute
                              Count -= &H1%
                              Position += &H2%
                           Loop
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
                  Case &H12%
                     Select Case DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
                        Case VideoModesE.Text80x25Mono
                           Success = True
                     End Select
               End Select
            Case &H11%
               CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=CPU.GetWord(AddressesE.EquipmentFlags))
               Success = True
            Case &H12%
               CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=CPU.GetWord(AddressesE.MemorySize))
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
                     Flags = SetBit(Flags, (CInt(CPU.Registers(CPU8086Class.Registers16BitE.AX)) = &H0%), ZERO_FLAG_INDEX)
                     Success = True
               End Select
            Case Else
               Success = HandleMSDOSInterrupt(Number, AH, Flags)
         End Select

         CPU.PutWord((CInt(CPU.Registers((CPU8086Class.SegmentRegistersE.SS))) << &H4%) + CInt(CPU.Registers(CPU8086Class.Registers16BitE.SP)) + &H4%, Flags)

         If Success Then CPU.ExecuteOpcode(CPU8086Class.OpcodesE.IRET)

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function

   'This procedure scrolls the screen.
   Private Sub ScrollScreen(Up As Boolean, ScrollArea As ScreenAreaStr, Count As Integer)
      Try
         Dim Attribute As New Integer
         Dim CharacterCell As New Integer

         Select Case DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
            Case VideoModesE.Text80x25Mono
               If Count = &H0% OrElse Count > TEXT_80_X_25_LINE_COUNT Then
                  ClearScreen()
               Else
                  Attribute = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.BH))
                  For Scroll As Integer = &H1% To Count
                     Select Case Up
                        Case True
                           For Row As Integer = ScrollArea.ULCRow To ScrollArea.LRCRow
                              For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
                                 If Row <= ScrollArea.LRCRow Then
                                    CharacterCell = CPU.GetWord(AddressesE.Text80x25Mono + ((Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)))
                                    CPU.PutWord(AddressesE.Text80x25Mono + ((Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)), Attribute)
                                 Else
                                    CharacterCell = Attribute
                                 End If
                                 If Row >= ScrollArea.ULCRow Then
                                    CPU.PutWord(AddressesE.Text80x25Mono + (((Row - &H1%) * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)), CharacterCell)
                                 End If
                              Next Column
                           Next Row
                        Case False
                           For Row As Integer = ScrollArea.LRCRow To ScrollArea.ULCRow Step -&H1%
                              For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
                                 If Row >= ScrollArea.ULCRow Then
                                    CharacterCell = CPU.GetWord(AddressesE.Text80x25Mono + ((Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)))
                                    CPU.PutWord(AddressesE.Text80x25Mono + ((Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)), Attribute)
                                 Else
                                    CharacterCell = Attribute
                                 End If
                                 If Row <= ScrollArea.LRCRow Then
                                    CPU.PutWord(AddressesE.Text80x25Mono + (((Row + &H1%) * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)), CharacterCell)
                                 End If
                              Next Column
                           Next Row
                     End Select
                  Next Scroll
               End If
         End Select
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure emulates character output in Teletype mode.
   Public Sub TeleType(Character As Byte)
      Try
         Dim ScrollArea As New ScreenAreaStr With {.ULCColumn = &H0%, .ULCRow = &H0%, .LRCColumn = TEXT_80_X_25_COLUMN_COUNT - &H1%, .LRCRow = TEXT_80_X_25_LINE_COUNT - &H1%}

         Select Case DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
            Case VideoModesE.Text80x25Mono
               CursorBlink.Enabled = False

               Select Case DirectCast(Character, TeletypeE)
                  Case TeletypeE.BEL
                     Console.Beep()
                  Case TeletypeE.BS
                     If Cursor.X > &H0% Then Cursor.X -= &H1%
                  Case TeletypeE.CR
                     Cursor.X = &H0%
                  Case TeletypeE.LF
                     If Cursor.Y < TEXT_80_X_25_LINE_COUNT Then
                        Cursor.Y += &H1%
                     Else
                        ScrollScreen(Up:=True, ScrollArea, Count:=&H1%)
                     End If
                  Case TeletypeE.TAB
                     Cursor.X = ((((Cursor.X + &H1%) \ &H8%) + &H1%) * &H8%) - &H1%
                     If Cursor.X >= TEXT_80_X_25_COLUMN_COUNT Then
                        Cursor.X = &H0%
                        If Cursor.Y < TEXT_80_X_25_LINE_COUNT Then
                           Cursor.Y += &H1%
                        Else
                           ScrollScreen(Up:=True, ScrollArea, Count:=&H1%)
                        End If
                     End If
                  Case Else
                     CPU.Memory(AddressesE.Text80x25Mono + (Cursor.Y * &HA0%) + (Cursor.X * &H2%)) = Character

                     If Cursor.X < TEXT_80_X_25_COLUMN_COUNT Then
                        Cursor.X += &H1%
                     Else
                        Cursor.X = &H0%
                        If Cursor.Y < TEXT_80_X_25_LINE_COUNT Then
                           Cursor.Y += &H1%
                        Else
                           ScrollScreen(Up:=True, ScrollArea, Count:=&H1%)
                        End If
                     End If
               End Select

               CPU.Memory(AddressesE.CursorPositions) = CByte(Cursor.X)
               CPU.Memory(AddressesE.CursorPositions + &H1%) = CByte(Cursor.Y)
               CursorBlink.Enabled = True
         End Select
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Module
