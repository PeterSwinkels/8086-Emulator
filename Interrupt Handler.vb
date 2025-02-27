'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Windows.Forms

'This module contains the default interrupt handler.
Public Module InterruptHandlerModule
   Private Const CURSOR_MASK As Integer = &H1F1F%      'Defines the cursor end/start bits.
   Private Const MS_DOS As Integer = &HFF00%           'Indicates that the operating is MS-DOS.
   Private Const MS_DOS_VERSION As Integer = &H1606%   'Defines the emulated MS-DOS version as 6.22.
   Private Const VIDEO_MODE_MASK As Byte = &H7F%       'Defines the bits indicating a video mode.

   'This procedure handles the specified interrupt and returns whether or not is succeeded.
   Public Function HandleInterrupt(Number As Integer, AH As Integer) As Boolean
      Try
         Dim Attribute As New Byte
         Dim Character As New Byte
         Dim Count As New Integer
         Dim Offset As New Integer
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
                     If [Enum].IsDefined(GetType(VideoModesE), CByte(VideoMode)) Then
                        CPU.Memory(AddressesE.VideoMode) = VideoMode
                        CPU.Memory(AddressesE.VideoModeOptions) = CByte(SetBit(CPU.Memory(AddressesE.VideoModeOptions), VideoModeBit7, Index:=&H7%))
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
                  Case &H9%
                     Select Case DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
                        Case VideoModesE.Text80x25Mono
                           Offset = AddressesE.Text80x25Mono + (Cursor.Y * &HA0%) + (Cursor.X * &H2%)
                           Character = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL))
                           Attribute = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.BL))
                           Count = CInt(CPU.Registers(CPU8086Class.Registers16BitE.CX))
                           Position = Offset
                           Do While Count > &H0%
                              CPU.Memory(Position) = Character
                              CPU.Memory(Position + &H1%) = Attribute
                              Count -= &H1%
                              Position += &H2%
                           Loop
                     End Select
                     Success = True
                  Case &HF%
                     VideoMode = CPU.Memory(AddressesE.VideoMode)
                     CPU.Registers(CPU8086Class.SubRegisters8BitE.AH, NewValue:=If(VideoMode = VideoModesE.Text40x25Color OrElse VideoMode = VideoModesE.Text40x25Mono, &H28%, &H50%))
                     VideoMode = VideoMode Or (CPU.Memory(AddressesE.VideoModeOptions) >> &H7%)
                     CPU.Registers(CPU8086Class.SubRegisters8BitE.AL, NewValue:=VideoMode)
                     CPU.Registers(CPU8086Class.SubRegisters8BitE.BH, NewValue:=CPU.Memory(AddressesE.VideoPage))
                     Success = True
               End Select
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
                     CPU.Registers(CPU8086Class.FlagRegistersE.ZF, NewValue:=(CInt(CPU.Registers(CPU8086Class.Registers16BitE.AX)) = &H0%))
                     Success = True
               End Select
            Case &H21%
               Select Case AH
                  Case &H30%
                     CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=MS_DOS_VERSION)
                     CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=MS_DOS)
                     Success = True
               End Select
         End Select

         If Success Then CPU.ExecuteOpcode(CPU8086Class.OpcodesE.IRET)

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function
End Module
