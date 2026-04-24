'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Emulator8086Program.CPU8086Class
Imports System
Imports System.IO
Imports System.Threading.Tasks

'This module contains BIOS related procedures.
Public Module BIOSModule
   'This enumeration lists the flat addresses of BIOS data area locations.
   Public Enum AddressesE As Integer
      BIOS = &HF0000%                   'BIOS.
      BIOSMemorySize = &H413%           'BIOS Memory size.
      BIOSStack = &HF3000%              'BIOS stack segment.
      CGA320x200 = &HB8000%             '320x200 CGA video buffer.
      Characters = &HFFA6E%             'Character bitmaps.
      Clock = &H46C%                    'Clock.
      ClockRollover = &H470%            'Clock rollover flag.
      ColumnCount = &H44A%              'Column count.
      CRTControllerBasePOrt = &H463%    'CRT controller base port.
      CursorPositions = &H450%          'Cursor positions.
      CursorScanLines = &H460%          'Cursor scan line start/end.
      EquipmentFlags = &H410%           'Equipment flags.
      ExtendedCharacters = &HC0000%     'Extended character bitmaps.
      KeyboardFlags = &H417%            'Keyboard flags.
      MachineID = &HFFFFE%              'Machine ID.
      Text80x25ColorPage0 = &HB8000%    '80x25 color text video buffer.
      Text80x25MonoPage0 = &HB0000%     '80x25 monochrome text video buffer.
      VGA320x200 = &HA0000%             '320x200 VGA video buffer.
      VideoMode = &H449%                'Current video mode.
      VideoModeOptions = &H487%         'Video mode options.
      VideoPage = &H462%                'Current video page.
   End Enum

   'This enumeration lists the supported Teletype control characters
   Public Enum TeletypeE As Byte
      BEL = &H7%   'Bell.
      BS = &H8%    'Backspace.
      TAB = &H9%   'Tab.
      LF = &HA%    'Line feed.
      CR = &HD%    'Carriage return.
   End Enum

   'This enumeration lists the video modes.
   Public Enum VideoModesE As Byte
      None = &HFF%             'None.
      CGA320x200A = &H4%       '320x200 CGA.
      CGA320x200B = &H5%       '320x200 CGA.
      CGA640x200 = &H6%        '640x200 CGA.
      EGA320x200 = &HD%        '320x200 EGA.
      EGA640x200 = &HE%        '640x200 EGA.
      EGA640x350 = &H10%       '640x350 EGA.
      EGA640x350Mono = &HF%    '640x350 monochrome EGA.
      PCjr160x200 = &H8%       '160x200 PCjr.
      PCjr320x200 = &H9%       '320x200 PCjr.
      PCjr640x200 = &HA%       '640x200 PCjr.
      Text40x25Color = &H1%    '40x25 color text.
      Text40x25Mono = &H0%     '40x25 monochrome text.
      Text80x25Color = &H3%    '80x25 color text.
      Text80x25Gray = &H2%     '80x25 gray text.
      Text80x25Mono = &H7%     '80x25 monochrome text.
      VGA320x200 = &H13%       '320x200 VGA.
      VGA640x480 = &H12%       '640x480 VGA.
      VGA640x480Mono = &H11%   '640x480 monochrome vga.
   End Enum

   Public Const CGA_320_x_200_BUFFER_SIZE As Integer = &H4000%               'Defines the 320x200 CGA mode video memory's size.
   Public Const CGA_320_X_200_BYTES_PER_ROW As Integer = &H50%               'Defines the number of bytes per row used by 320x200 CGA mode 
   Public Const CGA_320_X_200_LINES_PER_CHARACTER As Integer = &H8%          'Defines the number of lines per character used by 320x200 CGA mode 
   Public Const CGA_320_X_200_PIXELS_PER_BYTE As Integer = &H4%              'Defines the number of pixels per byte used by 320x200 CGA mode.
   Public Const EXTENDED_CHARACTERS_VECTOR As Integer = &H1F%                'Defines the extended character bitmap pointer's location.
   Public Const INITIAL_MODE_FLAGS_MDA As Integer = &H30%                    'Defines the initial video mode in the equipment flags as MDA.
   Public Const INITIAL_MODE_FLAGS_NOT_MDA As Integer = &H20%                'Defines the initial video mode in the equipment flags as not MDA.
   Public Const INITIAL_STACK_SIZE As Integer = &H10%                        'Defines the stack size at start up.
   Public Const MAXIMUM_CLOCK_VALUE As Integer = &H1800B0%                   'Defines the highest value reached by the clock just before midnight.
   Public Const MAXIMUM_VIDEO_PAGE_COUNT As Integer = &H8%                   'Defines the maximum number of video pages.
   Public Const TEXT_80_X_25_BYTES_PER_ROW As Integer = &HA0%                'Defines the number of bytes per row used by 80x25 monochrome text mode 
   Public Const TEXT_80_X_25_COLOR_BUFFER_SIZE As Integer = &HFA0%           'Defines the 80x25 color text mode video memory's size.
   Public Const TEXT_80_X_25_MONO_BUFFER_SIZE As Integer = &HFA0%            'Defines the 80x25 monochrome text mode video memory's size.
   Public Const VGA_320_x_200_BUFFER_SIZE As Integer = &HFA00%               'Defines the 320x200 VGA mode video memory's size.
   Public Const VGA_320_X_200_BYTES_PER_ROW As Integer = &H140%              'Defines the number of bytes per row used by 320x200 VGA mode 
   Public Const VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE As Integer = &H8%    'Defines the number of per character side used by 320x200 VGA mode 
   Private Const BIOS_MEMORY_SIZE As Integer = &H280%                         'Defines the BIOS memory size in 1kb blocks.
   Private Const MACHINE_ID As Byte = &HFF%                                   'Defines the machine ID.
   Private Const TICKS_PER_SECOND As Double = 18.2064814814815                'Defines the number of clock ticks per second.

   Public ClockCounter As Integer = &H0%         'Contains the system clock counter.

   'This procedure loads the BIOS into memory.
   Public Sub LoadBIOS()
      Try
         Dim Address As New Integer
         Dim Data() As Byte = {}
         Dim Offset As Integer = AddressesE.BIOS

         For Vector As Integer = &H0% To &HFF%
            Address = Vector * &H4%
            CPU.PutWord(Address + &H2%, AddressesE.BIOS >> &H4%)
            CPU.PutWord(Address, Offset)
            Offset = WriteBytesToMemory({OpcodesE.EXT_INT, CByte(Vector), OpcodesE.IRET}, Offset)
         Next Vector

         If MCC.IsMDA Then
            CurrentVideoMode = VideoModesE.Text80x25Mono
            VideoAdapter = New Text80x25MonoClass
            CPU.PutWord(AddressesE.EquipmentFlags, INITIAL_MODE_FLAGS_MDA)
            CPU.Memory(AddressesE.VideoMode) = CurrentVideoMode
         Else
            CurrentVideoMode = VideoModesE.Text80x25Color
            VideoAdapter = New Text80x25ColorClass
            CPU.PutWord(AddressesE.EquipmentFlags, INITIAL_MODE_FLAGS_NOT_MDA)
            CPU.Memory(AddressesE.VideoMode) = CurrentVideoMode

            Address = EXTENDED_CHARACTERS_VECTOR * &H4%
            CPU.PutWord(Address + &H2%, AddressesE.ExtendedCharacters >> &H4%)
            CPU.PutWord(Address, AddressesE.ExtendedCharacters And &HFFFF%)
            Data = File.ReadAllBytes("EXTFONT.BIN")
            Data.CopyTo(CPU.Memory, AddressesE.ExtendedCharacters)
            Data = File.ReadAllBytes("FONT.BIN")
            Data.CopyTo(CPU.Memory, AddressesE.Characters)

            Array.Copy(VGA.STATIC_FUNCTIONALITY, &H0%, CPU.Memory, VGA.STATIC_FUNCTIONALITY_ADDRESS, VGA.STATIC_FUNCTIONALITY.Length)
         End If

         SwitchVideoAdapter()

         PIC.WriteMask(&HFC%)

         ClockCounter = CInt(DateTime.Now.TimeOfDay.TotalSeconds * TICKS_PER_SECOND)
         UpdateClockCounter()

         CPU.PutWord(AddressesE.BIOSMemorySize, BIOS_MEMORY_SIZE)
         CPU.Memory(AddressesE.ColumnCount) = MCC.ColumnCount()
         CPU.PutWord(AddressesE.CRTControllerBasePOrt, IOPortsE.MDA3B0)
         CPU.Memory(AddressesE.MachineID) = MACHINE_ID

         CPU.Registers(SegmentRegistersE.SS, NewValue:=AddressesE.BIOSStack)
         CPU.Registers(Registers16BitE.SP, NewValue:=INITIAL_STACK_SIZE)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure emulates character output in Teletype mode.
   Public Sub TeleType(Character As Byte, Optional Attribute As Integer? = Nothing)
      Try
         Dim ScrollArea As New VideoAdapterClass.ScreenAreaStr With {.ULCColumn = &H0%, .ULCRow = &H0%, .LRCColumn = MCC.ColumnCount() - &H1%, .LRCRow = MCC.RowCount()}
         Dim ScrollAttribute As New Byte
         Dim VideoPage As Integer = CInt(CPU.Registers(SubRegisters8BitE.BH))
         Dim VideoPageAddress As New Integer

         CursorPositionUpdate()

         Select Case CurrentVideoMode
            Case VideoModesE.CGA320x200A, VideoModesE.CGA320x200B
               VideoPageAddress = AddressesE.CGA320x200
               Attribute = CInt(CPU.Registers(SubRegisters8BitE.BL))

               Select Case DirectCast(Character, TeletypeE)
                  Case TeletypeE.BEL
                     Task.Run(Sub() Console.Beep())
                  Case TeletypeE.BS
                     If Cursor.X > &H0% Then Cursor.X -= &H1%
                  Case TeletypeE.CR
                     Cursor.X = &H0%
                  Case TeletypeE.LF
                     If Cursor.Y < MCC.RowCount() - &H1% Then
                        Cursor.Y += &H1%
                     Else
                        VideoAdapter.ScrollBuffer(Up:=True, ScrollArea, Count:=&H1%)
                     End If
                  Case TeletypeE.TAB
                     Cursor.X = ((((Cursor.X + &H1%) \ &H8%) + &H1%) * &H8%) - &H1%
                     If Cursor.X >= MCC.ColumnCount() - &H1% Then
                        Cursor.X = &H0%
                        If Cursor.Y < MCC.RowCount() - &H1% Then
                           Cursor.Y += &H1%
                        Else
                           VideoAdapter.ScrollBuffer(Up:=True, ScrollArea, Count:=&H1%)
                        End If
                     End If
                  Case Else
                     VideoAdapter.DrawCharacter(Character, Attribute.Value)

                     If Cursor.X < MCC.ColumnCount() - &H1% Then
                        Cursor.X += &H1%
                     Else
                        Cursor.X = &H0%
                        If Cursor.Y < MCC.RowCount() - &H1% Then
                           Cursor.Y += &H1%
                        Else
                           VideoAdapter.ScrollBuffer(Up:=True, ScrollArea, Count:=&H1%)
                        End If
                     End If
               End Select

               CPU.Memory(AddressesE.CursorPositions) = CByte(Cursor.X)
               CPU.Memory(AddressesE.CursorPositions + &H1%) = CByte(Cursor.Y)
            Case VideoModesE.Text80x25Color, VideoModesE.Text80x25Mono
               Select Case CurrentVideoMode
                  Case VideoModesE.Text80x25Color
                     VideoPageAddress = AddressesE.Text80x25ColorPage0
                  Case VideoModesE.Text80x25Mono
                     VideoPageAddress = AddressesE.Text80x25MonoPage0
               End Select

               ScrollAttribute = CPU.Memory(VideoPageAddress + ((Cursor.Y * TEXT_80_X_25_BYTES_PER_ROW) + (Cursor.X * &H2%)) + &H1%)
               CPU.Registers(SubRegisters8BitE.BH, NewValue:=ScrollAttribute)

               Select Case DirectCast(Character, TeletypeE)
                  Case TeletypeE.BEL
                     Task.Run(Sub() Console.Beep())
                  Case TeletypeE.BS
                     If Cursor.X > &H0% Then Cursor.X -= &H1%
                  Case TeletypeE.CR
                     Cursor.X = &H0%
                  Case TeletypeE.LF
                     If Cursor.Y < MCC.RowCount() - &H1% Then
                        Cursor.Y += &H1%
                     Else
                        VideoAdapter.ScrollBuffer(Up:=True, ScrollArea, Count:=&H1%)
                     End If
                  Case TeletypeE.TAB
                     Cursor.X = ((((Cursor.X + &H1%) \ &H8%) + &H1%) * &H8%) - &H1%
                     If Cursor.X >= MCC.ColumnCount() - &H1% Then
                        Cursor.X = &H0%
                        If Cursor.Y < MCC.RowCount() - &H1% Then
                           Cursor.Y += &H1%
                        Else
                           VideoAdapter.ScrollBuffer(Up:=True, ScrollArea, Count:=&H1%)
                        End If
                     End If
                  Case Else
                     CPU.Memory(VideoPageAddress + (Cursor.Y * TEXT_80_X_25_BYTES_PER_ROW) + (Cursor.X * &H2%)) = Character
                     If Attribute IsNot Nothing Then
                        CPU.Memory(VideoPageAddress + (Cursor.Y * TEXT_80_X_25_BYTES_PER_ROW) + (Cursor.X * &H2%) + &H1%) = CByte(Attribute.Value)
                     End If

                     If Cursor.X < MCC.ColumnCount() - &H1% Then
                        Cursor.X += &H1%
                     Else
                        Cursor.X = &H0%
                        If Cursor.Y < MCC.RowCount() - &H1% Then
                           Cursor.Y += &H1%
                        Else
                           VideoAdapter.ScrollBuffer(Up:=True, ScrollArea, Count:=&H1%)
                        End If
                     End If
               End Select
         End Select

         CPU.Memory(AddressesE.CursorPositions) = CByte(Cursor.X)
         CPU.Memory(AddressesE.CursorPositions + &H1%) = CByte(Cursor.Y)
         CursorPositionUpdate()
         CPU.Registers(SubRegisters8BitE.BH, NewValue:=VideoPage)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure updates the clock counter.
   Public Sub UpdateClockCounter()
      Try
         If ClockCounter = MAXIMUM_CLOCK_VALUE Then
            ClockCounter = &H0%
            CPU.Memory(AddressesE.ClockRollover) = &H1%
         Else
            ClockCounter += &H1%
            If ClockCounter = MAXIMUM_CLOCK_VALUE Then CPU.Memory(AddressesE.ClockRollover) = &HFF%
         End If

         CPU.PutWord(AddressesE.Clock, ClockCounter And &HFFFF%)
         CPU.PutWord(AddressesE.Clock + &H2%, ClockCounter >> &H10%)

         SyncLock SYNCHRONIZER
            CPU.DoSystemTimerTick = True
         End SyncLock
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure writes a string to the screen.
   Public Sub WriteString()
      Dim AL As Integer = CInt(CPU.Registers(SubRegisters8BitE.AL)) And &H3%
      Dim Attribute As New Byte
      Dim Character As New Byte
      Dim Column As Byte = CByte(CPU.Registers(SubRegisters8BitE.DL))
      Dim Count As Integer = CInt(CPU.Registers(Registers16BitE.CX))
      Dim HasAttributes As Boolean = ((AL = &H2% OrElse AL = &H3%))
      Dim MoveCursor As Boolean = ((AL = &H1% OrElse AL = &H3%))
      Dim Offset As Integer = CInt(CPU.Registers(Registers16BitE.BP))
      Dim PreviousColumn As New Byte
      Dim PreviousRow As New Byte
      Dim Row As Byte = CByte(CPU.Registers(SubRegisters8BitE.DH))
      Dim Segment As Integer = CInt(CPU.Registers(SegmentRegistersE.ES))
      Dim VideoPage As Integer = CInt(CPU.Registers(SubRegisters8BitE.BH))

      If Not HasAttributes Then
         Attribute = CByte(CPU.Registers(SubRegisters8BitE.BL))
      End If
      If Not MoveCursor Then
         PreviousColumn = CPU.Memory(AddressesE.CursorPositions + &H1%)
         PreviousRow = CPU.Memory(AddressesE.CursorPositions)
      End If
      If VideoPage >= MCC.VideoPageCount() Then
         VideoPage = &H0%
      End If

      CPU.Memory(AddressesE.CursorPositions) = Column
      CPU.Memory(AddressesE.CursorPositions + &H1%) = Row
      CursorPositionUpdate()

      Do While Count > &H0%
         Character = CPU.Memory((Segment << &H4%) + (Offset And &HFFFF%))
         If HasAttributes Then
            Attribute = CPU.Memory((Segment << &H4%) + ((Offset + &H1%) And &HFFFF%))
            Offset += &H2%
         Else
            Offset += &H1%
         End If
         TeleType(Character, Attribute)
         Count -= &H1%
      Loop

      If Not MoveCursor Then
         CPU.Memory(AddressesE.CursorPositions) = PreviousColumn
         CPU.Memory(AddressesE.CursorPositions + &H1%) = PreviousRow
         CursorPositionUpdate()
      End If
   End Sub
End Module