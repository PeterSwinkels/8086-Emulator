'This modules's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Drawing

'This module contains BIOS related procedures.
Public Module BIOSModule
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

   'This enumeration lists the flat addresses of BIOS data area locations.
   Public Enum AddressesE As Integer
      BIOS = &HF0000%              'BIOS.
      BIOSStack = &HF30000%        'BIOS stack segment.
      CGA320x200 = &HB8000%        '320x200 CGA video buffer.
      Clock = &H46C%               'Clock.
      ClockRollover = &H470%       'Clock rollover flag.
      CursorPositions = &H450%     'Cursor positions.
      CursorScanLines = &H460%     'Cursor scan line start/end.
      EquipmentFlags = &H410%      'Equipment flags.
      MemorySize = &H413%          'Memory size.
      Text80x25Mono = &HB0000%     '80x25 monochrome text video buffer.
      VGA320x200 = &HA0000%        '320x200 VGA video buffer.
      VideoMode = &H449%           'Current video mode.
      VideoModeOptions = &H487%    'Video mode options.
      VideoPage = &H462%           'Current video page.
   End Enum

   'This enumeration lists the supported Teletype control characters
   Public Enum TeletypeE As Byte
      BEL = &H7%    'Bell.
      BS = &H8%     'Backspace.
      TAB = &H9%    'Tab.
      LF = &HA%     'Line feed.
      CR = &HD%     'Carriage return.
   End Enum

   Public Const CGA_320_X_200_BYTES_PER_ROW As Integer = &H50%               'Defines the number of bytes per row used by 320x200 CGA mode 
   Public Const CGA_320_X_200_LINES_PER_CHARACTER As Integer = &H8%          'Defines the number of lines per character used by 320x200 CGA mode 
   Public Const CGA_320_X_200_LINE_COUNT As Integer = &H19%                  'Defines the number of lines used by 320x200 CGA mode 
   Public Const CGA_320_x_200_BUFFER_SIZE As Integer = &H4000%               'Defines the 320x200 CGA mode video memory's size.
   Public Const INITIAL_STACK_SIZE As Integer = &H10%                        'Defines the stack size at start up.
   Public Const EQUIPMENT_FLAGS As Integer = &H30%                           'Defines the equipment flags.
   Public Const MAXIMUM_CLOCK_VALUE As Integer = &H1800B0%                   'Defines the highest value reached by the clock just before midnight.
   Public Const MAXIMUM_VIDEO_PAGE_COUNT As Integer = &H8%                   'Defines the maximum number of video pages.
   Public Const TEXT_80_X_25_BYTES_PER_ROW As Integer = &HA0%                'Defines the number of bytes per row used by 80x25 monochrome text mode 
   Public Const TEXT_80_X_25_COLUMN_COUNT As Integer = &H50%                 'Defines the number of columns used by 80x25 monochrome text mode 
   Public Const TEXT_80_X_25_LINE_COUNT As Integer = &H19%                   'Defines the number of lines used by 80x25 monochrome text mode 
   Public Const TEXT_80_X_25_MONO_BUFFER_SIZE As Integer = &HFA0%            'Defines the 80x25 monochrome text mode video memory's size.
   Public Const VGA_320_x_200_BUFFER_SIZE As Integer = &HFA00%               'Defines the 320x200 VGA mode video memory's size.
   Public Const VGA_320_X_200_BYTES_PER_ROW As Integer = &H140%              'Defines the number of bytes per row used by 320x200 VGA mode 
   Public Const VGA_320_X_200_LINE_COUNT As Integer = &H19%                  'Defines the number of lines used by 320x200 VGA mode 
   Public Const VGA_320_X_200_PIXELS_PER_CHARACTER_SIDE As Integer = &H8%    'Defines the number of per character side used by 320x200 VGA mode 

   Public ReadOnly BACKGROUND_COLORS() As Color = {Color.Black, Color.DarkBlue, Color.DarkGreen, Color.DarkCyan, Color.DarkRed, Color.Purple, Color.Brown, Color.White, Color.Gray, Color.Blue, Color.Green, Color.Cyan, Color.Red, Color.Pink, Color.Yellow, Color.White}   'Contains the background colors.
   Public ReadOnly PALETTE0() As Color = {Nothing, Color.DarkGreen, Color.DarkRed, Color.Brown}                                                                                                                                                                              'Contains palette #0 colors.
   Public ReadOnly PALETTE1() As Color = {Nothing, Color.DarkCyan, Color.Purple, Color.White}                                                                                                                                                                                'Contains palette #1 colors.

   Public BackgroundColor As Color = Color.Black   'Contains the current background color.
   Public ClockCounter As Integer = &H0%           'Contains the system clock counter.
   Public Palette() As Color = Nothing             'Contains the current palette.

   'This procedure loads the BIOS into memory.
   Public Sub LoadBIOS()
      Try
         Dim Address As New Integer
         Dim Offset As Integer = AddressesE.BIOS

         For Interrupt As Integer = &H0% To &H21%
            Address = Interrupt * &H4%
            CPU.PutWord(Address + &H2%, AddressesE.BIOS >> &H4%)
            CPU.PutWord(Address, Offset)
            Offset = WriteBytesToMemory({CPU8086Class.OpcodesE.EXT_INT, CByte(Interrupt), CPU8086Class.OpcodesE.IRET}, Offset)
         Next Interrupt

         CPU.PutWord(AddressesE.EquipmentFlags, EQUIPMENT_FLAGS)
         CPU.PutWord(AddressesE.MemorySize, (CPU.Memory.Length \ &H400%))
         CPU.Memory(AddressesE.VideoMode) = VideoModesE.Text80x25Mono

         CPU.Registers(CPU8086Class.SegmentRegistersE.SS, NewValue:=AddressesE.BIOSStack)
         CPU.Registers(CPU8086Class.Registers16BitE.SP, NewValue:=INITIAL_STACK_SIZE)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure updates the clock counter.
   Public Sub UpdateClockCounter()
      Try
         If ClockCounter = MAXIMUM_CLOCK_VALUE Then
            ClockCounter = &H0%
            CPU.Memory(AddressesE.ClockRollover) = &H0%
         Else
            ClockCounter += &H1%
            If ClockCounter = MAXIMUM_CLOCK_VALUE Then CPU.Memory(AddressesE.ClockRollover) = &HFF%
         End If

         CPU.PutWord(AddressesE.Clock, ClockCounter And &HFF%)
         CPU.PutWord(AddressesE.Clock + &H2%, ClockCounter And &HFF00%)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Module