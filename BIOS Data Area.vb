'This modules's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Drawing

'This module contains BIOS data area related information.
Public Module BIOSDataAreaClass
   'This enumeration lists the video modes.
   Public Enum VideoModesE As Byte
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
      CGA320x200 = &HB8000%        '320x200 CGA video buffer.
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

   Public Const CGA_320_x_200_BUFFER_SIZE As Integer = &H3E80%      'Defines the 320x200 CGA mode video memory's size.
   Public Const MAXIMUM_VIDEO_PAGE_COUNT As Integer = &H8%          'Defines the maximum number of video pages.
   Public Const TEXT_80_X_25_BYTES_PER_ROW As Integer = &HA0%       'Defines the number of bytes per row used by 80x25 monochrome text mode 
   Public Const TEXT_80_X_25_COLUMN_COUNT As Integer = &H50%        'Defines the number of columns used by 80x25 monochrome text mode 
   Public Const TEXT_80_X_25_LINE_COUNT As Integer = &H19%          'Defines the number of lines used by 80x25 monochrome text mode 
   Public Const TEXT_80_X_25_MONO_BUFFER_SIZE As Integer = &HFA0%   'Defines the 80x25 monochrome text mode video memory's size.
   Public Const VGA_320_x_200_BUFFER_SIZE As Integer = &HFA00%      'Defines the 320x200 VGA mode video memory's size.

   Public ReadOnly BACKGROUND_COLORS() As Color = {Color.Black, Color.DarkBlue, Color.DarkGreen, Color.DarkCyan, Color.DarkRed, Color.Purple, Color.Brown, Color.White, Color.Gray, Color.Blue, Color.Green, Color.Cyan, Color.Red, Color.Pink, Color.Yellow, Color.White}   'Contains the background colors.
   Public ReadOnly PALETTE0() As Color = {Nothing, Color.DarkGreen, Color.DarkRed, Color.Brown}                                                                                                                                                                              'Contains palette #0 colors.
   Public ReadOnly PALETTE1() As Color = {Nothing, Color.DarkCyan, Color.Purple, Color.White}                                                                                                                                                                                'Contains palette #1 colors.

   Public BackgroundColor As Color = Color.Black   'Contains the current background color.
   Public Palette() As Color = Nothing             'Contains the current palette.
End Module