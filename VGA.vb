'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Drawing
Imports System.Linq

'This class contains the VGA related procedures.
Public Class VGAClass
   Public STATIC_FUNCTIONALITY_ADDRESS As Integer = &HC2700%   'Defines the static video functionality table address.
   Private DYNAMIC_FUNCTIONALITY_SIZE As Integer = &H40%        'Defines the dynamic video functionality table buffer size.

   Public ReadOnly STATIC_FUNCTIONALITY() As Byte = {&HFF%, &HFF%, &HF%, &H0%, &H0%, &H0%, &H0%, &H7%, &H4%, &H2%, &HFF%, &HE%, &H0%, &H0%, &H0%, &H0%}   'Defines the static functionality table.
   Public ReadOnly VGA_DEFAULT_PALETTE() As Integer = {&H0%, &HA8%, &HA800%, &HA8A8%, &HA80000%, &HA800A8%, &HA85400%, &HA8A8A8%, &H545454%, &H5454FC%, &H54FC54%, &H54FCFC%, &HFC5454%, &HFC54FC%, &HFCFC54%, &HFCFCFC%, &H0%, &H141414%, &H202020%, &H2C2C2C%, &H383838%, &H444444%, &H505050%, &H606060%, &H707070%, &H808080%, &H909090%, &HA0A0A0%, &HB4B4B4%, &HC8C8C8%, &HE0E0E0%, &HFCFCFC%, &HFC%, &H4000FC%, &H7C00FC%, &HBC00FC%, &HFC00FC%, &HFC00BC%, &HFC007C%, &HFC0040%, &HFC0000%, &HFC4000%, &HFC7C00%, &HFCBC00%, &HFCFC00%, &HBCFC00%, &H7CFC00%, &H40FC00%, &HFC00%, &HFC40%, &HFC7C%, &HFCBC%, &HFCFC%, &HBCFC%, &H7CFC%, &H40FC%, &H7C7CFC%, &H9C7CFC%, &HBC7CFC%, &HDC7CFC%, &HFC7CFC%, &HFC7CDC%, &HFC7CBC%, &HFC7C9C%, &HFC7C7C%, &HFC9C7C%, &HFCBC7C%, &HFCDC7C%, &HFCFC7C%, &HDCFC7C%, &HBCFC7C%, &H9CFC7C%, &H7CFC7C%, &H7CFC9C%, &H7CFCBC%, &H7CFCDC%, &H7CFCFC%, &H7CDCFC%, &H7CBCFC%, &H7C9CFC%, &HB4B4FC%, &HC4B4FC%, &HD8B4FC%, &HE8B4FC%, &HFCB4FC%, &HFCB4E8%, &HFCB4D8%, &HFCB4C4%, &HFCB4B4%, &HFCC4B4%, &HFCD8B4%, &HFCE8B4%, &HFCFCB4%, &HE8FCB4%, &HD8FCB4%, &HC4FCB4%, &HB4FCB4%, &HB4FCC4%, &HB4FCD8%, &HB4FCE8%, &HB4FCFC%, &HB4E8FC%, &HB4D8FC%, &HB4C4FC%, &H70%, &H1C0070%, &H380070%, &H540070%, &H700070%, &H700054%, &H700038%, &H70001C%, &H700000%, &H701C00%, &H703800%, &H705400%, &H707000%, &H547000%, &H387000%, &H1C7000%, &H7000%, &H701C%, &H7038%, &H7054%, &H7070%, &H5470%, &H3870%, &H1C70%, &H383870%, &H443870%, &H543870%, &H603870%, &H703870%, &H703860%, &H703854%, &H703844%, &H703838%, &H704438%, &H705438%, &H706038%, &H707038%, &H607038%, &H547038%, &H447038%, &H387038%, &H387044%, &H387054%, &H387060%, &H387070%, &H386070%, &H385470%, &H384470%, &H505070%, &H585070%, &H605070%, &H685070%, &H705070%, &H705068%, &H705060%, &H705058%, &H705050%, &H705850%, &H706050%, &H706850%, &H707050%, &H687050%, &H607050%, &H587050%, &H507050%, &H507058%, &H507060%, &H507068%, &H507070%, &H506870%, &H506070%, &H505870%, &H40%, &H100040%, &H200040%, &H300040%, &H400040%, &H400030%, &H400020%, &H400010%, &H400000%, &H401000%, &H402000%, &H403000%, &H404000%, &H304000%, &H204000%, &H104000%, &H4000%, &H4010%, &H4020%, &H4030%, &H4040%, &H3040%, &H2040%, &H1040%, &H202040%, &H282040%, &H302040%, &H382040%, &H402040%, &H402038%, &H402030%, &H402028%, &H402020%, &H402820%, &H403020%, &H403820%, &H404020%, &H384020%, &H304020%, &H284020%, &H204020%, &H204028%, &H204030%, &H204038%, &H204040%, &H203840%, &H203040%, &H202840%, &H2C2C40%, &H302C40%, &H342C40%, &H3C2C40%, &H402C40%, &H402C3C%, &H402C34%, &H402C30%, &H402C2C%, &H40302C%, &H40342C%, &H403C2C%, &H40402C%, &H3C402C%, &H34402C%, &H30402C%, &H2C402C%, &H2C4030%, &H2C4034%, &H2C403C%, &H2C4040%, &H2C3C40%, &H2C3440%, &H2C3040%, &H0%, &H0%, &H0%, &H0%, &H0%, &H0%, &H0%, &H0%}  'Defines the VGA default palette.
   Private ReadOnly TO_8_BIT As Double = (&HFF% / &H3F%)   'Defines the value used to convert 6-bit values to 8-bit values.

   Public VGAPaintBrushes(&H0% To &HFF%) As Brush   'Contains the paint brushes created using the palette.

   Private SelectedAddress As New Integer   'Contains the VGA video DAC PEL address.

   'This procedure returns the dynamic video functionality table.
   Public Function GetDynamicFunctionality() As Byte()
      Dim Offset As UShort = CUShort(STATIC_FUNCTIONALITY_ADDRESS And &HFFFF%)
      Dim Segment As UShort = CUShort((STATIC_FUNCTIONALITY_ADDRESS And &HF0000%) >> &H4%)
      Dim Table As New List(Of Byte)

      With Table
         .AddRange(BitConverter.GetBytes(Offset))
         .AddRange(BitConverter.GetBytes(Segment))
         .Add(CurrentVideoMode)
         .AddRange(BitConverter.GetBytes(CUShort(MCC.ColumnCount())))
         .AddRange(BitConverter.GetBytes(CUShort({&H4000%, &H4000%, &H4000%, &H2000%, &H4000%, &H8000%, &H8000%, Nothing, Nothing, Nothing, &H800%, &H800%, &H1000%, &H1000%, Nothing, &H2000%, &HA000%, &HA000%}(CurrentVideoMode))))
         .AddRange({&H0%, &H0%})
         For VideoPage As Integer = &H0% To MAXIMUM_VIDEO_PAGE_COUNT - &H1%
            .AddRange(BitConverter.GetBytes(CUShort(CPU.GET_WORD(AddressesE.CursorPositions + (VideoPage * &H2%)))))
         Next VideoPage
         .Add(CPU.Memory(AddressesE.CursorScanLines))
         .Add(CPU.Memory(AddressesE.CursorScanLines + &H1%))
         .Add(CPU.Memory(AddressesE.VideoPage))
         .AddRange(BitConverter.GetBytes(CUShort(IOPortsE.CGAIndex)))
         .Add(ToByte(ReadIOPort(IOPortsE.CGAMode).Value))
         .Add(ToByte(ReadIOPort(IOPortsE.CGAColor).Value))
         .Add(MCC.RowCount())
         .AddRange(BitConverter.GetBytes(CUShort(MCC.CharacterScanLineCount())))
         .Add(&H0%)
         .Add(&H0%)
         .AddRange(BitConverter.GetBytes(CUShort(MCC.ColorCount())))
         .Add(MCC.VideoPageCount())
         .Add(MCC.RasterScanLineCount())
         .Add(&H0%)
         .Add(&H0%)
         .Add(&H0%)
         .AddRange({&H0%, &H0%, &H0%})
         .Add(MCCClass.VIDEO_RAM_256KB)
         .Add(&H0%)
         .AddRange({&H0%, &H0%, &H0%, &H0%})
      End With

      Table.AddRange(Enumerable.Repeat(CByte(&H0%), DYNAMIC_FUNCTIONALITY_SIZE - Table.Count).ToArray())

      Return Table.ToArray()
   End Function

   'This procedure selects the specified VGA video DAC PEL address.
   Public Sub SelectAddress(NewAddress As Integer)
      SelectedAddress = NewAddress And &HFF%
   End Sub

   'This procedure writes to the VGA video DAC.
   Public Sub WriteToDAC(Value As Integer)
      Static Values As New List(Of Integer)

      Values.Add(CInt((Value And &H3F%) * TO_8_BIT))

      If Values.Count = &H3% Then
         VGA_DEFAULT_PALETTE(SelectedAddress) = (Values(&H0%) << &H10%) Or (Values(&H1%) << &H8%) Or Values(&H2%)
         VGA.VGAPaintBrushes(SelectedAddress) = New SolidBrush(Color.FromArgb(VGA_DEFAULT_PALETTE(SelectedAddress) Or &HFF000000%))
         Values.Clear()
      End If
   End Sub
End Class
