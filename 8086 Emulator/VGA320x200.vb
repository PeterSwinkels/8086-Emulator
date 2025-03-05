'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Drawing

'This class emulates VGA 320x200.
Public Class VGA320x200Class
   Implements VideoAdapterClass

   Private ReadOnly DEFAULT_PALETTE() As Integer = {&H0%, &HA8%, &HA800%, &HA8A8%, &HA80000%, &HA800A8%, &HA85400%, &HA8A8A8%, &H545454%, &H5454FC%, &H54FC54%, &H54FCFC%, &HFC5454%, &HFC54FC%, &HFCFC54%, &HFCFCFC%, &H0%, &H141414%, &H202020%, &H2C2C2C%, &H383838%, &H444444%, &H505050%, &H606060%, &H707070%, &H808080%, &H909090%, &HA0A0A0%, &HB4B4B4%, &HC8C8C8%, &HE0E0E0%, &HFCFCFC%, &HFC%, &H4000FC%, &H7C00FC%, &HBC00FC%, &HFC00FC%, &HFC00BC%, &HFC007C%, &HFC0040%, &HFC0000%, &HFC4000%, &HFC7C00%, &HFCBC00%, &HFCFC00%, &HBCFC00%, &H7CFC00%, &H40FC00%, &HFC00%, &HFC40%, &HFC7C%, &HFCBC%, &HFCFC%, &HBCFC%, &H7CFC%, &H40FC%, &H7C7CFC%, &H9C7CFC%, &HBC7CFC%, &HDC7CFC%, &HFC7CFC%, &HFC7CDC%, &HFC7CBC%, &HFC7C9C%, &HFC7C7C%, &HFC9C7C%, &HFCBC7C%, &HFCDC7C%, &HFCFC7C%, &HDCFC7C%, &HBCFC7C%, &H9CFC7C%, &H7CFC7C%, &H7CFC9C%, &H7CFCBC%, &H7CFCDC%, &H7CFCFC%, &H7CDCFC%, &H7CBCFC%, &H7C9CFC%, &HB4B4FC%, &HC4B4FC%, &HD8B4FC%, &HE8B4FC%, &HFCB4FC%, &HFCB4E8%, &HFCB4D8%, &HFCB4C4%, &HFCB4B4%, &HFCC4B4%, &HFCD8B4%, &HFCE8B4%, &HFCFCB4%, &HE8FCB4%, &HD8FCB4%, &HC4FCB4%, &HB4FCB4%, &HB4FCC4%, &HB4FCD8%, &HB4FCE8%, &HB4FCFC%, &HB4E8FC%, &HB4D8FC%, &HB4C4FC%, &H70%, &H1C0070%, &H380070%, &H540070%, &H700070%, &H700054%, &H700038%, &H70001C%, &H700000%, &H701C00%, &H703800%, &H705400%, &H707000%, &H547000%, &H387000%, &H1C7000%, &H7000%, &H701C%, &H7038%, &H7054%, &H7070%, &H5470%, &H3870%, &H1C70%, &H383870%, &H443870%, &H543870%, &H603870%, &H703870%, &H703860%, &H703854%, &H703844%, &H703838%, &H704438%, &H705438%, &H706038%, &H707038%, &H607038%, &H547038%, &H447038%, &H387038%, &H387044%, &H387054%, &H387060%, &H387070%, &H386070%, &H385470%, &H384470%, &H505070%, &H585070%, &H605070%, &H685070%, &H705070%, &H705068%, &H705060%, &H705058%, &H705050%, &H705850%, &H706050%, &H706850%, &H707050%, &H687050%, &H607050%, &H587050%, &H507050%, &H507058%, &H507060%, &H507068%, &H507070%, &H506870%, &H506070%, &H505870%, &H40%, &H100040%, &H200040%, &H300040%, &H400040%, &H400030%, &H400020%, &H400010%, &H400000%, &H401000%, &H402000%, &H403000%, &H404000%, &H304000%, &H204000%, &H104000%, &H4000%, &H4010%, &H4020%, &H4030%, &H4040%, &H3040%, &H2040%, &H1040%, &H202040%, &H282040%, &H302040%, &H382040%, &H402040%, &H402038%, &H402030%, &H402028%, &H402020%, &H402820%, &H403020%, &H403820%, &H404020%, &H384020%, &H304020%, &H284020%, &H204020%, &H204028%, &H204030%, &H204038%, &H204040%, &H203840%, &H203040%, &H202840%, &H2C2C40%, &H302C40%, &H342C40%, &H3C2C40%, &H402C40%, &H402C3C%, &H402C34%, &H402C30%, &H402C2C%, &H40302C%, &H40342C%, &H403C2C%, &H40402C%, &H3C402C%, &H34402C%, &H30402C%, &H2C402C%, &H2C4030%, &H2C4034%, &H2C403C%, &H2C4040%, &H2C3C40%, &H2C3440%, &H2C3040%, &H0%, &H0%, &H0%, &H0%, &H0%, &H0%, &H0%, &H0%}  'Defines the default palette.

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte, CodePage() As Integer) Implements VideoAdapterClass.Display
      For y As Integer = 0 To ScreenSize.Height - 1
         For x As Integer = 0 To ScreenSize.Width - 1
            With DirectCast(Screen, Bitmap)
               .SetPixel(x, y, Color.FromArgb(DEFAULT_PALETTE(Memory(AddressesE.VGA320x200 + ((y * ScreenSize.Width) + x))) Or &HFF000000%))
            End With
         Next x
      Next y
   End Sub

   'This procedure returns the screen size used by a video adapter.
   Public Function ScreenSize() As Size Implements VideoAdapterClass.ScreenSize
      Return New Size(320, 200)
   End Function
End Class
