'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

'This class contains the screen output window.
Public Class ScreenWindow
   Private ReadOnly CODE_PAGE_437() As Integer = {&H0%, &H263A%, &H263B%, &H2665%, &H2666%, &H2663%, &H2660%, &H2022%, &H25D8%, &H25CB%, &H25D9%, &H2642%, &H2640%, &H266A%, &H266B%, &H263C%, &H25BA%, &H25C4%, &H2195%, &H203C%, &HB6%, &HA7%, &H25AC%, &H21A8%, &H2191%, &H2193%, &H2192%, &H2190%, &H221F%, &H2194%, &H25B2%, &H25BC%, &H20%, &H21%, &H22%, &H23%, &H24%, &H25%, &H26%, &H27%, &H28%, &H29%, &H2A%, &H2B%, &H2C%, &H2D%, &H2E%, &H2F%, &H30%, &H31%, &H32%, &H33%, &H34%, &H35%, &H36%, &H37%, &H38%, &H39%, &H3A%, &H3B%, &H3C%, &H3D%, &H3E%, &H3F%, &H40%, &H41%, &H42%, &H43%, &H44%, &H45%, &H46%, &H47%, &H48%, &H49%, &H4A%, &H4B%, &H4C%, &H4D%, &H4E%, &H4F%, &H50%, &H51%, &H52%, &H53%, &H54%, &H55%, &H56%, &H57%, &H58%, &H59%, &H5A%, &H5B%, &H5C%, &H5D%, &H5E%, &H5F%, &H60%, &H61%, &H62%, &H63%, &H64%, &H65%, &H66%, &H67%, &H68%, &H69%, &H6A%, &H6B%, &H6C%, &H6D%, &H6E%, &H6F%, &H70%, &H71%, &H72%, &H73%, &H74%, &H75%, &H76%, &H77%, &H78%, &H79%, &H7A%, &H7B%, &H7C%, &H7D%, &H7E%, &H2302%, &HC7%, &HFC%, &HE9%, &HE2%, &HE4%, &HE0%, &HE5%, &HE7%, &HEA%, &HEB%, &HE8%, &HEF%, &HEE%, &HEC%, &HC4%, &HC5%, &HC9%, &HE6%, &HC6%, &HF4%, &HF6%, &HF2%, &HFB%, &HF9%, &HFF%, &HD6%, &HDC%, &HA2%, &HA3%, &HA5%, &H20A7%, &H192%, &HE1%, &HED%, &HF3%, &HFA%, &HF1%, &HD1%, &HAA%, &HBA%, &HBF%, &H2310%, &HAC%, &HBD%, &HBC%, &HA1%, &HAB%, &HBB%, &H2591%, &H2592%, &H2593%, &H2502%, &H2524%, &H2561%, &H2562%, &H2556%, &H2555%, &H2563%, &H2551%, &H2557%, &H255D%, &H255C%, &H255B%, &H2510%, &H2514%, &H2534%, &H252C%, &H251C%, &H2500%, &H253C%, &H255E%, &H255F%, &H255A%, &H2554%, &H2569%, &H2566%, &H2560%, &H2550%, &H256C%, &H2567%, &H2568%, &H2564%, &H2565%, &H2559%, &H2558%, &H2552%, &H2553%, &H256B%, &H256A%, &H2518%, &H250C%, &H2588%, &H2584%, &H258C%, &H2590%, &H2580%, &H3B1%, &HDF%, &H393%, &H3C0%, &H3A3%, &H3C3%, &HB5%, &H3C4%, &H3A6%, &H398%, &H3A9%, &H3B4%, &H221E%, &H3C6%, &H3B5%, &H2229%, &H2261%, &HB1%, &H2265%, &H2264%, &H2320%, &H2321%, &HF7%, &H2248%, &HB0%, &H2219%, &HB7%, &H221A%, &H207F%, &HB2%, &H25A0%, &HA0%}  'Contains the code page 437 to unicode mappings.

   Private VideoAdapter As VideoAdapterClass = New Text80x25MonoClass   'Contains a reference to the video adapter used.

   Private WithEvents ScreenRefresh As New Timer With {.Enabled = True, .Interval = 56}  'Contains the screen refresh timer.

   'This procedure initalizes this window.
   Public Sub New()
      Try
         InitializeComponent()

         Me.ClientSize = VideoAdapter.ScreenSize

         Me.BackgroundImage = New Bitmap(Me.ClientSize.Width, Me.ClientSize.Height)
         Graphics.FromImage(Me.BackgroundImage).Clear(Color.Black)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure refreshes the screen's output.
   Private Sub ScreenRefresh_Tick(sender As Object, e As EventArgs) Handles ScreenRefresh.Tick
      Try
         Dim VideoMode As VideoModesE = DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
         Static CurrentVideoMode As VideoModesE = Nothing

         If VideoMode = CurrentVideoMode Then
            Me.Invalidate()
         Else
            CurrentVideoMode = VideoMode

            If [Enum].IsDefined(GetType(VideoModesE), CByte(CurrentVideoMode)) Then
               Me.Text = $"Screen - {CurrentVideoMode} (0x{CInt(CurrentVideoMode).ToString("X")})"
            Else
               Me.Text = "Screen - Unknown mode."
            End If

            VideoMode = DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)
            Select Case VideoMode
               Case VideoModesE.VGA320x200
                  VideoAdapter = New VGA320x200Class
               Case VideoModesE.Text80x25Mono
                  VideoAdapter = New Text80x25MonoClass
               Case Else
                  VideoAdapter = New Text80x25MonoClass
            End Select

            Me.ClientSize = VideoAdapter.ScreenSize
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles any keystrokes made by the user when a key is pressed.
   Private Sub ScreenWindow_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
      Try
         Dim BIOSKeyCode As Integer? = GetBIOSKeyCode(e)

         If BIOSKeyCode IsNot Nothing Then LastBIOSKeyCode(NewLastBIOSKeyCode:=CInt(BIOSKeyCode))
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure gives the video adapter the command to update the screen's content.
   Private Sub ScreenWindow_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
      Try
         VideoAdapter.Display(Me.BackgroundImage, CPU.Memory, CODE_PAGE_437)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Class