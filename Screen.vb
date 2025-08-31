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
   'This procedure initializes this window.
   Public Sub New()
      Try
         InitializeComponent()

         ToolTip.SetToolTip(Me, "This is where the emulation’s screen output is displayed and keyboard input is accepted.")

         ScreenActive = True
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure sets the screen window's state to inactive when it is closed.
   Private Sub ScreenWindow_Closed(sender As Object, e As EventArgs) Handles MyBase.Closed
      Try
         ScreenActive = False
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

   'This procedure handles any keystrokes made by the user when a key is released.
   Private Sub ScreenWindow_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
      Try
         GetBIOSKeyCode(e)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure gives the video adapter the command to update the screen's content.
   Private Sub ScreenWindow_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
      Try
         Dim VideoMode As VideoModesE = DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)

         Me.Text = $"Screen {If([Enum].IsDefined(GetType(VideoModesE), VideoMode), $"{VideoMode} (0x{CInt(VideoMode).ToString("X")})", "Unknown mode.")}"

         If VideoAdapter IsNot Nothing Then
            Me.ClientSize = VideoAdapter.Resolution

            If Me.BackgroundImage Is Nothing Then
               Me.BackgroundImage = New Bitmap(Me.ClientSize.Width, Me.ClientSize.Height)
               Graphics.FromImage(Me.BackgroundImage).Clear(Color.Black)
            End If

            VideoAdapter.Display(Me.BackgroundImage, CPU.Memory, CODE_PAGE_437)
         Else
            Me.ClientSize = New Size(320, 200)
            Me.BackColor = Color.Black
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Class