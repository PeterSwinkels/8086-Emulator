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
   'This procedure initalizes this window.
   Public Sub New()
      Try
         Dim VideoMode As VideoModesE = DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)

         InitializeComponent()

         Me.Text = $"Screen {If([Enum].IsDefined(GetType(VideoModesE), VideoMode), $"{VideoMode} (0x{CInt(VideoMode).ToString("X")})", "Unknown mode.")}"
         Me.ClientSize = If(VideoAdapter Is Nothing, New Size(320, 200), VideoAdapter.Resolution)

         If Me.BackgroundImage Is Nothing Then
            Me.BackgroundImage = New Bitmap(Me.ClientSize.Width, Me.ClientSize.Height)
            Graphics.FromImage(Me.BackgroundImage).Clear(Color.Black)
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
         If VideoAdapter IsNot Nothing Then
            VideoAdapter.Display(Me.BackgroundImage, CPU.Memory, CODE_PAGE_437)
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Class