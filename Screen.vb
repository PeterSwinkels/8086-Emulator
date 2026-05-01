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
   Private Const WM_KEYDOWN As Integer = &H100%   'Defines the key down window message.
   Private Const WM_KEYUP As Integer = &H101%     'Defines the key up window message.

   'This procedure initializes this window.
   Public Sub New()
      Try
         InitializeComponent()

         ToolTip.SetToolTip(Me, "This is where the emulation’s screen output is displayed and keyboard input is accepted.")
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
         Dim VideoMode As VideoModesE = VideoAdapterToVideoMode()

         Me.Text = $"Screen {If([Enum].IsDefined(GetType(VideoModesE), VideoMode), $"{VideoMode} (0x{CInt(VideoMode).ToString("X")})", "Unknown mode.")}"

         If VideoAdapter Is Nothing Then
            Me.ClientSize = New Size(320, 200)
            Me.BackColor = Color.Black
         Else
            Me.ClientSize = VideoAdapter.Resolution

            If Me.BackgroundImage Is Nothing Then
               Me.BackgroundImage = New Bitmap(Me.ClientSize.Width, Me.ClientSize.Height)
               Graphics.FromImage(Me.BackgroundImage).Clear(Color.Black)
            End If

            VideoAdapter.Display(Me.BackgroundImage, CPU.Memory, CODE_PAGE_437)
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure sets the key scan code to be returned by the keyboard I/O port.
   Protected Overrides Sub WndProc(ByRef m As Message)
      Try
         Dim Scancode As Integer = (m.LParam.ToInt32 >> &H10%) And &HFF%

         Select Case m.Msg
            Case WM_KEYDOWN
               KeyScancode = CByte(Scancode)
            Case WM_KEYUP
               KeyScancode = CByte(Scancode Or &H80%)
         End Select

         MyBase.WndProc(m)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Class