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
   Private ReadOnly VideoAdapter As New MDAClass   'Contains a reference to the video adapter used.

   Private WithEvents Refresher As New Timer With {.Enabled = True, .Interval = 56}   'Contains the screen refresher.

   'This procedure initalizes this window.
   Public Sub New()
      Try
         InitializeComponent()

         Me.ClientSize = VideoAdapter.SCREEN_SIZE

         Me.BackgroundImage = New Bitmap(Me.ClientSize.Width, Me.ClientSize.Height)
         Graphics.FromImage(Me.BackgroundImage).Clear(Color.Black)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure refreshes the screen output.
   Private Sub Refresher_Tick(sender As Object, e As EventArgs) Handles Refresher.Tick
      Try
         Me.Invalidate()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure gives the video adapter the command to update the screen's content.
   Private Sub ScreenWindow_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
      Try
         VideoAdapter.Display(Me.BackgroundImage, CPU.Memory)

      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Class