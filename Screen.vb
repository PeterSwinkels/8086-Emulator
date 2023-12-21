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
   'This enumeration lists the supported video modes.
   Private Enum VideoModesE As Byte
      LowResolutionVGA = &H13%   'Defines Low resolution VGA.
      MDA = &H7%                 'Defines MDA.
   End Enum

   Private Const VIDEO_MODE_ADDRESS As Integer = &H449%    'Defines the memory address where the current video mode is stored.

   Private VideoAdapter As VideoAdapterClass = New MDAClass  'Contains a reference to the video adapter used.

   Private WithEvents Updater As New Timer With {.Enabled = True, .Interval = 56}   'Contains the screen updater.

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

   'This procedure gives the video adapter the command to update the screen's content.
   Private Sub ScreenWindow_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
      Try
         VideoAdapter.Display(Me.BackgroundImage, CPU.Memory)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure updates the screen output.
   Private Sub Updater_Tick(sender As Object, e As EventArgs) Handles Updater.Tick
      Static CurrentVideoMode As VideoModesE = DirectCast(CPU.Memory(VIDEO_MODE_ADDRESS), VideoModesE)

      Try
         If DirectCast(CPU.Memory(VIDEO_MODE_ADDRESS), VideoModesE) = CurrentVideoMode Then
            Me.Invalidate()
         Else
            CurrentVideoMode = DirectCast(CPU.Memory(VIDEO_MODE_ADDRESS), VideoModesE)
            Select Case CurrentVideoMode
               Case VideoModesE.LowResolutionVGA
                  VideoAdapter = New LowResolutionVGA
               Case VideoModesE.MDA
                  VideoAdapter = New MDAClass
               Case Else
                  VideoAdapter = New MDAClass
            End Select

            Me.ClientSize = VideoAdapter.ScreenSize
         End If

      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Class