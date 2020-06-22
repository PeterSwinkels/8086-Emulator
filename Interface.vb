'This class' imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.ComponentModel
Imports System.Environment
Imports System.Linq
Imports System.Windows.Forms

'This class contains this program's main interface window.
Public Class InterfaceWindow

   'This procedure initializes this window.
   Public Sub New()
      Try
         InitializeComponent()

         With Me
            .Height = CInt(My.Computer.Screen.WorkingArea.Height / 2)
            .Text = $"{My.Application.Info.Title} v{My.Application.Info.Version} - by: {My.Application.Info.CompanyName}"
            .Width = CInt(My.Computer.Screen.WorkingArea.Width / 2)
         End With

         ToolTip.SetToolTip(CommandBox, "Enter commands here. Specify ""?"" for help.")
         ToolTip.SetToolTip(EnterButton, "Sends a commmand to the CPU emulator.")
         ToolTip.SetToolTip(OutputBox, "This box displays CPU memory contents, register values, etc.")

         Output = Me.OutputBox
         UpdateStatus()

         If GetCommandLineArgs.Count > 1 Then RunCommandScript(GetCommandLineArgs.Last)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure gives the command to parse the user's input.
   Private Sub EnterButton_Click(sender As Object, e As EventArgs) Handles EnterButton.Click
      Try
         MouseCursorStatus(Busy:=True)

         If AssemblyModeOn Then
            Assemble(CommandBox.Text)
         Else
            ParseCommand(CommandBox.Text)
         End If

         CommandBox.Clear()
         UpdateStatus()
         MouseCursorStatus(Busy:=False)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure moves the focus to the command box when this window gets the focus.
   Private Sub InterfaceWindow_GotFocus(sender As Object, e As EventArgs) Handles MyBase.GotFocus
      Try
         CommandBox.Select()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure ensures the emulated CPU is stopped and the output to the interface stopped when this window closes.
   Private Sub InterfaceWindow_Closing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing
      Try
         CPU.Clock.Stop()
         Tracer.Stop()
         Output = Nothing
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure responds to changes in the output.
   Private Sub OutputBox_TextChanged(sender As Object, e As EventArgs) Handles OutputBox.TextChanged
      Try
         UpdateStatus()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure updates the status bar.
   Private Sub UpdateStatus()
      Try
         CPUActiveLabel.Text = $"CPU {If(CPU.Clock.Enabled OrElse Tracer.Enabled, "active.", "halted.")}"
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Class