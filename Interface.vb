'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.ComponentModel
Imports System.Environment
Imports System.Linq
Imports System.Threading.Tasks
Imports System.Windows.Forms

'This class contains this program's main interface window.
Public Class InterfaceWindow
   Private LastCommand As String = Nothing   'Contains the last command entered by the user.

   Private WithEvents Updater As New Timer With {.Enabled = True, .Interval = 56}   'Contains the interface's updater.

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

         CursorBlink = New Timer With {.Enabled = True, .Interval = 500}

         Output = Me.OutputBox

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

         If Not CommandBox.Text.Trim() = Nothing Then LastCommand = CommandBox.Text

         CommandBox.Clear()
         MouseCursorStatus(Busy:=False)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure moves the focus to the command box when this window is activated.
   Private Sub InterfaceWindow_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
      Try
         CommandBox.Select()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure ensures the emulated CPU is stopped and the output to the interface stopped when this window closes.
   Private Sub InterfaceWindow_Closing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing
      Try
         CPU.ClockToken.Cancel()
         Output = Nothing

         Updater.Stop()

         ScreenWindow.Close()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles the user's keystrokes.
   Private Sub InterfaceWindow_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
      Try
         If e.KeyCode = Keys.F3 AndAlso Not LastCommand = Nothing Then
            CommandBox.Text = LastCommand
            CommandBox.Select(CommandBox.Text.Length, 0)
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure gives the command to load the file dropped into the data box.
   Private Sub OutputBox_DragDrop(sender As Object, e As DragEventArgs) Handles OutputBox.DragDrop
      Try
         If e.Data.GetDataPresent(DataFormats.FileDrop) Then RunCommandScript(DirectCast(e.Data.GetData(DataFormats.FileDrop), String()).First)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles objects being dragged into the data box.
   Private Sub OutputBox_DragEnter(sender As Object, e As DragEventArgs) Handles OutputBox.DragEnter
      Try
         If e.Data.GetDataPresent(DataFormats.FileDrop) Then e.Effect = DragDropEffects.All
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure updates the interface.
   Private Sub UpdateStatus(sender As Object, e As EventArgs) Handles Updater.Tick
      Try
         CPUActiveLabel.Text = $"CPU {If(CPU.Clock.Status = TaskStatus.Running, "active", "inactive")}."

         If Not CPUEvent.ToString() = Nothing Then
            OutputBox.AppendText(CPUEvent.ToString())
            CPUEvent.Clear()
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Class