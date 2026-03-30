'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Windows.Forms

'This module contains information about the cursor.
Public Module CursorModule
   Public Structure CursorStr
      Public Off As Boolean             'Indicates whether or not the cursor is off
      Public ScanLineEnd As Integer     'Defines the cursor's end scan line.
      Public ScanLineStart As Integer   'Defines the cursor's start scan line.
      Public Visible As Boolean         'Indicates whether or not the cursor is visible.
      Public X As Integer               'Defines the cursor's horizontal position.
      Public Y As Integer               'Defines the cursor's vertical position.
   End Structure

   Public Const CURSOR_DEFAULT As Integer = &HD0E%     'Defines the default cursor's scan lines.
   Public Const CURSOR_DISABLED As Integer = &H2000%   'Defines the disabled cursor indicator.

   Public Cursor As New CursorStr With {.Off = False, .X = 0, .Y = 0, .Visible = False}   'Contains the cursor.

   Public WithEvents CursorBlink As New Timer With {.Enabled = False, .Interval = 500}  'Contains the cursor blink timer.

   'This procedure ensures the cursor will blink.
   Private Sub CursorBlink_Tick(sender As Object, e As EventArgs) Handles CursorBlink.Tick
      Try
         If Not Cursor.Off Then
            Cursor.ScanLineStart = CPU.Memory(AddressesE.CursorScanLines + &H1%)
            Cursor.ScanLineEnd = CPU.Memory(AddressesE.CursorScanLines)
            Cursor.Visible = Not Cursor.Visible
         End If

         Cursor.Off = (((Cursor.ScanLineStart << &H8%) Or Cursor.ScanLineEnd) = CURSOR_DISABLED) OrElse (Cursor.ScanLineStart > Cursor.ScanLineEnd)
         CursorPositionUpdate()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure updates the cursor's position. 
   Public Sub CursorPositionUpdate()
      Try
         Dim VideoPage As Integer = CPU.Memory(AddressesE.VideoPage)

         Cursor.X = CPU.Memory(AddressesE.CursorPositions + (VideoPage * &H2%))
         Cursor.Y = CPU.Memory(AddressesE.CursorPositions + ((VideoPage * &H2%) + &H1%))
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Module
