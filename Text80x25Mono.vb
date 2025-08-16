'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Convert
Imports System.Drawing
Imports System.Linq
Imports System.Math
Imports System.Windows.Forms

'This class emulates a text 80x25 monochrome graphics adapter.
Public Class Text80x25MonoClass
   Implements VideoAdapterClass

   Private Const BLACK_ON_GREEN As Integer = &H70&         'Defines the black on green attribute bit.
   Private Const DARK_GREEN_ON_GREEN As Integer = &H78&    'Defines the black on green attribute bit.
   Private Const BLINK_BITMASK As Integer = &H80%          'Defines the character blink attribute bit.
   Private Const BRIGHT_BITMASK As Integer = &H8%          'Defines the bright character attribute bit.
   Private Const NON_BLINK_ATTRIBUTES As Integer = &H7F%   'Defines the attribute bits not related to blinking.
   Private Const UNDERLINE_BITMASK As Integer = &H7%       'Defines the character underline attribute bits.

   Private ReadOnly BLACK_ATTRIBUTES() As Integer = {&H0%, &H8%, &H80%, &H88%}                                        'Defines the black character attributes.
   Private ReadOnly CHARACTER_SIZE As Size = New Size(11, 16)                                                         'Defines the character size.
   Private ReadOnly FONT_NORMAL As New Font("Px437 IBM MDA", emSize:=14)                                              'Defines the normal font.
   Private ReadOnly FONT_UNDERLINE As New Font("Px437 IBM MDA", emSize:=14, FontStyle.Underline)                      'Defines the underlined font.
   Private ReadOnly HORIZONTAL_MARGIN As Integer = CInt(Ceiling(CHARACTER_SIZE.Width / 2))                            'Defines the horizontal screen size margin.
   Private ReadOnly TEXT_SCREEN_SIZE As Size = New Size(&H50% * CHARACTER_SIZE.Width, &H19% * CHARACTER_SIZE.Height)  'Defines the screen size measured in characters.

   Private BlinkCharactersVisible As Boolean = True  'Indicates whether or not the blinking characters are visible.

   Private WithEvents CharacterBlink As New Timer With {.Enabled = True, .Interval = 500}  'Contains the character blink timer.

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte, CodePage() As Integer) Implements VideoAdapterClass.Display
      Dim Attribute As New Byte
      Dim Character As New Char
      Dim CharacterColor As Brush = Nothing
      Dim CharacterFont As Font = Nothing
      Dim Target As New Point(0, 0)
      Dim VideoPageAddress As Integer = If(CPU.Memory(AddressesE.VideoPage) = 0, AddressesE.Text80x25MonoPage0, AddressesE.Text80x25MonoPage1)

      With Graphics.FromImage(Screen)
         .Clear(Color.Black)

         For Position As Integer = VideoPageAddress To VideoPageAddress + TEXT_80_X_25_MONO_BUFFER_SIZE Step &H2%
            Character = ToChar(CodePage(Memory(Position)))
            Attribute = Memory(Position + &H1%)

            If Attribute > &H0% AndAlso Not BLACK_ATTRIBUTES.Contains(Attribute) Then
               If (Attribute And NON_BLINK_ATTRIBUTES) = BLACK_ON_GREEN Then
                  .FillRectangle(New SolidBrush(Color.Lime), Target.X, Target.Y, CHARACTER_SIZE.Width, CHARACTER_SIZE.Height)
                  CharacterColor = New SolidBrush(Color.Green)
               ElseIf (Attribute And NON_BLINK_ATTRIBUTES) = DARK_GREEN_ON_GREEN Then
                  .FillRectangle(New SolidBrush(Color.Lime), Target.X, Target.Y, CHARACTER_SIZE.Width, CHARACTER_SIZE.Height)
                  CharacterColor = New SolidBrush(Color.Green)
               ElseIf (Attribute And BRIGHT_BITMASK) = &H0% Then
                  CharacterColor = New SolidBrush(Color.Green)
               Else
                  CharacterColor = New SolidBrush(Color.Lime)
               End If

               If (Attribute And UNDERLINE_BITMASK) = &H1% Then
                  CharacterFont = FONT_UNDERLINE
               Else
                  CharacterFont = FONT_NORMAL
               End If

               If ((Attribute And BLINK_BITMASK) = &H0%) OrElse BlinkCharactersVisible Then
                  .DrawString(Character, CharacterFont, CharacterColor, Target)
               End If
            End If

            If Target.X < TEXT_SCREEN_SIZE.Width - CHARACTER_SIZE.Width Then
               Target.X += CHARACTER_SIZE.Width
            Else
               Target.X = 0
               If Target.Y < TEXT_SCREEN_SIZE.Height Then Target.Y += CHARACTER_SIZE.Height
            End If
         Next Position

         If (Not Cursor.Off) AndAlso Cursor.Visible Then
            With Graphics.FromImage(Screen)
               .FillRectangle(New SolidBrush(Color.Lime), Cursor.X * CHARACTER_SIZE.Width, (Cursor.Y * CHARACTER_SIZE.Height) + Cursor.ScanLineStart, CHARACTER_SIZE.Width, Cursor.ScanLineEnd - Cursor.ScanLineStart)
            End With
         End If
      End With
   End Sub

   'This procedure regulates the blinking character's blinking.
   Private Sub CharacterBlink_Tick(sender As Object, e As EventArgs) Handles CharacterBlink.Tick
      BlinkCharactersVisible = Not BlinkCharactersVisible
   End Sub

   'This procedure clears video adapter's buffer.
   Public Sub ClearBuffer() Implements VideoAdapterClass.ClearBuffer
      Dim Count As Integer = TEXT_80_X_25_MONO_BUFFER_SIZE \ &H2%
      Dim Position As Integer = &H0%
      Dim VideoPageAddress As Integer = If(CPU.Memory(AddressesE.VideoPage) = 0, AddressesE.Text80x25MonoPage0, AddressesE.Text80x25MonoPage1)

      Do While Count > &H0%
         CPU.PutWord(VideoPageAddress + Position, &H200%)
         Count -= &H1%
         Position += &H2%
      Loop
   End Sub

   'This procedure initializes the video adapter.
   Public Sub Initialize() Implements VideoAdapterClass.Initialize
      ClearBuffer()

      CPU.Memory(AddressesE.VideoPage) = &H0%
      CPU.PutWord(AddressesE.CursorPositions, Word:=&H0%)
      CPU.PutWord(AddressesE.CursorScanLines, Word:=CURSOR_DEFAULT)
      CursorBlink.Enabled = True
   End Sub

   'This procedure returns the screen size used by a video adapter.
   Public Function Resolution() As Size Implements VideoAdapterClass.Resolution
      Return New Size(TEXT_SCREEN_SIZE.Width + HORIZONTAL_MARGIN, TEXT_SCREEN_SIZE.Height)
   End Function

   'This procedure scrolls the video adapter's buffer.
   Public Sub ScrollBuffer(Up As Boolean, ScrollArea As VideoAdapterClass.ScreenAreaStr, Count As Integer) Implements VideoAdapterClass.ScrollBuffer
      Dim Attribute As Integer = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.BH))
      Dim CharacterCell As New Integer
      Dim Position As New Integer
      Dim VideoPageAddress As Integer = If(CPU.Memory(AddressesE.VideoPage) = 0, AddressesE.Text80x25MonoPage0, AddressesE.Text80x25MonoPage1)

      If Count = &H0% OrElse Count > TEXT_80_X_25_LINE_COUNT Then
         For Row As Integer = ScrollArea.ULCRow To ScrollArea.LRCRow
            For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
               CPU.PutWord(VideoPageAddress + ((Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)), Attribute << &H8%)
            Next Column
         Next Row
      Else
         For Scroll As Integer = &H1% To Count
            Select Case Up
               Case True
                  For Row As Integer = ScrollArea.ULCRow + &H1% To ScrollArea.LRCRow
                     For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
                        If Row <= ScrollArea.LRCRow Then
                           CharacterCell = CPU.GET_WORD(VideoPageAddress + ((Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)))
                           CPU.PutWord(VideoPageAddress + ((Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)), Attribute)
                        Else
                           CharacterCell = Attribute << &H8%
                        End If
                        If Row >= ScrollArea.ULCRow Then
                           CPU.PutWord(VideoPageAddress + (((Row - &H1%) * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)), CharacterCell)
                        End If
                     Next Column
                  Next Row
               Case False
                  For Row As Integer = ScrollArea.LRCRow - &H1% To ScrollArea.ULCRow Step -&H1%
                     For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
                        If Row >= ScrollArea.ULCRow Then
                           CharacterCell = CPU.GET_WORD(VideoPageAddress + ((Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)))
                           CPU.PutWord(VideoPageAddress + ((Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)), Attribute)
                        Else
                           CharacterCell = (Attribute << &H8%)
                        End If
                        If Row <= ScrollArea.LRCRow Then
                           CPU.PutWord(VideoPageAddress + (((Row + &H1%) * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)), CharacterCell)
                        End If
                     Next Column
                  Next Row
            End Select
         Next Scroll
      End If
   End Sub
End Class
