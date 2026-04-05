'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Convert
Imports System.Drawing
Imports System.Windows.Forms

'This class emulates the text 80x25 color video mode.
Public Class Text80x25ColorClass
   Implements VideoAdapterClass

   Private Const BLINK_BITMASK As Integer = &H80%          'Defines the character blink attribute bit.
   Private Const SCANLINE_COUNT As Integer = &HE%          'Defines the number of scanlines per character.

   Private ReadOnly COLORS() As Color = {Color.Black, Color.DarkBlue, Color.DarkGreen, Color.DarkCyan, Color.DarkRed, Color.Purple, Color.Brown, Color.DarkGray, Color.Gray, Color.Blue, Color.Green, Color.Cyan, Color.Red, Color.Magenta, Color.Yellow, Color.White}  'Defines the colors.
   Private ReadOnly CHARACTER_SIZE As Size = New Size(14, 24)                                                         'Defines the character size.
   Private ReadOnly FONT As New Font("Px437 IBM VGA 8x14", emSize:=21)                                                'Defines the font.
   Private ReadOnly PIXELS_PER_SCANLINE As Integer = CInt(CHARACTER_SIZE.Height / SCANLINE_COUNT)                     'Defines the number of pixels per scanline.
   Private ReadOnly TEXT_SCREEN_SIZE As Size = New Size(&H50% * CHARACTER_SIZE.Width, &H19% * CHARACTER_SIZE.Height)  'Defines the screen size measured in characters.

   Private BlinkCharactersVisible As Boolean = True  'Indicates whether or not the blinking characters are visible.

   Private WithEvents CharacterBlink As New Timer With {.Enabled = True, .Interval = 500}  'Contains the character blink timer.

   'This procedure regulates the blinking character's blinking.
   Private Sub CharacterBlink_Tick(sender As Object, e As EventArgs) Handles CharacterBlink.Tick
      BlinkCharactersVisible = Not BlinkCharactersVisible
   End Sub

   'This procedure clears video adapter's buffer.
   Public Sub ClearBuffer() Implements VideoAdapterClass.ClearBuffer
      Dim Count As Integer = TEXT_80_X_25_COLOR_BUFFER_SIZE \ &H2%
      Dim Position As Integer = &H0%
      Dim VideoPageAddress As Integer = AddressesE.Text80x25ColorPage0

      Do While Count > &H0%
         CPU.PutWord(VideoPageAddress + Position, &H700%)
         Count -= &H1%
         Position += &H2%
      Loop
   End Sub

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte, CodePage() As Integer) Implements VideoAdapterClass.Display
      Dim Attribute As New Byte
      Dim BitmapO As Bitmap = DirectCast(Screen, Bitmap)
      Dim Character As New Char
      Dim CharacterColor As Brush = Nothing
      Dim ColorO As New Color
      Dim Target As New Point(0, 0)
      Dim VideoPageAddress As Integer = AddressesE.Text80x25ColorPage0

      With Graphics.FromImage(Screen)
         .Clear(Color.Black)

         For Position As Integer = VideoPageAddress To VideoPageAddress + TEXT_80_X_25_COLOR_BUFFER_SIZE Step &H2%
            Character = ToChar(CodePage(Memory(Position)))
            Attribute = ToByte((Memory(Position + &H1%) And &H7F%) \ &H10%)

            .FillRectangle(New SolidBrush(COLORS(Attribute)), Target.X, Target.Y, CHARACTER_SIZE.Width, CHARACTER_SIZE.Height)

            If Target.X < TEXT_SCREEN_SIZE.Width - CHARACTER_SIZE.Width Then
               Target.X += CHARACTER_SIZE.Width
            Else
               Target.X = 0
               If Target.Y < TEXT_SCREEN_SIZE.Height Then Target.Y += CHARACTER_SIZE.Height
            End If
         Next Position

         Target = New Point(0, 0)

         For Position As Integer = VideoPageAddress To VideoPageAddress + TEXT_80_X_25_MONO_BUFFER_SIZE Step &H2%
            Character = ToChar(CodePage(Memory(Position)))
            Attribute = Memory(Position + &H1%)

            CharacterColor = New SolidBrush(COLORS(Attribute And &HF%))

            If ((Attribute And BLINK_BITMASK) = &H0%) OrElse BlinkCharactersVisible Then
               .DrawString(Character, FONT, CharacterColor, Target.X - CInt(CHARACTER_SIZE.Width / 4), Target.Y)
            End If

            If Target.X < TEXT_SCREEN_SIZE.Width - CHARACTER_SIZE.Width Then
               Target.X += CHARACTER_SIZE.Width
            Else
               Target.X = 0
               If Target.Y < TEXT_SCREEN_SIZE.Height Then Target.Y += CHARACTER_SIZE.Height
            End If
         Next Position

         If (Not Cursor.Off) AndAlso Cursor.Visible Then
            Attribute = ToByte(Memory((AddressesE.Text80x25MonoPage0 + (Cursor.Y * &HA0%) + (Cursor.X * &H2%)) + &H1%) And &HF%)
            .FillRectangle(New SolidBrush(COLORS(Attribute)), Cursor.X * CHARACTER_SIZE.Width, (Cursor.Y * CHARACTER_SIZE.Height) + (Cursor.ScanLineStart * PIXELS_PER_SCANLINE), CHARACTER_SIZE.Width, (Cursor.ScanLineEnd * PIXELS_PER_SCANLINE) - (Cursor.ScanLineStart * PIXELS_PER_SCANLINE))
         End If
      End With
   End Sub

   'This procedure is ignored.
   Public Sub DrawCharacter(Index As Integer, Attribute As Integer) Implements VideoAdapterClass.DrawCharacter
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
      Return New Size(TEXT_SCREEN_SIZE.Width, TEXT_SCREEN_SIZE.Height)
   End Function

   'This procedure scrolls the video adapter's buffer.
   Public Sub ScrollBuffer(Up As Boolean, ScrollArea As VideoAdapterClass.ScreenAreaStr, Count As Integer) Implements VideoAdapterClass.ScrollBuffer
      Dim Attribute As Integer = CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.BH))
      Dim CharacterCell As New Integer
      Dim Position As New Integer
      Dim VideoPageAddress As Integer = AddressesE.Text80x25ColorPage0

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
