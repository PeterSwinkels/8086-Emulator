'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Emulator8086Program.CPU8086Class
Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Drawing
Imports System.Windows.Forms

'This class emulates the text 80x25 color video mode.
Public Class Text80x25ColorClass
   Implements VideoAdapterClass

   Private Const BLINK_BITMASK As Integer = &H80%   'Defines the character blink attribute bit.
   Private Const SCANLINE_COUNT As Integer = &HE%   'Defines the number of scanlines per character.

   Private ReadOnly COLORS() As Color = {Color.Black, Color.DarkBlue, Color.DarkGreen, Color.DarkCyan, Color.DarkRed, Color.Purple, Color.Brown, Color.DarkGray, Color.Gray, Color.Blue, Color.LimeGreen, Color.Cyan, Color.Red, Color.Magenta, Color.Yellow, Color.White}  'Defines the colors.
   Private ReadOnly CHARACTER_SIZE As Size = New Size(16, 24)                                                         'Defines the character size.
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
      Dim Count As Integer = VideoPageSizesE.Text80x25Color \ &H2%
      Dim Position As Integer = &H0%
      Dim VideoPageAddress As Integer = AddressesE.CGABuffer

      Do While Count > &H0%
         CPU.PutWord(VideoPageAddress + Position, &H700%)
         Count -= &H1%
         Position += &H2%
      Loop
   End Sub

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte, ByRef CodePage() As Integer) Implements VideoAdapterClass.Display
      Dim Attribute As New Byte
      Dim BitSet(,,) As Boolean = GetCharacterBits()
      Dim CharacterColor As Brush = Nothing
      Dim CursorScanlineEnd As Integer = If(Cursor.ScanLineEnd > &H3%, SCANLINE_COUNT, Cursor.ScanLineEnd)
      Dim CursorScanlineStart As Integer = If(Cursor.ScanLineStart > &H3%, SCANLINE_COUNT - &H1%, Cursor.ScanLineStart)
      Dim GraphicsO As Graphics = Nothing
      Dim Index As New Integer
      Dim Shift As New Integer
      Dim Target As New Point(0, 0)
      Dim VideoPageAddress As Integer = AddressesE.CGABuffer

      Try
         GraphicsO = Graphics.FromImage(Screen)

         With GraphicsO
            For Position As Integer = VideoPageAddress To VideoPageAddress + VideoPageSizesE.Text80x25Color Step &H2%
               If MCC.BlinkingOn Then
                  Attribute = ToByte((Memory(Position + &H1%) And &H7F%) \ &H10%)
               Else
                  Attribute = ToByte(Memory(Position + &H1%) \ &H10%)
               End If
               .FillRectangle(EGA.EGABrushes(Attribute), Target.X, Target.Y, CHARACTER_SIZE.Width, CHARACTER_SIZE.Height)

               If Target.X < TEXT_SCREEN_SIZE.Width - CHARACTER_SIZE.Width Then
                  Target.X += CHARACTER_SIZE.Width
               Else
                  Target.X = 0
                  If Target.Y < TEXT_SCREEN_SIZE.Height Then Target.Y += CHARACTER_SIZE.Height
               End If
            Next Position

            Target = New Point(0, 0)

            For Position As Integer = VideoPageAddress To VideoPageAddress + VideoPageSizesE.Text80x25Color Step &H2%
               Index = Memory(Position)
               Attribute = Memory(Position + &H1%)

               CharacterColor = EGA.EGABrushes(Attribute And &HF%)

               If ((Attribute And BLINK_BITMASK) = &H0%) OrElse BlinkCharactersVisible OrElse Not MCC.BlinkingOn Then
                  For y As Integer = &H0% To &H7%
                     Shift = &H7%
                     For Bit As Integer = &H0% To &H7%
                        If BitSet(Index, y, Bit) Then
                           .FillRectangle(CharacterColor, Target.X + (Shift * 2), Target.Y + (y * 3), 2, 3)
                        End If
                        Shift -= &H1%
                     Next Bit
                  Next y
               End If

               If Target.X < TEXT_SCREEN_SIZE.Width - CHARACTER_SIZE.Width Then
                  Target.X += CHARACTER_SIZE.Width
               Else
                  Target.X = 0
                  If Target.Y < TEXT_SCREEN_SIZE.Height Then Target.Y += CHARACTER_SIZE.Height
               End If
            Next Position

            If (Not Cursor.Off) AndAlso Cursor.Visible Then
               Attribute = ToByte(Memory((AddressesE.Text80x25ColorBuffer + (Cursor.Y * &HA0%) + (Cursor.X * &H2%)) + &H1%) And &HF%)
               .FillRectangle(EGA.EGABrushes(Attribute), Cursor.X * CHARACTER_SIZE.Width, (Cursor.Y * CHARACTER_SIZE.Height) + (CursorScanlineStart * PIXELS_PER_SCANLINE) - &H4%, CHARACTER_SIZE.Width, (CursorScanlineEnd * PIXELS_PER_SCANLINE) - (CursorScanlineStart * PIXELS_PER_SCANLINE))
            End If
         End With
      Catch
      Finally
         If GraphicsO IsNot Nothing Then GraphicsO.Dispose()
      End Try
   End Sub

   'This procedure is ignored.
   Public Sub DrawCharacter(Index As Integer, Attribute As Integer) Implements VideoAdapterClass.DrawCharacter
   End Sub

   'This procedure returns the bits from the current character bitmaps in video memory.
   Private Function GetCharacterBits() As Boolean(,,)
      Dim BitSet(&H0% To &HFF%, &H0% To &H7%, &H0% To &H7%) As Boolean
      Dim Character(&H0% To &H7%) As Byte
      Dim RemainingBits As New Integer
      Dim y As Integer = 0

      For Index As Integer = &H0% To &HFF%
         Array.Copy(CPU.Memory, If(Index < &H80%, AddressesE.Characters + (Index * &H8%), AddressesE.ExtendedCharacters + ((Index - &H80%) * &H8%)), Character, &H0%, Character.Length)

         y = 0
         For Each ScanLine As Byte In Character
            RemainingBits = ScanLine
            For Bit As Integer = &H0% To &H7%
               BitSet(Index, y, Bit) = CBool(RemainingBits And &H1%)
               RemainingBits >>= &H1%
            Next Bit
            y += 1
         Next ScanLine
      Next Index

      Return BitSet
   End Function

   'This procedure initializes the video adapter.
   Public Sub Initialize() Implements VideoAdapterClass.Initialize
      ClearBuffer()

      CPU.Memory(AddressesE.VideoPage) = &H0%
      ResetCursor()
      MCC.BlinkingOn = True

      EGA.EGABrushes = New List(Of SolidBrush)
      For Each [Color] As Color In COLORS
         EGA.EGABrushes.Add(New SolidBrush([Color]))
      Next [Color]
   End Sub

   'This procedure returns the screen size used by a video adapter.
   Public Function Resolution() As Size Implements VideoAdapterClass.Resolution
      Return New Size(TEXT_SCREEN_SIZE.Width, TEXT_SCREEN_SIZE.Height)
   End Function

   'This procedure scrolls the video adapter's buffer.
   Public Sub ScrollBuffer(Up As Boolean, ScrollArea As VideoAdapterClass.ScreenAreaStr, Count As Integer) Implements VideoAdapterClass.ScrollBuffer
      Dim Attribute As Integer = CPU.Registers(SubRegisters8BitE.BH)
      Dim BlankCell As New Integer
      Dim CharacterCell As New Integer
      Dim Position As New Integer
      Dim VideoPageAddress As Integer = AddressesE.CGABuffer

      If Count = &H0% OrElse Count > MCC.RowCount() Then
         For Row As Integer = ScrollArea.ULCRow To ScrollArea.LRCRow
            For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
               CPU.PutWord(VideoPageAddress + ((Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%)), Attribute << &H8%)
            Next Column
         Next Row
      Else
         For Scroll As Integer = &H1% To Count
            Select Case Up
               Case True
                  BlankCell = Attribute << &H8%

                  For Row As Integer = ScrollArea.ULCRow + &H1% To ScrollArea.LRCRow
                     For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
                        If Row < MCC.RowCount() Then
                           CharacterCell = CPU.GetWord(VideoPageAddress + (Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%))
                           CPU.PutWord(VideoPageAddress + ((Row - &H1%) * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%), CharacterCell)
                        Else
                           CPU.PutWord(VideoPageAddress + ((Row - &H1%) * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%), BlankCell)
                        End If
                     Next Column
                  Next Row

                  For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
                     CPU.PutWord(VideoPageAddress + (ScrollArea.LRCRow * TEXT_80_X_25_BYTES_PER_ROW) + (Column * 2), BlankCell)
                  Next Column
               Case False
                  BlankCell = Attribute << &H8%

                  For Row As Integer = ScrollArea.LRCRow - &H1% To ScrollArea.ULCRow - &H1% Step -&H1%
                     For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
                        If Row > &H0% Then
                           CharacterCell = CPU.GetWord(VideoPageAddress + (Row * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%))
                           CPU.PutWord(VideoPageAddress + ((Row + &H1%) * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%), CharacterCell)
                        Else
                           CPU.PutWord(VideoPageAddress + ((Row + &H1%) * TEXT_80_X_25_BYTES_PER_ROW) + (Column * &H2%), BlankCell)
                        End If
                     Next Column
                  Next Row

                  For Column As Integer = ScrollArea.ULCColumn To ScrollArea.LRCColumn
                     CPU.PutWord(VideoPageAddress + (ScrollArea.ULCRow * TEXT_80_X_25_BYTES_PER_ROW) + (Column * 2), BlankCell)
                  Next Column
            End Select
         Next Scroll
      End If
   End Sub
End Class
