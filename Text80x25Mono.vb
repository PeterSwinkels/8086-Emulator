﻿'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Convert
Imports System.Drawing
Imports System.Linq
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
   Private Const VIDEO_SIZE As Integer = &HFA0%            'Defines the video memory's size.

   Private ReadOnly BLACK_ATTRIBUTES() As Integer = {&H0%, &H8%, &H80%, &H88%}                                  'Defines the black character attributes.
   Private ReadOnly CHARACTER_SIZE As Size = New Size(9, 14)                                                    'Defines the character size.
   Private ReadOnly FONT_NORMAL As New Font("Px437 IBM MDA", emSize:=9)                                         'Defines the normal font.
   Private ReadOnly FONT_UNDERLINE As New Font("Px437 IBM MDA", emSize:=9, FontStyle.Underline)                 'Defines the underlined font.
   Private ReadOnly TEXT_SCREEN_SIZE As Size = New Size(80 * CHARACTER_SIZE.Width, 25 * CHARACTER_SIZE.Height)  'Defines the screen size measured in characters.

   Private BlinkCharactersVisible As Boolean = True  'Indicates whether or not the blinking characters are visible.

   Private WithEvents CharacterBlink As New Timer With {.Enabled = True, .Interval = 500}  'Contains the character blink timer.

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte, CodePage() As Integer) Implements VideoAdapterClass.Display
      Dim Attribute As New Byte
      Dim Character As New Char
      Dim CharacterColor As Brush = Nothing
      Dim CharacterFont As Font = Nothing
      Dim Target As New Point(0, 0)

      With Graphics.FromImage(Screen)
         .Clear(Color.Black)

         For Position As Integer = AddressesE.Text80x25Mono To AddressesE.Text80x25Mono + VIDEO_SIZE Step &H2%
            Character = ToChar(CodePage(Memory(Position)))
            Attribute = Memory(Position + &H1%)

            If Attribute > &H0% AndAlso Not BLACK_ATTRIBUTES.Contains(Attribute) Then
               If (Attribute And NON_BLINK_ATTRIBUTES) = BLACK_ON_GREEN Then
                  .FillRectangle(New SolidBrush(Color.LightGreen), Target.X, Target.Y, CHARACTER_SIZE.Width, CHARACTER_SIZE.Height)
                  CharacterColor = New SolidBrush(Color.Green)
               ElseIf (Attribute And NON_BLINK_ATTRIBUTES) = DARK_GREEN_ON_GREEN Then
                  .FillRectangle(New SolidBrush(Color.LightGreen), Target.X, Target.Y, CHARACTER_SIZE.Width, CHARACTER_SIZE.Height)
                  CharacterColor = New SolidBrush(Color.Green)
               ElseIf (Attribute And BRIGHT_BITMASK) = &H0% Then
                  CharacterColor = New SolidBrush(Color.Green)
               Else
                  CharacterColor = New SolidBrush(Color.LightGreen)
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
               .FillRectangle(New SolidBrush(Color.LightGreen), Cursor.X * CHARACTER_SIZE.Width, (Cursor.Y * CHARACTER_SIZE.Height) + Cursor.ScanLineStart, CHARACTER_SIZE.Width, Cursor.ScanLineEnd - Cursor.ScanLineStart)
            End With
         End If
      End With
   End Sub

   'This procedure regulates the blinking characters's blinking.
   Private Sub CharacterBlink_Tick(sender As Object, e As EventArgs) Handles CharacterBlink.Tick
      Try
         BlinkCharactersVisible = Not BlinkCharactersVisible
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure returns the screen size used by a video adapter.
   Public Function ScreenSize() As Size Implements VideoAdapterClass.ScreenSize
      Return New Size(TEXT_SCREEN_SIZE.Width, TEXT_SCREEN_SIZE.Height)
   End Function
End Class
