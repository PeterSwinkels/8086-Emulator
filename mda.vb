'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Generic
Imports System.Drawing

'This class emulates an MDA graphics adapter.
Public Class MDAClass
   'Defines the supported character attributes.
   Private Enum AttributesE As Byte
      Invisible          'Invisible.
      Normal             'Normal.
      Underline          'Underline.
      Bright             'Bright.
      BrightUnderline    'Bright underline.
      ReverseVideo       'Reverse video.
      InvisibleReverse   'Invisible reverse.
   End Enum

   Private Const MAP_ROW_LENGTH As Integer = &H20%     'Defines the number of characters per row in the character map.
   Private Const VIDEO_ADDRESS As Integer = &HB8000%   'Defines the video memory's base address.
   Private Const VIDEO_SIZE As Integer = &HFA0%        'Defines the video memory's size.

   Private ReadOnly CHARACTER_MAP_BRIGHT As Bitmap = My.Resources.cp437_bright     'Contains the bright character map.
   Private ReadOnly CHARACTER_MAP_NORMAL As Bitmap = My.Resources.cp437_normal     'Contains the normal character map.
   Private ReadOnly CHARACTER_MAP_REVERSE As Bitmap = My.Resources.cp437_reverse   'Contains the reverse character map.
   Private ReadOnly CHARACTER_SIZE As Size = New Size(9, 14)                       'Defines the characters's sizes.
   Private ReadOnly TEXT_SCREEN_SIZE As Size = New Size(80, 25)                    'Defines the screen size measured in characters.

   Public ReadOnly SCREEN_SIZE As Size = New Size(TEXT_SCREEN_SIZE.Width * CHARACTER_SIZE.Width, TEXT_SCREEN_SIZE.Height * CHARACTER_SIZE.Height)   'Contains the screen size measured in pixels.

   'This procedure draws the specified video buffer's context on the specified image.
   Public Sub Display(Screen As Image, Memory() As Byte)
      Dim Character As New Byte
      Dim Source As New Point
      Dim Target As New Point(1, 1)
      Dim TargetArea As New Rectangle

      With Graphics.FromImage(Screen)
         .Clear(Color.Black)

         For Position As Integer = VIDEO_ADDRESS To VIDEO_ADDRESS + VIDEO_SIZE Step &H2%
            Character = Memory(Position)
            Source.Y = Character \ MAP_ROW_LENGTH
            Source.X = (Character - (Source.Y * MAP_ROW_LENGTH)) * CHARACTER_SIZE.Width
            Source.Y *= CHARACTER_SIZE.Height

            TargetArea = New Rectangle(New Point((Target.X - 1) * CHARACTER_SIZE.Width, (Target.Y - 1) * CHARACTER_SIZE.Height), CHARACTER_SIZE)

            Select Case DirectCast(Memory(Position + &H1%), AttributesE)
               Case AttributesE.Bright
                  .DrawImage(CHARACTER_MAP_BRIGHT, TargetArea, New Rectangle(Source, CHARACTER_SIZE), GraphicsUnit.Pixel)
               Case AttributesE.BrightUnderline
                  .DrawImage(CHARACTER_MAP_BRIGHT, TargetArea, New Rectangle(Source, CHARACTER_SIZE), GraphicsUnit.Pixel)
                  .DrawLine(Pens.LightGreen, TargetArea.X, (TargetArea.Y - 1) + CHARACTER_SIZE.Height, TargetArea.X + CHARACTER_SIZE.Width, (TargetArea.Y - 1) + CHARACTER_SIZE.Height)
               Case AttributesE.Normal
                  .DrawImage(CHARACTER_MAP_NORMAL, TargetArea, New Rectangle(Source, CHARACTER_SIZE), GraphicsUnit.Pixel)
               Case AttributesE.InvisibleReverse
                  .FillRectangle(New SolidBrush(Color.LightGreen), TargetArea)
               Case AttributesE.ReverseVideo
                  .DrawImage(CHARACTER_MAP_REVERSE, TargetArea, New Rectangle(Source, CHARACTER_SIZE), GraphicsUnit.Pixel)
               Case AttributesE.Underline
                  .DrawImage(CHARACTER_MAP_NORMAL, TargetArea, New Rectangle(Source, CHARACTER_SIZE), GraphicsUnit.Pixel)
                  .DrawLine(Pens.Green, TargetArea.X, (TargetArea.Y - 1) + CHARACTER_SIZE.Height, TargetArea.X + CHARACTER_SIZE.Width, (TargetArea.Y - 1) + CHARACTER_SIZE.Height)
            End Select

            If Target.X < TEXT_SCREEN_SIZE.Width Then
               Target.X += 1
            Else
               Target.X = 1
               If Target.Y < TEXT_SCREEN_SIZE.Height Then Target.Y += 1
            End If
         Next Position
      End With
   End Sub
End Class
