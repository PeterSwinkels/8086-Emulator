'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Convert
Imports System.Linq
Imports System.Windows.Forms

'This module contains the keyboard related procedures.
Public Module KeyboardModule
   Private Const ASCII_0 As Integer = 48                                        'Defines the ASCII code for the zero ("0") character.
   Private Const ASCII_1 As Integer = 49                                        'Defines the ASCII code for the one ("1") character.
   Private Const ASCII_2 As Integer = 50                                        'Defines the ASCII code for the two ("2") character.
   Private Const ASCII_6 As Integer = 54                                        'Defines the ASCII code for the six ("6") character.
   Private Const ASCII_9 As Integer = 57                                        'Defines the ASCII code for the nine ("9") character.
   Private Const ASCII_A As Integer = 65                                        'Defines the ASCII code for the letter "A" character.
   Private Const ASCII_Z As Integer = 90                                        'Defines the ASCII code for the letter "Z" character.
   Private Const BIOS_KEY_CODE_1 As Integer = &H2%                              'Defines the BIOS key code for the one ("1") character.
   Private Const BIOS_KEY_CODE_1_ALT As Integer = &H78%                         'Defines the BIOS code for the one ("1") character combined with the Alt key.
   Private Const BIOS_KEY_CODE_F1 As Integer = &H3B%                            'Defines the BIOS code for the F1 key.
   Private Const BIOS_KEY_CODE_F1_ALT As Integer = &H68%                        'Defines the BIOS code for the F1 key combined with the Alt key.
   Private Const BIOS_KEY_CODE_F1_CONTROL As Integer = &H5E%                    'Defines the BIOS code for the F1 key combined with the Control key.
   Private Const BIOS_KEY_CODE_F1_SHIFT As Integer = &H54%                      'Defines the BIOS code for the F1 key combined with the Shift key.
   Private Const BIOS_KEY_CODE_F11 As Integer = &H85%                           'Defines the BIOS code for the F11 key.
   Private Const BIOS_KEY_CODE_F11_ALT As Integer = &H8B%                       'Defines the BIOS code for the F11 key combined with the Alt key.
   Private Const BIOS_KEY_CODE_F11_CONTROL As Integer = &H89%                   'Defines the BIOS code for the F11 key combined with the Control key.
   Private Const BIOS_KEY_CODE_F11_SHIFT As Integer = &H87%                     'Defines the BIOS code for the F11 key combined with the Shift key.
   Private Const CHARACTER_CODE_0_SHIFT As Integer = &H29%                      'Defines the character code for the zero ("0") character  combined with the Shift key.
   Private Const CHARACTER_CODE_6_CONTROL As Integer = &H1E%                    'Defines the character code for the six ("6") character combined with the Control key.
   Private Const PUNCTUATION_CHARACTERS As String = "-=[];'`\,./_+{}:""~|<>?"   'Defines the punctuation characters.
   Private Const PUNCTUATION_CHARACTERS_SHIFT_INDEX As Integer = 12             'Defines the index of the first shifted punctuation character.

   Private ReadOnly CHARACTER_CODES_1_9_SHIFT() As Integer = {&H21%, &H40%, &H23%, &H24%, &H25%, &H5E%, &H26%, &H2A%, &H28%}                                                                                                                                    'Contains the codes for the characters one to nine ("1"-"9") combined with the Shift key.
   Private ReadOnly CHARACTER_CODES_PUNCTUATION_CONTROL() As Integer = {&H1F%, Nothing, &H1B%, &H1D%, Nothing, Nothing, Nothing, &H1C%, Nothing, Nothing, Nothing}                                                                                              'Contains the codes for the punctuation characters combined with the Control key.
   Private ReadOnly BIOS_KEY_CODES_A_Z() As Integer = {&H1E%, &H30%, &H2E%, &H20%, &H12%, &H21%, &H22%, &H23%, &H17%, &H24%, &H25%, &H26%, &H32%, &H31%, &H18%, &H19%, &H10%, &H13%, &H1F%, &H14%, &H16%, &H2F%, &H11%, &H2D%, &H15%, &H2C%}                    'Contains the BIOS key codes for the letters ("A"-"Z").
   Private ReadOnly BIOS_KEY_CODES_COMMAND() As Integer = {&HE08%, &H5300%, &H5000%, &H4F00%, &H1C0D%, &H11B%, &H4700%, &H5200%, Nothing, &H372A%, &H4A2D%, &H4E2B%, &H352F%, &H4B00%, &H5100%, &H4900%, Nothing, &H4D00%, &H3920%, &HF09%, &H4800%}            'Contains the BIOS key codes for the command keys.
   Private ReadOnly BIOS_KEY_CODES_COMMAND_ALT() As Integer = {&HE%, &HA3%, &HA0%, &H9F%, &HA6%, &H1%, &H97%, &HA2%, Nothing, &H37%, &H4A%, &H4E%, &HA4%, &H9B%, &HA1%, &H99%, Nothing, &H9D%, &H39%, &HA5%, &H98%}                                             'Contains the BIOS key codes for the command keys combined with the Alt key.
   Private ReadOnly BIOS_KEY_CODES_COMMAND_CONTROL() As Integer = {&HE7F%, &H9300%, &H9100%, &H7500%, &H1C0A%, &H11B%, &H7700%, &H9200%, &H8F00%, &H9600%, &H8E00%, Nothing, &H9500%, &H7300%, &H7600%, &H8400%, &H7200%, &H7400%, &H3920%, &H9400%, &H8D00%}   'Contains the BIOS key codes for the command keys combined with the Control key.
   Private ReadOnly BIOS_KEY_CODES_COMMAND_SHIFT() As Integer = {&HE08%, &H532E%, &H5032%, &H4F31%, &H1C0D%, &H11B%, &H4737%, &H5230%, &H4C35%, Nothing, &H4A2D%, &H4E2B%, &H352F%, &H4B34%, &H5133%, &H4939%, Nothing, &H4D36%, &H3920%, &HF00%, &H4838%}      'Contains the BIOS key codes for the command keys combined with the Shift key.
   Private ReadOnly BIOS_KEY_CODES_PUNCTUATION() As Integer = {&HC%, &HD%, &H1A%, &H1B%, &H27%, &H28%, &H29%, &H2B%, &H33%, &H34%, &H35%}                                                                                                                       'Contains the BIOS key codes for the punctuation characters.
   Private ReadOnly BIOS_KEY_CODES_PUNCTUATION_ALT() As Integer = {&H82%, &H83%, &H1A%, &H1B%, &H27%, Nothing, Nothing, &H26%, Nothing, Nothing, Nothing}                                                                                                       'Contains the BIOS key codes for the punctuation characters combined with the Alt key.
   Private ReadOnly BIOS_KEY_CODES_PUNCTUATION_SHIFT() As Integer = {&HE08%, &H532E%, &H5032%, &H4F31%, &H1C0D%, &H11B%, &H4737%, &H5230%, &H4C35%, Nothing, &H4A2D%, &H4E2B%, &H352F%, &H4B34%, &H5133%, &H4939%, Nothing, &H4D36%, &H3920%, &HF00%, &H4838%}  'Contains the BIOS key codes for the punctuation characters combined with the Shift key.
   Private ReadOnly COMMAND_KEYS() As Integer = {Keys.Back, Keys.Delete, Keys.Down, Keys.End, Keys.Enter, Keys.Escape, Keys.Home, Keys.Insert, Keys.NumPad5, Keys.Multiply, Keys.Subtract, Keys.Add, Keys.Divide, Keys.Left, Keys.PageDown, Keys.PageUp, Keys.PrintScreen, Keys.Right, Keys.Space, Keys.Tab, Keys.Up}  'Contains the command keys.

   'This procedure converts the specified ASCII code combined with the specified modifier keys to a BIOS scan and character code and returns the result.
   Private Function GetBIOSCKeyCode(KeyASCII As Integer, Alt As Boolean, Control As Boolean, Shift As Boolean) As Integer?
      Try
         Dim BIOSKeyCode As New Integer?
         Dim CharacterCode As Integer = &H0%
         Dim PunctuationIndex As New Integer

         Select Case KeyASCII
            Case ASCII_0 To ASCII_9
               If Alt Then
                  BIOSKeyCode = BIOS_KEY_CODE_1_ALT + If(KeyASCII = ASCII_0, 9, (KeyASCII - ASCII_1))
               ElseIf Control Then
                  Select Case KeyASCII
                     Case ASCII_2
                        BIOSKeyCode = BIOS_KEY_CODE_1 + If(KeyASCII = ASCII_0, 9, (KeyASCII - ASCII_1))
                     Case ASCII_6
                        CharacterCode = CHARACTER_CODE_6_CONTROL
                        BIOSKeyCode = BIOS_KEY_CODE_1 + If(KeyASCII = ASCII_0, 9, (KeyASCII - ASCII_1))
                  End Select
               ElseIf Shift Then
                  CharacterCode = If(KeyASCII = ASCII_0, CHARACTER_CODE_0_SHIFT, CHARACTER_CODES_1_9_SHIFT(KeyASCII - ASCII_1))
                  BIOSKeyCode = BIOS_KEY_CODE_1 + If(KeyASCII = ASCII_0, 9, (KeyASCII - ASCII_1))
               Else
                  CharacterCode = KeyASCII
                  BIOSKeyCode = BIOS_KEY_CODE_1 + If(KeyASCII = ASCII_0, 9, (KeyASCII - ASCII_1))
               End If
            Case ASCII_A To ASCII_Z
               If Control Then
                  CharacterCode = KeyASCII - &H40%
               ElseIf Shift Then
                  CharacterCode = KeyASCII
               ElseIf Not Alt Then
                  CharacterCode = KeyASCII + &H20%
               End If
               BIOSKeyCode = BIOS_KEY_CODES_A_Z(KeyASCII - ASCII_A)
            Case Else
               PunctuationIndex = PUNCTUATION_CHARACTERS.IndexOf(ToChar(KeyASCII))
               Shift = (PunctuationIndex >= PUNCTUATION_CHARACTERS_SHIFT_INDEX)
               If PunctuationIndex >= 0 Then
                  If Alt Then
                     If Not BIOS_KEY_CODES_PUNCTUATION_ALT(PunctuationIndex) = Nothing Then BIOSKeyCode = BIOS_KEY_CODES_PUNCTUATION_ALT(PunctuationIndex)
                  ElseIf Control Then
                     CharacterCode = CHARACTER_CODES_PUNCTUATION_CONTROL(PunctuationIndex)
                     If Not CharacterCode = Nothing Then BIOSKeyCode = BIOS_KEY_CODES_PUNCTUATION(PunctuationIndex)
                  ElseIf Shift Then
                     CharacterCode = KeyASCII
                     BIOSKeyCode = BIOS_KEY_CODES_PUNCTUATION(PunctuationIndex - (PUNCTUATION_CHARACTERS_SHIFT_INDEX - 1))
                  Else
                     CharacterCode = KeyASCII
                     BIOSKeyCode = BIOS_KEY_CODES_PUNCTUATION(PunctuationIndex)
                  End If
               End If
         End Select

         If BIOSKeyCode IsNot Nothing Then BIOSKeyCode = ((BIOSKeyCode << &H8%) Or CharacterCode)

         Return BIOSKeyCode
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure converts the specified command key combined with the specified modifier keys to a BIOS scan and character code and returns the result.
   Private Function GetBIOSCMDKeyCode(Key As Keys, Alt As Boolean, Control As Boolean, Shift As Boolean) As Integer?
      Try
         Dim BIOSKeyCode As New Integer?
         Dim CommandKeyIndex As Integer = Array.IndexOf(COMMAND_KEYS, CInt(Key))

         If Alt Then
            BIOSKeyCode = BIOS_KEY_CODES_COMMAND_ALT(CommandKeyIndex) << &H8%
         ElseIf Control Then
            BIOSKeyCode = BIOS_KEY_CODES_COMMAND_CONTROL(CommandKeyIndex)
         ElseIf Shift Then
            BIOSKeyCode = BIOS_KEY_CODES_COMMAND_SHIFT(CommandKeyIndex)
         Else
            BIOSKeyCode = BIOS_KEY_CODES_COMMAND(CommandKeyIndex)
         End If

         Return BIOSKeyCode
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure converts the specified function key combined with the specified modifier keys to a BIOS scan and character code and returns the result.
   Private Function GetBIOSFKeyCode(Key As Keys, Alt As Boolean, Control As Boolean, Shift As Boolean) As Integer?
      Try
         Dim BIOSKeyCode As New Integer?

         Select Case Key
            Case Keys.F1 To Keys.F10
               If Alt Then
                  BIOSKeyCode = BIOS_KEY_CODE_F1_ALT + (Key - Keys.F1)
               ElseIf Control Then
                  BIOSKeyCode = BIOS_KEY_CODE_F1_CONTROL + (Key - Keys.F1)
               ElseIf Shift Then
                  BIOSKeyCode = BIOS_KEY_CODE_F1_SHIFT + (Key - Keys.F1)
               Else
                  BIOSKeyCode = BIOS_KEY_CODE_F1 + (Key - Keys.F1)
               End If
            Case Keys.F11 To Keys.F12
               If Alt Then
                  BIOSKeyCode = BIOS_KEY_CODE_F11_ALT + (Key - Keys.F11)
               ElseIf Control Then
                  BIOSKeyCode = BIOS_KEY_CODE_F11_CONTROL + (Key - Keys.F11)
               ElseIf Shift Then
                  BIOSKeyCode = BIOS_KEY_CODE_F11_SHIFT + (Key - Keys.F11)
               Else
                  BIOSKeyCode = BIOS_KEY_CODE_F11 + (Key - Keys.F11)
               End If
         End Select

         If BIOSKeyCode IsNot Nothing Then BIOSKeyCode = (BIOSKeyCode << &H8%)

         Return BIOSKeyCode
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure converts the specified key combined with the specified modifier keys to a BIOS scan and character code and returns the result.
   Public Function GetBIOSKeyCode(Key As KeyEventArgs, ByRef KeyASCII As Integer?) As Integer?
      Try
         Dim BIOSKeyCode As New Integer?

         If KeyASCII Is Nothing Then
            Select Case Key.KeyCode
               Case Keys.F1 To Keys.F12
                  BIOSKeyCode = GetBIOSFKeyCode(Key.KeyCode, Key.Alt, Key.Control, Key.Shift)
               Case Else
                  If COMMAND_KEYS.Contains(Key.KeyCode) Then
                     BIOSKeyCode = GetBIOSCMDKeyCode(Key.KeyCode, Key.Alt, Key.Control, Key.Shift)
                  Else
                     KeyASCII = PunctuationKeyToASCII(Key.KeyCode)
                     If KeyASCII Is Nothing Then KeyASCII = Key.KeyValue
                     BIOSKeyCode = GetBIOSCKeyCode(CInt(KeyASCII), Key.Alt, Key.Control, Key.Shift)
                  End If
            End Select
         Else
            BIOSKeyCode = GetBIOSCKeyCode(CInt(KeyASCII), Key.Alt, Key.Control, Key.Shift)
         End If

         Return BIOSKeyCode
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure keeps track of the most recently pressed key's BIOS key code.
   Public Function LastBIOSKeyCode(Optional NewLastBIOSKeyCode As Integer = Nothing, Optional Clear As Boolean = False) As Integer
      Try
         Static CurrentLastBIOSKeyCode As Integer = Nothing

         If Clear Then
            CurrentLastBIOSKeyCode = Nothing
         ElseIf Not NewLastBIOSKeyCode = Nothing Then
            CurrentLastBIOSKeyCode = CInt(NewLastBIOSKeyCode)
         End If

         Return CurrentLastBIOSKeyCode
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns the ASCII value of the specified punctuation key's character if it recognized.
   Public Function PunctuationKeyToASCII(Key As Keys) As Integer?
      Try
         Dim KeyASCII As New Integer?

         Select Case Key
            Case Keys.Oem7
               KeyASCII = ToInt32("'"c)
            Case Keys.OemBackslash, Keys.Oem5
               KeyASCII = ToInt32("\"c)
            Case Keys.OemCloseBrackets
               KeyASCII = ToInt32("]"c)
            Case Keys.Oemcomma
               KeyASCII = ToInt32(","c)
            Case Keys.OemMinus
               KeyASCII = ToInt32("-"c)
            Case Keys.Oemplus
               KeyASCII = ToInt32("="c)
            Case Keys.OemOpenBrackets
               KeyASCII = ToInt32("["c)
            Case Keys.OemPeriod
               KeyASCII = ToInt32("."c)
            Case Keys.OemQuestion
               KeyASCII = ToInt32("/"c)
            Case Keys.OemSemicolon
               KeyASCII = ToInt32(";"c)
            Case Keys.Oemtilde
               KeyASCII = ToInt32("`"c)
         End Select

         Return KeyASCII
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function
End Module
