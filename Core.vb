'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Environment
Imports System.IO
Imports System.Globalization
Imports System.Linq
Imports System.Math
Imports System.Text
Imports System.Windows.Forms

'This module contains this program's core procedures.
Public Module CoreModule
   'This structure defines the a command's remainder yet to be parsed and an element that has been extracted.
   Private Structure ParsedStr
      Public Element As String     'Defines an element in a command that has been extracted during parsing.
      Public Remainder As String   'Defines the command's remainder yet to be parsed.
   End Structure

   Private Const ASSIGNMENT_OPERATOR As Char = "="c              'Defines the assignment operator for operands.
   Private Const CHARACTER_OPERAND_DELIMITER As Char = "'"c      'Defines a character operand's delimiter.
   Private Const DEFAULT_DISASSEMBLY_COUNT As Integer = &H10%    'Defines the default number of bytes disassembled.
   Private Const DEFAULT_MEMORY_DUMP_COUNT As Integer = &H100%   'Defines the default memory size.
   Private Const ESCAPE_CHARACTER As Byte = &H2F%                'Defines the escape character used in output.
   Private Const MEMORY_OPERAND_END As Char = "]"c               'Defines a memory operand's first character.
   Private Const MEMORY_OPERAND_START As Char = "["c             'Defines a memory operand's last character.
   Private Const SCRIPT_HEADER As String = "[SCRIPT]"            'Defines a mandatory header for all script files.
   Private Const STRING_OPERAND_DELIMITER As Char = """"c        'Defines a string operand's delimiter.
   Private Const VALUES_OPERAND_END As Char = "}"c               'Defines a values operand's end.
   Private Const VALUES_OPERAND_START As Char = "{"c             'Defines a values operand's start.

   Public WithEvents CPU As New CPU8086Class                                        'Contains a reference to the CPU 8086 class.
   Public WithEvents Tracer As New Timer With {.Enabled = False, .Interval = 100}   'Contains the clock that handles CPU execution tracing.
   Private WithEvents Assembler As New AssemblerClass                               'Contains a reference to the assembler.
   Private WithEvents Disassembler As New DisassemblerClass                         'Contains a reference to the disassembler.

   Public AssemblyModeOn As Boolean = False   'Indicates whether input is interpreted as assembly language.
   Public Output As TextBox = Nothing         'Contains a reference to an output.

   Private ReadOnly GET_OPERAND As Func(Of String, String) = Function(Input As String) (Input.Substring(Input.IndexOf(ASSIGNMENT_OPERATOR) + 1))                                                                                                             'Returns the specified input's operand.
   Private ReadOnly IS_CHARACTER_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (Operand.Trim().StartsWith(CHARACTER_OPERAND_DELIMITER) AndAlso Operand.Trim().EndsWith(CHARACTER_OPERAND_DELIMITER) AndAlso Operand.Trim().Length = 3)   'Indicates whether the specified operand is a character.
   Private ReadOnly IS_MEMORY_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (Operand.Trim().StartsWith(MEMORY_OPERAND_START) AndAlso Operand.Trim().EndsWith(MEMORY_OPERAND_END))                                                        'Indicates whether the specified operand is a memory location.
   Private ReadOnly IS_STRING_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (Operand.Trim().StartsWith(STRING_OPERAND_DELIMITER) AndAlso Operand.Trim().EndsWith(STRING_OPERAND_DELIMITER))                                              'Indicates whether the specified operand is a string.
   Private ReadOnly IS_VALUES_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (Operand.Trim().StartsWith(VALUES_OPERAND_START) AndAlso Operand.Trim().EndsWith(VALUES_OPERAND_END))                                                        'Indicates whether the specified operand is an array of values.
   Private ReadOnly WITHOUT_DELIMITERS As Func(Of String, String) = Function(Operand As String) (Operand.Substring(1, Operand.Length - 2))                                                                                                                   'Returns the specified operand with its first and last characters removed.

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         InterfaceWindow.Show()

         Do While InterfaceWindow.Visible
            Application.DoEvents()
         Loop
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure translates the specified memory operands to a flat memory address.
   Private Function AddressFromOperand(Operand As String) As Integer?
      Try
         Dim Offset As New Integer?
         Dim Parsed As New ParsedStr
         Dim Segment As New Integer?

         If IS_MEMORY_OPERAND(Operand) Then Operand = WITHOUT_DELIMITERS(Operand)

         Parsed.Remainder = Operand

         Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=Nothing, Ending:=":"c)
         Segment = GetLiteral(Parsed.Element)

         Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=Nothing, Ending:=Nothing)
         Offset = GetLiteral(Parsed.Element)

         If Segment IsNot Nothing AndAlso Offset IsNot Nothing Then Return (CInt(Segment) << &H4%) + CInt(Offset)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure translates the specified assembly code to opcodes.
   Public Sub Assemble(Optional Input As String = Nothing, Optional StartAddress As Integer? = Nothing)
      Try
         Dim Opcodes As New List(Of Byte)
         Static Address As New Integer

         If AssemblyModeOn Then
            If Input.Trim() = Nothing Then
               AssemblyModeOn = False
               Output.AppendText($"Done.{NewLine}")
            Else
               Opcodes = Assembler.Assemble(Input)
               If Opcodes Is Nothing Then
                  Output.AppendText($"{Input}{NewLine}")
               ElseIf Opcodes.Count = 0 Then
                  Output.AppendText($"{Input}{NewLine}")
                  Output.AppendText($"Undefined error.{NewLine}")
               Else
                  Output.AppendText($"{Address:X8}   {Input,-25} -> {String.Join(Nothing, (From Opcode In Opcodes Select Opcode.ToString("X2")).ToArray())}{NewLine}")
                  Array.Copy(Opcodes.ToArray(), 0, CPU.Memory, Address, Opcodes.Count)
                  Address += Opcodes.Count
               End If
            End If
         ElseIf StartAddress IsNot Nothing Then
            Address = CInt(StartAddress)
            AssemblyModeOn = True
            Output.AppendText($"Assembler started.{NewLine}")
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles any exceptions raised by the assembler.
   Private Sub Assembler_HandleError(ExceptionO As Exception) Handles Assembler.HandleError
      Dim Message As String = ExceptionO.Message

      Try
         Output.AppendText($"Assembler error: {Message}{NewLine}")
      Catch Exception2 As Exception
         DisplayException(Exception2.Message)
      End Try
   End Sub

   'This procedure handles the emulated CPU's halt events.
   Private Sub CPU_Halt() Handles CPU.Halt
      Try
         CPU.Clock.Stop()
         Tracer.Stop()
         Output.AppendText($"Halted.{NewLine}")
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles the emulated CPU's interrupt events.
   Private Sub CPU_Interrupt(Number As Integer, AH As Integer) Handles CPU.Interrupt
      Try
         CPU.Clock.Stop()
         Tracer.Stop()
         Output.AppendText($"INT {Number:x}, {AH:X}{NewLine}")
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles the emulated CPU's I/O read events.
   Private Sub CPU_ReadIOPort(Port As Integer, ByRef Value As Integer, Is8Bit As Boolean) Handles CPU.ReadIOPort
      Try
         CPU.Clock.Stop()
         Tracer.Stop()
         Output.AppendText($"IN {If(Is8Bit, "AL", "AX")}, {Port:X}{NewLine}")
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles the emulated CPU's I/O write events.
   Private Sub CPU_WriteIOPort(Port As Integer, Value As Integer, Is8Bit As Boolean) Handles CPU.WriteIOPort
      Try
         CPU.Clock.Stop()
         Tracer.Stop()
         Output.AppendText($"OUT {Port:X}, {If(Is8Bit, $"{Value:X2}", $"{Value:X4}")}{NewLine}")
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure disassembles the specified code and returns the resulting lines of code.
   Private Function Disasemble(Code As List(Of Byte), Position As Integer, Optional Count As Integer? = Nothing) As String
      Try
         Dim Disassembly As New StringBuilder
         Dim EndPosition As New Integer
         Dim HexadecimalCode As String = Nothing
         Dim Instruction As String = Nothing
         Dim PreviousPosition As New Integer

         If Count Is Nothing Then Count = DEFAULT_DISASSEMBLY_COUNT

         EndPosition = If(CInt(Count) = &H0%, Code.Count, Position + CInt(Count))

         With Disassembler
            Do Until Position >= EndPosition
               PreviousPosition = Position + &H1%
               Instruction = .Disassemble(Code, Position)
               HexadecimalCode = .BytesToHexadecimal(.GetBytes(Code, PreviousPosition - &H1%, (Position - PreviousPosition) + &H1%), NoPrefix:=True, Reverse:=False)
               Disassembly.Append($"{(PreviousPosition - &H1%):X8} {HexadecimalCode,-25}{Instruction}{NewLine}")
            Loop
         End With

         Return Disassembly.ToString()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure displays the specified exception message.
   Public Sub DisplayException(Message As String)
      Try
         If Output Is Nothing Then
            MessageBox.Show(Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
         Else
            Output.AppendText($"{Message}{NewLine}")
         End If
      Catch
         Application.Exit()
      End Try
   End Sub

   'This procedure escapes the specified byte or convert it to a character and returns the result.
   Private Function EscapeByte([Byte] As Byte) As String
      Try
         Return If([Byte] >= ToByte(" "c) AndAlso [Byte] <= ToByte("~"c), If([Byte] = ESCAPE_CHARACTER, New String(ToChar(ESCAPE_CHARACTER), count:=2), ToChar([Byte])), $"{ToChar(ESCAPE_CHARACTER)}{[Byte]:X2}")
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns the flat memory address for the emulated CPU's CS:IP registers.
   Public Function GetFlatCSIP() As Integer
      Try
         Return (CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.CS)) << &H4%) Or CInt(CPU.Registers(CPU8086Class.Registers16BitE.IP))
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns the literal value represented by a command element.
   Private Function GetLiteral(Element As String, Optional Is8Bit As Boolean = True) As Integer?
      Try
         Dim Address As New Integer?
         Dim Buffer As New Integer
         Dim Literal As New Integer?
         Dim Parsed As New ParsedStr
         Dim Register As Object = Nothing

         If Element IsNot Nothing Then
            If IS_CHARACTER_OPERAND(Element) Then
               Literal = ToByte(Element.Chars(1))
            ElseIf IS_MEMORY_OPERAND(Element) Then
               Address = AddressFromOperand(Element.Trim())
               If Address IsNot Nothing Then Literal = If(Is8Bit, CPU.Memory(CInt(Address)), CPU.GetWord(CInt(Address)))
            ElseIf Integer.TryParse(Element, NumberStyles.HexNumber, Nothing, Buffer) Then
               Literal = Buffer
            Else
               Register = GetRegisterByName(Element.Trim().ToUpper())
               If Register IsNot Nothing Then Literal = CInt(CPU.Registers(Register))
            End If
         End If

            Return Literal
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns a memory dump.
   Private Function GetMemoryDump(AllHexadecimal As Boolean, Offset As Integer, Optional Count As Integer? = Nothing) As String
      Try
         Dim Dump As New StringBuilder

         If Count Is Nothing Then Count = DEFAULT_MEMORY_DUMP_COUNT
         If Offset + Count >= CPU.Memory.Length Then Count = CPU.Memory.Length - Offset

         With Dump
            If AllHexadecimal Then
               CPU.Memory.ToList().GetRange(Offset, CInt(Count)).ForEach(Sub([Byte] As Byte) .Append($"{[Byte]:X2} "))
            Else
               CPU.Memory.ToList().GetRange(Offset, CInt(Count)).ForEach(Sub([Byte] As Byte) .Append(EscapeByte([Byte])))
            End If

            .Append($"{NewLine}{NewLine}")
         End With

         Return Dump.ToString()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified memory location's value.
   Private Function GetMemoryValue(Address As Integer) As String
      Try
         Return $"Byte = 0x{CPU.Memory(Address):X2}   Word = 0x{CPU.GetWord(Address):X4}   Characters = '{EscapeByte(CPU.Memory(Address))}{EscapeByte(CPU.Memory(Address + &H1%))}'{NewLine}"
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns the emulated CPU register with the specified name.
   Private Function GetRegisterByName(Name As String, Optional ByRef Is8Bit As Boolean = Nothing) As Object
      Try
         Is8Bit = True

         For Each Register As CPU8086Class.FlagRegistersE In [Enum].GetValues(GetType(CPU8086Class.FlagRegistersE))
            If Register.ToString() = Name Then Return Register
         Next Register

         For Each Register As CPU8086Class.SubRegisters8BitE In [Enum].GetValues(GetType(CPU8086Class.SubRegisters8BitE))
            If Register.ToString() = Name Then Return Register
         Next Register

         Is8Bit = False

         For Each Register As CPU8086Class.Registers16BitE In [Enum].GetValues(GetType(CPU8086Class.Registers16BitE))
            If Register.ToString() = Name Then Return Register
         Next Register

         For Each Register As CPU8086Class.SegmentRegistersE In [Enum].GetValues(GetType(CPU8086Class.SegmentRegistersE))
            If Register.ToString() = Name Then Return Register
         Next Register
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns the values of all the emulated CPU registers.
   Private Function GetRegisterValues() As String
      Try
         Dim Values As New StringBuilder

         With Values
            For Each Flag As CPU8086Class.FlagRegistersE In [Enum].GetValues(GetType(CPU8086Class.FlagRegistersE))
               If Flag >= &H0% AndAlso Flag <= &HF% Then .Append($"{Flag} = {Abs(CInt(CPU.Registers(Flag)))} ")
            Next Flag
            .Append(NewLine)

            For Each Register As CPU8086Class.Registers16BitE In [Enum].GetValues(GetType(CPU8086Class.Registers16BitE))
               .Append($"{Register} = {CInt(CPU.Registers(Register)):X4} ")
            Next Register
            .Append(NewLine)

            For Each Register As CPU8086Class.SegmentRegistersE In [Enum].GetValues(GetType(CPU8086Class.SegmentRegistersE))
               .Append($"{Register} = {CInt(CPU.Registers(Register)):X4} ")
            Next Register

            Return .ToString()
         End With
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns the emulated CPU's stack contents.
   Private Function GetStack(SS As Integer, SP As Integer) As String
      Try
         Dim Stack As New StringBuilder

         With Stack
            For Offset As Integer = &HFFFD% To SP Step -&H2%
               Stack.Append($"{CPU.GetWord((SS << &H4%) + Offset):X4}{NewLine}")
            Next Offset

            Return .ToString()
         End With
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure loads the specified binary file into the emulated CPU's memory.
   Private Sub LoadBinary(FileName As String, Offset As Integer)
      Try
         Dim Bytes As New List(Of Byte)(File.ReadAllBytes(FileName))

         Output.AppendText($"Loading ""{FileName}"" ({Bytes.Count:X8} bytes) at address {Offset:X8}.{NewLine}")
         If Offset + Bytes.Count <= CPU.Memory.Length Then
            Bytes.CopyTo(CPU.Memory, Offset)
         Else
            Output.AppendText($"""{FileName}"" does not fit inside the emulated memory.{NewLine}")
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure manages the mouse cursor status.
   Public Sub MouseCursorStatus(Busy As Boolean)
      Try
         For Each Window As Form In Application.OpenForms
            Window.Cursor = If(Busy, Cursors.WaitCursor, Cursors.Default)
            Application.DoEvents()
         Next Window
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure parses the user's commands.
   Public Sub ParseCommand(Command As String)
      Try
         Dim Address As New Integer?
         Dim Count As New Integer?
         Dim FileName As String = Nothing
         Dim Is8Bit As New Boolean
         Dim NewValue As New Integer?
         Dim Parsed As New ParsedStr
         Dim Register As Object = GetRegisterByName(Command.Trim().ToUpper(), Is8Bit)
         Dim Value As String = Nothing

         Command = Command.Trim()
         Output.AppendText($"{Command}{NewLine}")

         If Not Command = Nothing Then
            If Register IsNot Nothing Then
               Output.AppendText($"{Register} = {If(Is8Bit, $"{CInt(CPU.Registers(Register)):X2}", $"{CInt(CPU.Registers(Register)):X4}") }{NewLine}")
            ElseIf IS_MEMORY_OPERAND(Command) Then
               Address = AddressFromOperand(Command)
               If Address Is Nothing Then
                  Output.AppendText($"Invalid address.{NewLine}")
               Else
                  Output.AppendText(GetMemoryValue(CInt(Address)))
               End If
            ElseIf Command = "?" Then
               Output.AppendText($"{My.Resources.Help}{NewLine}")
            ElseIf Command.ToUpper() = "C" Then
               Output.Clear()
            ElseIf Command.ToUpper() = "E" Then
               Output.AppendText($"{If(CPU.Clock.Enabled, "Already executing.", "Execution started.")}{NewLine}")
               CPU.Clock.Start()
            ElseIf Command.ToUpper() = "Q" Then
               Application.Exit()
            ElseIf Command.ToUpper() = "R" Then
               Output.AppendText($"{GetRegisterValues()}{NewLine}")
            ElseIf Command.ToUpper() = "S" Then
               Output.AppendText($"{If(Not CPU.Clock.Enabled, "Execution already stopped.", "Execution stopped.")}{NewLine}")
               CPU.Clock.Stop()
            ElseIf Command.ToUpper() = "ST" Then
               Output.AppendText(GetStack(CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.SS)), CInt(CPU.Registers(CPU8086Class.Registers16BitE.SP))))
            ElseIf Command.ToUpper() = "T" Then
               Tracer_Tick(Nothing, Nothing)
            ElseIf Command.ToUpper() = "TE" Then
               Output.AppendText($"{If(Tracer.Enabled, "Already tracing.", "Tracing started.")}{NewLine}")
               Tracer.Start()
            ElseIf Command.ToUpper() = "TS" Then
               Output.AppendText($"{If(Not Tracer.Enabled, "Tracing already stopped.", "Tracing stopped.")}{NewLine}")
               Tracer.Stop()
            ElseIf Command.ToUpper().StartsWith("EXE") Then
               FileName = If(Command.Contains(" "c), Command.Substring(Command.IndexOf(" "c) + 1), RequestFileName("Load executable."))
               If Not FileName = Nothing AndAlso Not LoadExecutable(FileName) Then Output.AppendText($"Invalid executable.{NewLine}")
            ElseIf Command.ToUpper().StartsWith("L") Then
               FileName = If(Command.Contains(" "c), Command.Substring(Command.IndexOf(" "c) + 1), RequestFileName("Load binary."))
               If Not FileName = Nothing Then LoadBinary(FileName, GetFlatCSIP())
            ElseIf Command.ToUpper().StartsWith("MA") Then
               Parsed.Remainder = Command
               Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=" "c, Ending:=" "c)
               Address = AddressFromOperand(Parsed.Element)
               If Address Is Nothing Then Address = GetFlatCSIP()
               Assemble(, StartAddress:=Address)
            ElseIf Command.ToUpper().StartsWith("M") OrElse Command.ToUpper().StartsWith("MD") OrElse Command.ToUpper().StartsWith("MT") Then
               Parsed.Remainder = Command

               Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=" "c, Ending:=" "c)
               Address = AddressFromOperand(Parsed.Element)

               Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=Nothing, Ending:=Nothing)
               Count = GetLiteral(Parsed.Element)

               If Address Is Nothing Then
                  Output.AppendText($"Invalid Or no address specified. CS:IP used instead.{NewLine}")
                  Address = (CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.CS)) << &H4%) Or CInt(CPU.Registers(CPU8086Class.Registers16BitE.IP))
               End If

               If Command.ToUpper().StartsWith("MD") Then
                  Output.AppendText(Disasemble(CPU.Memory.ToList(), CInt(Address), Count))
               Else
                  Output.AppendText(GetMemoryDump(AllHexadecimal:=Not Command.ToUpper().StartsWith("MT"), CInt(Address), Count))
               End If
            ElseIf Command.StartsWith("$"c) Then
               FileName = If(Command.Contains(" "c), Command.Substring(Command.IndexOf(" "c) + 1), RequestFileName("Run script."))
               If Not FileName = Nothing Then RunCommandScript(FileName)
            ElseIf Command.Contains(ASSIGNMENT_OPERATOR) Then
               If Command.StartsWith(MEMORY_OPERAND_START) AndAlso Command.Contains(MEMORY_OPERAND_END) Then
                  Parsed.Remainder = Command
                  Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=Nothing, Ending:=ASSIGNMENT_OPERATOR)
                  Address = AddressFromOperand(Parsed.Element.Trim())

                  If Address Is Nothing Then
                     Output.AppendText($"Invalid address.{NewLine}")
                  Else
                     Value = GET_OPERAND(Command).Trim()
                     If IS_VALUES_OPERAND(Value) Then
                        Value = WITHOUT_DELIMITERS(Value)
                        Output.AppendText($"Finished writing values at 0x{WriteValuesToMemory(Value, CInt(Address)):X8}.{NewLine}")
                     ElseIf IS_STRING_OPERAND(Value) Then
                        Value = WITHOUT_DELIMITERS(Value)
                        Output.AppendText($"Finished writing values at 0x{WriteStringToMemory(Value, CInt(Address)):X8}.{NewLine}")
                     Else
                        NewValue = GetLiteral(Value)
                        If NewValue Is Nothing Then
                           Output.AppendText($"Invalid parameter.{NewLine}")
                        Else
                           WriteValueToMemory(CInt(NewValue), CInt(Address), (CInt(NewValue) < &H100%))
                        End If
                     End If
                  End If
               Else
                  Register = GetRegisterByName(Command.ToUpper().Substring(0, Command.IndexOf(ASSIGNMENT_OPERATOR)).Trim(), Is8Bit)
                  Value = GET_OPERAND(Command).Trim()
                  If Register Is Nothing Then
                     Output.AppendText($"Invalid register.{NewLine}")
                  Else
                     NewValue = GetLiteral(Value, Is8Bit)
                     If NewValue Is Nothing Then
                        Output.AppendText($"Invalid parameter.{NewLine}")
                     Else
                        CPU.Registers(Register, NewValue:=NewValue)
                     End If
                  End If
               End If
            Else
               Output.AppendText($"Error.{NewLine}")
            End If
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure extracts an element from the specified command and returns the result.
   Private Function ParseElement(Command As String, Start As String, Ending As String) As ParsedStr
      Try
         Dim Position As Integer = If(Start = Nothing, 0, Command.IndexOf(Start) + 1)
         Dim NextPosition As Integer = If(Ending = Nothing, Command.Length, Command.IndexOf(Ending, Position + 1))

         If Position < 0 Then Position = 0
         If NextPosition < 0 Then NextPosition = Command.Length

         Return New ParsedStr With {.Element = Command.Substring(Position, NextPosition - Position), .Remainder = If(NextPosition >= Command.Length, "", Command.Substring(NextPosition + 1))}
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return New ParsedStr With {.Remainder = ""}
   End Function

   'This procedure displays a file dialog requesting the user to specify a file.
   Private Function RequestFileName(Title As String) As String
      Try
         With New OpenFileDialog With {.Title = Title}
            If .ShowDialog() = DialogResult.OK Then Return .FileName
         End With
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure runs the commands in the specified script file.
   Public Sub RunCommandScript(FileName As String)
      Try
         Dim PreviousPath As String = Directory.GetCurrentDirectory()
         Dim Script As New List(Of String)(File.ReadAllLines(FileName))

         If Script.First().Trim().ToUpper() = SCRIPT_HEADER Then
            Output.AppendText($"Running script ""{FileName}"".{NewLine}")
            Script.RemoveAt(0)
            Directory.SetCurrentDirectory(Path.GetDirectoryName(FileName))
            Script.ForEach(AddressOf ParseCommand)
            Directory.SetCurrentDirectory(PreviousPath)
         Else
            Output.AppendText($"Invalid script file.{NewLine}")
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure gives the CPU the command to execute an instruction and traces the results.
   Private Sub Tracer_Tick(sender As Object, e As EventArgs) Handles Tracer.Tick
      Try
         Dim Address As New Integer
         Dim Code As String = Disasemble(CPU.Memory.ToList(), GetFlatCSIP(), Count:=&H1%)

         If Not CPU.ExecuteOpcode() Then CPU.ExecuteInterrupt(CPU8086Class.OpcodesE.INT, Number:=CPU8086Class.INVALID_OPCODE)

         Output.AppendText($"{Code}{GetRegisterValues()}{NewLine}")
         If Code.Contains(MEMORY_OPERAND_START) AndAlso Code.Contains(MEMORY_OPERAND_END) Then
            Address = CInt(CPU.AddressFromOperand(CPU8086Class.MemoryOperandsE.LAST))
            Output.AppendText($"{MEMORY_OPERAND_START}0x{Address:X8}{MEMORY_OPERAND_END} =  {GetMemoryValue(Address)}{NewLine}")
         Else
            Output.AppendText($"{NewLine}")
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure writes the specified string to memory at the specified address.
   Private Function WriteStringToMemory([String] As String, Address As Integer) As Integer
      Try
         For Each Character As Char In [String].ToCharArray()
            CPU.Memory(Address) = ToByte(Character)
            Address += &H1%
         Next Character

         Return Address
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure writes the specified values to memory at the specified address.
   Private Function WriteValuesToMemory(Values As String, Address As Integer) As Integer
      Try
         Dim Character As New Char
         Dim InCharacterLiteral As Boolean = False
         Dim Is8Bit As New Boolean
         Dim NewValue As String = Nothing
         Dim Position As Integer = 0
         Dim Value As New Integer?

         Values = $"{Values} "
         Do Until Values.Trim() = Nothing
            Character = Values.Chars(Position)
            Is8Bit = True
            If Character = " "c Then
               NewValue = Values.Substring(0, Position).Trim()
               Value = GetLiteral(NewValue, Is8Bit)
               If Value Is Nothing Then
                  Output.AppendText($"Invalid value: ""{NewValue}"".{NewLine}")
                  Exit Do
               Else
                  Is8Bit = (CInt(Value) < &H100%)
                  Values = $"{Values.Substring(Position).Trim()} "
                  Position = 0
                  WriteValueToMemory(CInt(Value), Address, Is8Bit)
                  Address += If(Is8Bit, &H1%, &H2%)
               End If
            Else
               Position += 1
            End If
         Loop

         Return Address
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure writes the specified value to memory.
   Private Sub WriteValueToMemory(NewValue As Integer, Address As Integer, Is8Bit As Boolean)
      Try
         If Is8Bit Then
            CPU.Memory(Address) = CByte(NewValue And &HFF%)
         Else
            CPU.PutWord(Address, NewValue)
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Module
