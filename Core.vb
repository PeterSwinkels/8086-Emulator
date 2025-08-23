'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Environment
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Math
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
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
   Private Const SCRIPT_COMMENT As Char = "#"c                   'Defines a script comment.
   Private Const SCRIPT_HEADER As String = "[SCRIPT]"            'Defines the script file header.
   Private Const STRING_OPERAND_DELIMITER As Char = """"c        'Defines a string operand's delimiter.
   Private Const VALUES_OPERAND_END As Char = "}"c               'Defines a values operand's end.
   Private Const VALUES_OPERAND_START As Char = "{"c             'Defines a values operand's start.

   Public WithEvents CPU As New CPU8086Class                                                   'Contains a reference to the CPU 8086 class.
   Public WithEvents PIT As New PITClass                                                       'Contains the 8253 Programmable Interval Timer.
   Public WithEvents ScreenRefresh As New Forms.Timer With {.Enabled = True, .Interval = 56}   'Contains the screen refresh timer.
   Private WithEvents Assembler As New AssemblerClass                                          'Contains a reference to the assembler.
   Private WithEvents Disassembler As New DisassemblerClass                                    'Contains a reference to the disassembler.

   Public AssemblyModeOn As Boolean = False                    'Indicates whether input is interpreted as assembly language.
   Public CPUEvent As New StringBuilder                        'Contains CPU event specific text.
   Public CurrentVideoMode As VideoModesE = VideoModesE.None   'Contains the current video mode.
   Public Output As TextBox = Nothing                          'Contains a reference to an output.
   Public ScreenActive As Boolean = False                      'Indicates whether or not the screen window is active.
   Public Synchronizer As New Object                           'Contains the object used to synchronize threads.
   Public VideoAdapter As VideoAdapterClass = Nothing          'Contains a reference to the video adapter used.

   Public ReadOnly CODE_PAGE_437() As Integer = {&H0%, &H263A%, &H263B%, &H2665%, &H2666%, &H2663%, &H2660%, &H2022%, &H25D8%, &H25CB%, &H25D9%, &H2642%, &H2640%, &H266A%, &H266B%, &H263C%, &H25BA%, &H25C4%, &H2195%, &H203C%, &HB6%, &HA7%, &H25AC%, &H21A8%, &H2191%, &H2193%, &H2192%, &H2190%, &H221F%, &H2194%, &H25B2%, &H25BC%, &H20%, &H21%, &H22%, &H23%, &H24%, &H25%, &H26%, &H27%, &H28%, &H29%, &H2A%, &H2B%, &H2C%, &H2D%, &H2E%, &H2F%, &H30%, &H31%, &H32%, &H33%, &H34%, &H35%, &H36%, &H37%, &H38%, &H39%, &H3A%, &H3B%, &H3C%, &H3D%, &H3E%, &H3F%, &H40%, &H41%, &H42%, &H43%, &H44%, &H45%, &H46%, &H47%, &H48%, &H49%, &H4A%, &H4B%, &H4C%, &H4D%, &H4E%, &H4F%, &H50%, &H51%, &H52%, &H53%, &H54%, &H55%, &H56%, &H57%, &H58%, &H59%, &H5A%, &H5B%, &H5C%, &H5D%, &H5E%, &H5F%, &H60%, &H61%, &H62%, &H63%, &H64%, &H65%, &H66%, &H67%, &H68%, &H69%, &H6A%, &H6B%, &H6C%, &H6D%, &H6E%, &H6F%, &H70%, &H71%, &H72%, &H73%, &H74%, &H75%, &H76%, &H77%, &H78%, &H79%, &H7A%, &H7B%, &H7C%, &H7D%, &H7E%, &H2302%, &HC7%, &HFC%, &HE9%, &HE2%, &HE4%, &HE0%, &HE5%, &HE7%, &HEA%, &HEB%, &HE8%, &HEF%, &HEE%, &HEC%, &HC4%, &HC5%, &HC9%, &HE6%, &HC6%, &HF4%, &HF6%, &HF2%, &HFB%, &HF9%, &HFF%, &HD6%, &HDC%, &HA2%, &HA3%, &HA5%, &H20A7%, &H192%, &HE1%, &HED%, &HF3%, &HFA%, &HF1%, &HD1%, &HAA%, &HBA%, &HBF%, &H2310%, &HAC%, &HBD%, &HBC%, &HA1%, &HAB%, &HBB%, &H2591%, &H2592%, &H2593%, &H2502%, &H2524%, &H2561%, &H2562%, &H2556%, &H2555%, &H2563%, &H2551%, &H2557%, &H255D%, &H255C%, &H255B%, &H2510%, &H2514%, &H2534%, &H252C%, &H251C%, &H2500%, &H253C%, &H255E%, &H255F%, &H255A%, &H2554%, &H2569%, &H2566%, &H2560%, &H2550%, &H256C%, &H2567%, &H2568%, &H2564%, &H2565%, &H2559%, &H2558%, &H2552%, &H2553%, &H256B%, &H256A%, &H2518%, &H250C%, &H2588%, &H2584%, &H258C%, &H2590%, &H2580%, &H3B1%, &HDF%, &H393%, &H3C0%, &H3A3%, &H3C3%, &HB5%, &H3C4%, &H3A6%, &H398%, &H3A9%, &H3B4%, &H221E%, &H3C6%, &H3B5%, &H2229%, &H2261%, &HB1%, &H2265%, &H2264%, &H2320%, &H2321%, &HF7%, &H2248%, &HB0%, &H2219%, &HB7%, &H221A%, &H207F%, &HB2%, &H25A0%, &HA0%}  'Contains the code page 437 to unicode mappings.
   Public ReadOnly ESCAPE_BYTE As Func(Of Byte, String) = Function([Byte] As Byte) If([Byte] >= ToByte(" "c) AndAlso [Byte] <= ToByte("~"c), If([Byte] = ESCAPE_CHARACTER, New String(ToChar(ESCAPE_CHARACTER), count:=2), ToChar([Byte])), $"{ToChar(ESCAPE_CHARACTER)}{[Byte]:X2}")    'Returns the specified as either a character or escape sequence.
   Public ReadOnly SET_BIT As Func(Of Integer, Boolean, Integer, Integer) = Function(Value As Integer, Bit As Boolean, Index As Integer) If(Bit, Value Or (&H1% << Index), Value And ((&H1% << Index) Xor &HFFFF%))                                                                      'Returns the specified value with the specified bit set to the specified value.
   Private ReadOnly GET_MEMORY_VALUE As Func(Of Integer, String) = Function(Address As Integer) $"Byte = 0x{CPU.Memory(Address):X2}   Word = 0x{CPU.GET_WORD(Address):X4}   Characters = '{ESCAPE_BYTE(CPU.Memory(Address))}{ESCAPE_BYTE(CPU.Memory(Address + &H1%))}'{NewLine}"         'Returns the specified memory location's value.
   Private ReadOnly GET_OPERAND As Func(Of String, String) = Function(Input As String) (Input.Substring(Input.IndexOf(ASSIGNMENT_OPERATOR) + 1))                                                                                                                                         'Returns the specified input's operand.
   Private ReadOnly IS_CHARACTER_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (Operand.Trim().StartsWith(CHARACTER_OPERAND_DELIMITER) AndAlso Operand.Trim().EndsWith(CHARACTER_OPERAND_DELIMITER) AndAlso Operand.Trim().Length = 3)                               'Indicates whether the specified operand is a character.
   Private ReadOnly IS_MEMORY_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (Operand.Trim().StartsWith(MEMORY_OPERAND_START) AndAlso Operand.Trim().EndsWith(MEMORY_OPERAND_END))                                                                                    'Indicates whether the specified operand is a memory location.
   Private ReadOnly IS_SEGMENT_PREFIX As Func(Of CPU8086Class.OpcodesE, Boolean) = Function(Opcode As CPU8086Class.OpcodesE) {CPU8086Class.OpcodesE.CS, CPU8086Class.OpcodesE.DS, CPU8086Class.OpcodesE.ES, CPU8086Class.OpcodesE.SS}.Contains(Opcode)                                   'Indicates whether the specified opcode is a segment prefix.
   Private ReadOnly IS_STRING_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (Operand.Trim().StartsWith(STRING_OPERAND_DELIMITER) AndAlso Operand.Trim().EndsWith(STRING_OPERAND_DELIMITER))                                                                          'Indicates whether the specified operand is a string.
   Private ReadOnly IS_VALUES_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (Operand.Trim().StartsWith(VALUES_OPERAND_START) AndAlso Operand.Trim().EndsWith(VALUES_OPERAND_END))                                                                                    'Indicates whether the specified operand is an array of values.
   Private ReadOnly REMOVE_DELIMITERS As Func(Of String, String) = Function(Operand As String) (Operand.Substring(1, Operand.Length - 2))                                                                                                                                                'Returns the specified operand with its first and last character removed.

   'This procedure translates the specified memory operands to a flat memory address and returns the result.
   Private Function AddressFromOperand(Operand As String) As Integer?
      Try
         Dim Offset As New Integer?
         Dim Parsed As New ParsedStr
         Dim Segment As New Integer?

         If IS_MEMORY_OPERAND(Operand) Then Operand = REMOVE_DELIMITERS(Operand)

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
         Static PreviousAddress As New Integer

         If AssemblyModeOn Then
            Select Case Input.Trim()
               Case Nothing
                  AssemblyModeOn = False
                  Output.AppendText($"Done.{NewLine}")
               Case "*"c
                  Address = PreviousAddress
                  Output.AppendText($"Address reset to: {Address:X8}.{NewLine}")
               Case "?"c
                  Output.AppendText($"{My.Resources.Assembler}{NewLine}")
               Case Else
                  Opcodes = Assembler.Assemble(Address, Input)
                  If Opcodes Is Nothing Then
                     Output.AppendText($"{Input}{NewLine}")
                  ElseIf Opcodes.Count = 0 Then
                     Output.AppendText($"{Input}{NewLine}")
                     Output.AppendText($"Undefined error.{NewLine}")
                  Else
                     Output.AppendText($"{Address:X8}   {Input,-25} -> {String.Join(Nothing, (From Opcode In Opcodes Select Opcode.ToString("X2")).ToArray())}{NewLine}")
                     Array.Copy(Opcodes.ToArray(), 0, CPU.Memory, Address, Opcodes.Count)
                     PreviousAddress = Address
                     Address = (Address + Opcodes.Count) And CPU8086Class.ADDRESS_MASK
                  End If
            End Select
         ElseIf StartAddress IsNot Nothing Then
            Address = CInt(StartAddress) And CPU8086Class.ADDRESS_MASK
            PreviousAddress = Address
            AssemblyModeOn = True
            Output.AppendText($"Assembler started at 0x{Address:X8}.{NewLine}")
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles any exceptions raised by the assembler.
   Private Sub Assembler_HandleError(AssemblerExceptionO As Exception) Handles Assembler.HandleError
      Try
         Output.AppendText($"Assembler error: {AssemblerExceptionO.Message}{NewLine}")
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles the emulated CPU's halt events.
   Private Sub CPU_Halt() Handles CPU.Halt
      Try
         CPU.ClockToken.Cancel()
         SyncLock Synchronizer
            CPUEvent.Append($"Halted.{NewLine}")
         End SyncLock
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles the emulated CPU's interrupt events.
   Private Sub CPU_Interrupt(Number As Integer, AH As Integer) Handles CPU.Interrupt
      Try
         If Not HandleInterrupt(Number, AH) Then
            CPU.ClockToken.Cancel()
            SyncLock Synchronizer
               CPUEvent.Append($"INT {Number:X}, {AH:X}{NewLine}")
            End SyncLock
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles the emulated CPU's I/O read events.
   Private Sub CPU_ReadIOPort(Port As Integer, ByRef Value As Integer, Is8Bit As Boolean) Handles CPU.ReadIOPort
      Try
         Dim NewValue As Integer? = ReadIOPort(Port)

         If NewValue Is Nothing Then
            CPU.ClockToken.Cancel()
            SyncLock Synchronizer
               CPUEvent.Append($"IN {If(Is8Bit, "AL", "AX")}, {Port:X}{NewLine}")
            End SyncLock
         Else
            Value = CInt(NewValue)
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles CPU tracing events.
   Private Sub CPU_Trace(FlatCSIP As Integer) Handles CPU.Trace
      Try
         Dim Address As New Integer?
         Dim Opcode As New CPU8086Class.OpcodesE
         Dim Override As New CPU8086Class.SegmentRegistersE?
         Dim ParsedAddress As New Integer
         Static Code As String = Nothing
         Static CheckForAddress As Boolean = False
         Static SegmentPrefix As String = Nothing

         SyncLock Synchronizer
            If CheckForAddress Then
               If Code IsNot Nothing AndAlso Code.Contains(MEMORY_OPERAND_START) AndAlso Code.Contains(MEMORY_OPERAND_END) Then
                  Address = CPU.AddressFromOperand(CPU8086Class.MemoryOperandsE.LAST).FlatAddress

                  If Address Is Nothing AndAlso Integer.TryParse(ParseElement(Code, $"{MEMORY_OPERAND_START}{Disassembler.HEXADECIMAL_PREFIX}", MEMORY_OPERAND_END).Element, NumberStyles.HexNumber, CultureInfo.InvariantCulture, ParsedAddress) Then
                     Override = CPU.SegmentOverride(, Preserve:=True)
                     Address = (CInt(CPU.Registers(If(Override Is Nothing, CPU8086Class.SegmentRegistersE.DS, Override))) << &H4%) + ParsedAddress
                  End If

                  GenerateAddressContent(Address)
               ElseIf Code IsNot Nothing AndAlso (Code.Contains("LODSB") OrElse Code.Contains("LODSW")) Then
                  Address = AddressFromOperand("[DS:SI]") + (If(CBool(CPU.Registers(CPU8086Class.FlagRegistersE.DF)), 1, -1))
                  GenerateAddressContent(Address)
               ElseIf Code IsNot Nothing AndAlso {"SCASB", "SCASW", "STOSB", "STOSW"}.Any(Function(OpcodeText As String) Code.Contains(OpcodeText)) Then
                  Address = AddressFromOperand("[ES:DI]") + (If(CBool(CPU.Registers(CPU8086Class.FlagRegistersE.DF)), 1, -1))
                  GenerateAddressContent(Address)
               ElseIf Code IsNot Nothing AndAlso {"CMPSB", "CMPSW", "MOVSB", "MOVSW"}.Any(Function(OpcodeText As String) Code.Contains(OpcodeText)) Then
                  Address = AddressFromOperand("[DS:SI]") + (If(CBool(CPU.Registers(CPU8086Class.FlagRegistersE.DF)), 1, -1))
                  GenerateAddressContent(Address, AddNewLine:=False)

                  Address = AddressFromOperand("[ES:DI]") + (If(CBool(CPU.Registers(CPU8086Class.FlagRegistersE.DF)), 1, -1))
                  GenerateAddressContent(Address)
               Else
                  CPUEvent.Append($"{NewLine}")
               End If
            Else
               Opcode = DirectCast(CPU.Memory(FlatCSIP), CPU8086Class.OpcodesE)

               Select Case Opcode
                  Case CPU8086Class.OpcodesE.REPNE, CPU8086Class.OpcodesE.REPZ
                     Code = Disassemble(CPU.Memory.ToList(), FlatCSIP, Count:=&H2%)
                  Case Else
                     Code = Disassemble(CPU.Memory.ToList(), FlatCSIP, Count:=&H1%)
               End Select

               If Code IsNot Nothing Then
                  If SegmentPrefix IsNot Nothing Then
                     Code = $"{SegmentPrefix}{Code}"
                     SegmentPrefix = Nothing
                  End If
                  If IS_SEGMENT_PREFIX(Opcode) Then
                     CPUEvent.Append($"(segment prefix){NewLine}")
                  Else
                     CPUEvent.Append($"{Code}{GetRegisterValues()}{NewLine}")
                  End If
               End If

               If IS_SEGMENT_PREFIX(Opcode) Then
                  SegmentPrefix = Disassemble(CPU.Memory.ToList(), FlatCSIP, Count:=&H1%)
               End If
            End If

            CheckForAddress = Not CheckForAddress
         End SyncLock
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles the emulated CPU's I/O write events.
   Private Sub CPU_WriteIOPort(Port As Integer, Value As Integer, Is8Bit As Boolean) Handles CPU.WriteIOPort
      Try
         If Not WriteIOPort(Port, Value) Then
            CPU.ClockToken.Cancel()
            SyncLock Synchronizer
               CPUEvent.Append($"OUT {Port:X}, {If(Is8Bit, $"{Value:X2}", $"{Value:X4}")}{NewLine}")
            End SyncLock
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure disassembles the specified code and returns the resulting lines of code.
   Private Function Disassemble(Code As List(Of Byte), Position As Integer, Optional Count As Integer? = Nothing) As String
      Try
         Dim Disassembly As New StringBuilder
         Dim EndPosition As New Integer
         Dim HexadecimalCode As String = Nothing
         Dim Instruction As String = Nothing
         Dim PreviousPosition As New Integer

         If Count Is Nothing Then Count = DEFAULT_DISASSEMBLY_COUNT

         EndPosition = If(CInt(Count) = &H0%, Code.Count, Position + CInt(Count))

         With Disassembler
            Do Until Position >= EndPosition OrElse Position > CPU.Memory.GetUpperBound(0)
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

   'This procedure generates the output for the contents of the specified memory address.
   Private Sub GenerateAddressContent(Address As Integer?, Optional AddNewLine As Boolean = True)
      Try
         If Address Is Nothing Then
            CPUEvent.Append($"{MEMORY_OPERAND_START}0x???{MEMORY_OPERAND_END} = ???{NewLine}")
         Else
            Address = Address And CPU8086Class.ADDRESS_MASK
            CPUEvent.Append($"{MEMORY_OPERAND_START}0x{Address:X8}{MEMORY_OPERAND_END} = {GET_MEMORY_VALUE(CInt(Address))}")
         End If

         If AddNewLine Then CPUEvent.Append(NewLine)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

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
               If Address IsNot Nothing Then Literal = If(Is8Bit, CPU.Memory(CInt(Address)), CPU.GET_WORD(CInt(Address)))
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
               CPU.Memory.ToList().GetRange(Offset, CInt(Count)).ForEach(Sub([Byte] As Byte) .Append(ESCAPE_BYTE([Byte])))
            End If

            .Append($"{NewLine}{NewLine}")
         End With

         Return Dump.ToString()
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
            For Each Register As CPU8086Class.Registers16BitE In [Enum].GetValues(GetType(CPU8086Class.Registers16BitE))
               .Append($"{Register} = {CInt(CPU.Registers(Register)):X4} ")
            Next Register
            .Append(NewLine)

            For Each Register As CPU8086Class.SegmentRegistersE In [Enum].GetValues(GetType(CPU8086Class.SegmentRegistersE))
               .Append($"{Register} = {CInt(CPU.Registers(Register)):X4} ")
            Next Register
            .Append(NewLine)

            For Each Flag As CPU8086Class.FlagRegistersE In [Enum].GetValues(GetType(CPU8086Class.FlagRegistersE))
               If Flag >= &H0% AndAlso Flag <= &HF% Then .Append($"{Flag} = {Abs(CInt(CPU.Registers(Flag)))} ")
            Next Flag

            Return .ToString()
         End With
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns the emulated CPU's stack contents.
   Private Function GetStack() As String
      Try
         Dim SS As Integer = CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.SS))
         Dim Stack As New StringBuilder

         With Stack
            For Offset As Integer = CInt(CPU.Registers(CPU8086Class.Registers16BitE.BP)) To CInt(CPU.Registers(CPU8086Class.Registers16BitE.SP))
               Stack.Append($"{CPU.GET_WORD((SS << &H4%) + Offset):X4}{NewLine}")
            Next Offset

            Return .ToString()
         End With
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure reads a null terminated string from the specified position in memory and returns it.
   Public Function GetStringZ(Segment As Integer, Offset As Integer) As String
      Try
         Dim Position As Integer = (Segment << &H4%) + Offset
         Dim StringZ As New StringBuilder

         Do Until CPU.Memory(Position And CPU8086Class.ADDRESS_MASK) = &H0%
            StringZ.Append(ToChar(CPU.Memory(Position And CPU8086Class.ADDRESS_MASK)))
            Position += &H1%
         Loop

         Return StringZ.ToString()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure loads the specified binary file into the emulated CPU's memory.
   Private Function LoadBinary(FileName As String, Offset As Integer) As Boolean
      Try
         Dim Binary As New List(Of Byte)(File.ReadAllBytes(FileName))
         Dim Success As Boolean = True

         If Offset + Binary.Count <= CPU.Memory.Length Then
            Output.AppendText($"Loading ""{FileName}"" ({Binary.Count:X8} bytes) at address {Offset:X8}.{NewLine}")
            Binary.CopyTo(CPU.Memory, Offset)
         Else
            Output.AppendText($"""{FileName}"" does not fit inside the emulated memory.{NewLine}")
            Success = False
         End If

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function

   'This procedure manages the mouse cursor status.
   Public Sub MouseCursorStatus(Busy As Boolean)
      Try
         Dim Forms As New List(Of Form)(Application.OpenForms.Cast(Of Form)())

         For Each [Form] As Form In Forms
            [Form].Cursor = If(Busy, Cursors.WaitCursor, Cursors.Default)
            Application.DoEvents()
         Next [Form]
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure parses the specified command and returns whether or not a critical error occurred.
   Public Function ParseCommand(Input As String) As Boolean
      Try
         Dim Address As New Integer?
         Dim AH As New Integer?
         Dim Command As String = Nothing
         Dim Count As New Integer?
         Dim ErrorAt As New Integer
         Dim FileName As String = Nothing
         Dim Interrupt As New Integer?
         Dim Is8Bit As New Boolean
         Dim NewValue As New Integer?
         Dim Operands As String = Nothing
         Dim Parsed As New ParsedStr
         Dim Position As New Integer
         Dim Register As Object = GetRegisterByName(Input.Trim().ToUpper(), Is8Bit)
         Dim Success As Boolean = True
         Dim Value As String = Nothing

         Input = Input.Trim()

         If Not Input = Nothing Then
            Output.AppendText($"{Input}{NewLine}")

            If Not Input.StartsWith(SCRIPT_COMMENT) Then
               Position = Input.IndexOf(" "c)
               If Position < 0 Then
                  Command = Input.ToUpper()
               Else
                  Command = Input.Substring(0, Position).ToUpper()
                  Operands = Input.Substring(Position + 1).Trim()
               End If

               If Register IsNot Nothing Then
                  Output.AppendText($"{Register} = {If(Is8Bit, $"{CInt(CPU.Registers(Register)):X2}", $"{CInt(CPU.Registers(Register)):X4}") }{NewLine}")
               ElseIf IS_MEMORY_OPERAND(Input) Then
                  Address = AddressFromOperand(Input)
                  Output.AppendText(If(Address Is Nothing, $"Invalid address.{NewLine}", GET_MEMORY_VALUE(CInt(Address))))
               Else
                  Select Case Command
                     Case "$"
                        FileName = If(Operands Is Nothing, RequestFileName("Run script."), Operands)
                        If Not FileName = Nothing Then RunCommandScript(FileName)
                     Case "?"
                        Output.AppendText($"{My.Resources.Help}{NewLine}")
                     Case "ARG"
                        CommandTail = $" {If(Operands.Length > COMMAND_TAIL_MAXIMUM_LENGTH, Operands.Substring(1, COMMAND_TAIL_MAXIMUM_LENGTH), Operands)}"
                        Output.AppendText($"Command tail set to: ""{CommandTail}""{NewLine}")
                     Case "C"
                        Output.Clear()
                     Case "CD"
                        If Operands Is Nothing Then
                           Output.AppendText(CurrentDirectory())
                        Else
                           CurrentDirectory = Operands
                        End If
                     Case "E"
                        If CPU.Clock.Status = TaskStatus.Running Then
                           Output.AppendText("Already executing.")
                        Else
                           CPU.ClockToken = New CancellationTokenSource
                           CPU.Clock = New Task(AddressOf CPU.Execute)
                           CPU.Clock.Start()
                           Output.AppendText("Execution started.")
                        End If
                        Output.AppendText(NewLine)
                     Case "ECXZ"
                        Do Until CInt(CPU.Registers(CPU8086Class.Registers16BitE.CX)) = &H0%
                           Application.DoEvents()

                           If Not CPU.ExecuteOpcode() Then
                              CPU.ExecuteInterrupt(CPU8086Class.OpcodesE.INT, Number:=CPU8086Class.INVALID_OPCODE)
                           End If
                        Loop

                        Output.AppendText($"CX = 0{NewLine}")
                     Case "EXE"
                        FileName = If(Operands Is Nothing, RequestFileName("Load Executable."), Operands)
                        If Not FileName = Nothing Then Success = LoadMSDOSProgram(FileName)
                     Case "INT"
                        Parsed.Remainder = Input

                        Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=" "c, Ending:=" "c)
                        Interrupt = GetLiteral(Parsed.Element)
                        If Interrupt Is Nothing Then
                           Output.AppendText($"Invalid interrupt number.{NewLine}")
                        Else
                           Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=Nothing, Ending:=Nothing)
                           AH = GetLiteral(Parsed.Element)
                           If AH Is Nothing Then
                              Output.AppendText($"Invalid function number.{NewLine}")
                           Else
                              CPU.Registers(CPU8086Class.SubRegisters8BitE.AH, NewValue:=AH)

                              If CPU.Clock.Status = TaskStatus.Running Then
                                 CPU.ExecuteInterrupt(CPU8086Class.OpcodesE.INT, Interrupt)
                              Else
                                 CPU_Interrupt(CInt(Interrupt), CInt(AH))
                              End If
                           End If
                        End If
                     Case "IRET"
                        CPU.ExecuteOpcode(CPU8086Class.OpcodesE.IRET)
                     Case "L"
                        Address = (CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H4%) + (CInt(CPU.Registers(CPU8086Class.Registers16BitE.DI))) And CPU8086Class.ADDRESS_MASK
                        FileName = If(Operands Is Nothing, RequestFileName("Load binary."), Operands)
                        If Not FileName = Nothing Then Success = LoadBinary(FileName, CInt(Address))
                     Case "M", "MD", "MT"
                        Parsed.Remainder = Input

                        Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=" "c, Ending:=" "c)
                        Address = AddressFromOperand(Parsed.Element)

                        Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=Nothing, Ending:=Nothing)
                        Count = GetLiteral(Parsed.Element)

                        If Address Is Nothing Then
                           Output.AppendText($"Invalid or no address specified. CS:IP used instead.{NewLine}")
                           Address = (CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.CS)) << &H4%) + CInt(CPU.Registers(CPU8086Class.Registers16BitE.IP)) And CPU8086Class.ADDRESS_MASK
                        End If

                        Output.AppendText(If(Input.ToUpper().StartsWith("MD"), Disassemble(CPU.Memory.ToList(), CInt(Address), Count), GetMemoryDump(AllHexadecimal:=Not Input.ToUpper().StartsWith("MT"), CInt(Address), Count)))
                     Case "MA"
                        Parsed.Remainder = Input
                        Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=" "c, Ending:=" "c)
                        Address = AddressFromOperand(Parsed.Element)

                        If Address Is Nothing Then
                           Output.AppendText($"Invalid or no address specified. CS:IP used instead.{NewLine}")
                           Address = CPU.GET_FLAT_CS_IP()
                        Else
                           CPU.Registers(CPU8086Class.SegmentRegistersE.CS, NewValue:=(Address >> &H4%))
                           CPU.Registers(CPU8086Class.Registers16BitE.IP, NewValue:=Address - (Address And &HFFF0%))
                        End If

                        Assemble(, StartAddress:=Address)
                     Case "Q"
                        Application.Exit()
                     Case "R"
                        Output.AppendText($"{GetRegisterValues()}{NewLine}")
                     Case "RESET"
                        LastBIOSKeyCode(, Clear:=True)
                        ResetMSDOS()

                        CPU.ClockToken.Cancel()
                        CPU = New CPU8086Class
                        CursorPositionUpdate()
                        LoadBIOS()
                        LoadMSDOS()
                        Output.AppendText($"CPU reset.{NewLine}")
                     Case "S"
                        CPU.Tracing = False
                        Output.AppendText($"{If(Not CPU.Clock.Status = TaskStatus.Running, "Execution already stopped.", "Execution stopped.")}{NewLine}")
                        CPU.ClockToken.Cancel()
                     Case "SCR"
                        ScreenWindow.Show()
                     Case "ST"
                        Output.AppendText(GetStack())
                     Case "T"
                        CPU_Trace(CPU.GET_FLAT_CS_IP())

                        If Not CPU.ExecuteOpcode() Then
                           CPU.ExecuteInterrupt(CPU8086Class.OpcodesE.INT, Number:=CPU8086Class.INVALID_OPCODE)
                        End If

                        CPU_Trace(Nothing)
                     Case "TE"
                        If Not CPU.Clock.Status = TaskStatus.Running Then
                           CPU.Tracing = True
                           CPU.ClockToken = New CancellationTokenSource
                           CPU.Clock = New Task(AddressOf CPU.Execute)
                           CPU.Clock.Start()
                        End If
                     Case "TS"
                        CPU.ClockToken.Cancel()
                        CPU.Tracing = False
                        Output.AppendText($"Tracing {If(CPU.Clock.Status = TaskStatus.Running, "stopped.", " is not active.")}{NewLine}")
                     Case Else
                        If Input.Contains(ASSIGNMENT_OPERATOR) Then
                           If Input.StartsWith(MEMORY_OPERAND_START) AndAlso Input.Contains(MEMORY_OPERAND_END) Then
                              Parsed.Remainder = Input
                              Parsed = ParseElement(Parsed.Remainder.Trim(), Start:=Nothing, Ending:=ASSIGNMENT_OPERATOR)
                              Address = AddressFromOperand(Parsed.Element.Trim())

                              If Address Is Nothing Then
                                 Output.AppendText($"Invalid address.{NewLine}")
                              Else
                                 Value = GET_OPERAND(Input).Trim()
                                 If IS_VALUES_OPERAND(Value) Then
                                    Value = REMOVE_DELIMITERS(Value)
                                    Output.AppendText($"Finished writing values at 0x{WriteValuesToMemory(Value, CInt(Address)):X8}.{NewLine}")
                                 ElseIf IS_STRING_OPERAND(Value) Then
                                    Value = Unescape(REMOVE_DELIMITERS(Value),, ErrorAt)
                                    Output.AppendText(If(ErrorAt = 0, $"Finished writing string at 0x{WriteStringToMemory(Value, CInt(Address)):X8}.{NewLine}", $"Invalid escape sequence at: {ErrorAt}.{NewLine}"))
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
                              Register = GetRegisterByName(Input.ToUpper().Substring(0, Input.IndexOf(ASSIGNMENT_OPERATOR)).Trim(), Is8Bit)
                              Value = GET_OPERAND(Input).Trim()
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
                  End Select
               End If
            End If
         End If

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function

   'This procedure extracts an element from the specified text and returns the result.
   Private Function ParseElement(Text As String, Start As String, Ending As String) As ParsedStr
      Try
         Dim Position As Integer = If(Start = Nothing, 0, Text.IndexOf(Start) + Start.Length)
         Dim NextPosition As Integer = If(Ending = Nothing, Text.Length, Text.IndexOf(Ending, Position + 1))

         If Position < 0 Then Position = 0
         If NextPosition < 0 Then NextPosition = Text.Length

         Return New ParsedStr With {.Element = Text.Substring(Position, NextPosition - Position), .Remainder = If(NextPosition >= Text.Length, "", Text.Substring(NextPosition + 1))}
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return New ParsedStr With {.Remainder = ""}
   End Function

   'This procedure handles PIT events.
   Private Sub PIT_PITEvent(Counter As PITClass.CountersE, Mode As PITClass.ModesE) Handles PIT.PITEvent
      Try
         If Not CPU.UpdateClock Then
            CPU.UpdateClock = ((Counter = PITClass.CountersE.TimeOfDay) AndAlso (CPU.Clock.Status = TaskStatus.Running))
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure displays a file dialog requesting the user to specify a file and returns the user's selection.
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
         Dim ResetPath As Boolean = True
         Dim Script As New List(Of String)(File.ReadAllLines(FileName))

         If Script.First().Trim().ToUpper() = SCRIPT_HEADER Then
            Output.AppendText($"Running script ""{FileName}"".{NewLine}")
            Script.RemoveAt(0)
            Directory.SetCurrentDirectory(Path.GetDirectoryName(FileName))
            For Each Line As String In Script
               If AssemblyModeOn Then
                  Assemble(Line)
               Else
                  If Line.Trim().ToUpper().StartsWith("CD") Then ResetPath = False
                  If Not ParseCommand(Line) Then Exit For
               End If
            Next Line
            If ResetPath Then Directory.SetCurrentDirectory(PreviousPath)
         Else
            Output.AppendText($"Invalid script file.{NewLine}")
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure refreshes the screen's output.
   Private Sub ScreenRefresh_Tick(sender As Object, e As EventArgs) Handles ScreenRefresh.Tick
      Try
         Dim VideoMode As VideoModesE = DirectCast(CPU.Memory(AddressesE.VideoMode), VideoModesE)

         If VideoMode = CurrentVideoMode Then
            If ScreenActive Then ScreenWindow.Invalidate()
         Else
            CurrentVideoMode = VideoMode
            SwitchVideoAdapter()
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure switches to a video adapter based on the current screen mode.
   Public Sub SwitchVideoAdapter()
      Try
         Select Case CurrentVideoMode
            Case VideoModesE.CGA320x200A, VideoModesE.CGA320x200B
               VideoAdapter = New CGA320x200Class
            Case VideoModesE.Text80x25Mono
               VideoAdapter = New Text80x25MonoClass
            Case VideoModesE.VGA320x200
               VideoAdapter = New VGA320x200Class
            Case Else
               VideoAdapter = Nothing
         End Select

         If VideoAdapter IsNot Nothing Then VideoAdapter.Initialize()

         If ScreenActive Then ScreenWindow.Invalidate()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure converts any escape sequences in the specified text to characters and returns the result.
   Private Function Unescape(Text As String, Optional EscapeCharacter As Char = "/"c, Optional ByRef ErrorAt As Integer = 0) As String
      Try
         Dim Character As New Char
         Dim CharacterCode As New Integer
         Dim Index As Integer = 0
         Dim NextCharacter As New Char
         Dim Unescaped As New StringBuilder

         ErrorAt = 0

         Do Until Index >= Text.Length OrElse ErrorAt > 0
            Character = Text.Chars(Index)
            If Index < Text.Length - 1 Then NextCharacter = Text.Chars(Index + 1) Else NextCharacter = Nothing

            If Character = EscapeCharacter Then
               If NextCharacter = EscapeCharacter Then
                  Unescaped.Append(Character)
                  Index += 1
               Else
                  If NextCharacter = Nothing Then
                     ErrorAt = Index + 1
                  Else
                     If Index < Text.Length - 2 AndAlso Integer.TryParse(Text.Substring(Index + 1, 2), NumberStyles.HexNumber, Nothing, CharacterCode) Then
                        Unescaped.Append(ToChar(CharacterCode))
                        Index += 2
                     Else
                        ErrorAt = Index + 1
                     End If
                  End If
               End If
            Else
               Unescaped.Append(Character)
            End If
            Index += 1
         Loop

         Return Unescaped.ToString()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure writes the specified bytes to memory at the specified address and returns the last address written to.
   Public Function WriteBytesToMemory(Bytes() As Byte, Address As Integer) As Integer
      Try
         Array.Copy(Bytes, &H0%, CPU.Memory, Address, Bytes.Count)

         Return Address + Bytes.Count
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure writes the specified string to memory at the specified address and returns the last address written to.
   Public Function WriteStringToMemory([String] As String, Address As Integer) As Integer
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

   'This procedure writes the specified values to memory at the specified address and returns the last address written to.
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
