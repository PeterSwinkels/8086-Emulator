'This class' imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Linq
Imports System.Text

'This class contains the DAsm 8086 disassembler.
Public Class DisassemblerClass
   'These lists contain the names of instructions that share the same opcode.
   Private ReadOnly OPCODES_707F As New List(Of String)({"JO", "JNO", "JC", "JNC", "JZ", "JNZ", "JNA", "JA", "JS", "JNS", "JPE", "JPO", "JL", "JNL", "JNG", "JG"})
   Private ReadOnly OPCODES_8083 As New List(Of String)({"ADD", "OR", "ADC", "SBB", "AND", "SUB", "XOR", "CMP"})
   Private ReadOnly OPCODES_8F As New List(Of String)({"POP", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing})
   Private ReadOnly OPCODES_C0C1_D2D3 As New List(Of String)({"ROL", "ROR", "RCL", "RCR", "SHL", "SHR", Nothing, "SAR"})
   Private ReadOnly OPCODES_C6C7 As New List(Of String)({"MOV", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing})
   Private ReadOnly OPCODES_F6F7 As New List(Of String)({"TEST", Nothing, "NOT", "NEG", "MUL", "IMUL", "DIV", "IDIV"})
   Private ReadOnly OPCODES_FE__00BF As New List(Of String)({"INC BYTE", "DEC BYTE", "CALL WORD NEAR", "CALL WORD FAR", "JMP WORD NEAR", "JMP WORD FAR", "PUSH WORD", Nothing})
   Private ReadOnly OPCODES_FE__C0FF As New List(Of String)({"INC BYTE", "DEC BYTE", "CALL", Nothing, "JMP", Nothing, "PUSH", Nothing})
   Private ReadOnly OPCODES_FF__00BF As New List(Of String)({"INC WORD", "DEC WORD", "CALL WORD NEAR", "CALL WORD FAR", "JMP WORD NEAR", "JMP WORD FAR", "PUSH WORD", Nothing})
   Private ReadOnly OPCODES_FF__C0FF As New List(Of String)({"INC WORD", "DEC WORD", "CALL", Nothing, "JMP", Nothing, "PUSH", Nothing})

   'These lists contain the CPU register names.
   Private ReadOnly LH_REGISTERS As New List(Of String)({"AL", "CL", "DL", "BL", "AH", "CH", "DH", "BH"})
   Private ReadOnly SG_REGISTERS As New List(Of String)({"ES", "CS", "SS", "DS", Nothing, Nothing, Nothing, Nothing})
   Private ReadOnly XP_REGISTERS As New List(Of String)({"AX", "CX", "DX", "BX", "SP", "BP", "SI", "DI"})

   'This list contains the names of registers referring to memory locations.
   Private ReadOnly MEMORY_OPERANDS As New List(Of String)({"BX + SI", "BX + DI", "BP + SI", "BP + DI", "SI", "DI", "BP", "BX"})

   'These constants define the indexes of specific registers in their respective lists.
   Private Const ACCUMULATOR_REGISTER As Integer = &H0%
   Private Const BASE_POINTER As Integer = &H6%
   Private Const COUNTER_REGISTER As Integer = &H1%
   Private Const DATA_REGISTER As Integer = &H2%

   'This constant defines the hexadecimal number prefix.
   Public ReadOnly HEXADECIMAL_PREFIX As String = "0x"

   'The definitions for events that can be raised by this class.
   Public Event HandleError(ExceptionO As Exception)

   'This procedure returns the specified bytes as a hexadecimal number.
   Public Function BytesToHexadecimal(Bytes As List(Of Byte), Optional NoPrefix As Boolean = False, Optional Reverse As Boolean = True, Optional Signed As Boolean = False) As String
      Try
         Dim Hexadecimal As New StringBuilder()

         If Bytes Is Nothing OrElse Not Bytes.Any Then
            Hexadecimal.Append("?")
         Else
            If Reverse Then Bytes.Reverse()

            Bytes.ForEach(Sub(ByteO As Byte) Hexadecimal.Append($"{ByteO:X2}"))
         End If

         If Not NoPrefix Then Hexadecimal = New StringBuilder($"{HEXADECIMAL_PREFIX}{Hexadecimal}")
         If Signed Then Hexadecimal = New StringBuilder(SignHexadecimal(Hexadecimal.ToString()))

         Return Hexadecimal.ToString()
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure disassembles the opcodes at the specified position in the specified code.
   Public Overloads Function Disassemble(Code As List(Of Byte), ByRef Optional Position As Integer = &H0%) As String
      Try
         Dim Instruction As String = Nothing
         Dim Opcode As New Byte
         Dim Operand As New Byte
         Dim PreviousPosition As New Integer

         If Position < Code.Count Then
            Opcode = GetByte(Code, Position)
            PreviousPosition = Position

            Select Case Opcode
               Case &H27% : Instruction = "DAA"
               Case &H2F% : Instruction = "DAS"
               Case &H37% : Instruction = "AAA"
               Case &H3F% : Instruction = "AAS"
               Case &H90% : Instruction = "NOP"
               Case &H98% : Instruction = "CBW"
               Case &H99% : Instruction = "CWD"
               Case &H9B% : Instruction = "WAIT"
               Case &H9C% : Instruction = "PUSHFW"
               Case &H9D% : Instruction = "POPFW"
               Case &H9E% : Instruction = "SAHF"
               Case &H9F% : Instruction = "LAHF"
               Case &HA4% : Instruction = "MOVSB"
               Case &HA5% : Instruction = "MOVSW"
               Case &HA6% : Instruction = "CMPSB"
               Case &HA7% : Instruction = "CMPSW"
               Case &HAA% : Instruction = "STOSB"
               Case &HAB% : Instruction = "STOSW"
               Case &HAC% : Instruction = "LODSB"
               Case &HAD% : Instruction = "LODSW"
               Case &HAE% : Instruction = "SCASB"
               Case &HAF% : Instruction = "SCASW"
               Case &HC3% : Instruction = "RETN"
               Case &HCB% : Instruction = "RETF"
               Case &HCE% : Instruction = "INTO"
               Case &HCF% : Instruction = "IRET"
               Case &HD6% : Instruction = "SALC"
               Case &HD7% : Instruction = "XLATB"
               Case &HF0% : Instruction = "LOCK"
               Case &HF2% : Instruction = "REPNE"
               Case &HF3% : Instruction = "REPZ"
               Case &HF4% : Instruction = "HLT"
               Case &HF5% : Instruction = "CMC"
               Case &HF8% : Instruction = "CLC"
               Case &HF9% : Instruction = "STC"
               Case &HFA% : Instruction = "CLI"
               Case &HFB% : Instruction = "STI"
               Case &HFC% : Instruction = "CLD"
               Case &HFD% : Instruction = "STD"
               Case &H0% To &H5%, &H8% To &HD%, &H10% To &H15%, &H18% To &H1D%, &H20% To &H25%, &H28% To &H2D%, &H30% To &H35%, &H38% To &H3D%, &H88% To &H8B%
                  Select Case Opcode
                     Case &H0% To &H5% : Instruction = "ADD"
                     Case &H8% To &HD% : Instruction = "OR"
                     Case &H10% To &H15% : Instruction = "ADC"
                     Case &H18% To &H1D% : Instruction = "SBB"
                     Case &H20% To &H25% : Instruction = "AND"
                     Case &H28% To &H2D% : Instruction = "SUB"
                     Case &H30% To &H35% : Instruction = "XOR"
                     Case &H38% To &H3D% : Instruction = "CMP"
                     Case &H88% To &H8B% : Instruction = "MOV"
                  End Select

                  Operand = GetByte(Code, Position)
                  Select Case (Opcode And &H7%)
                     Case &H0% : Instruction &= $" {GetOperand(Code, Position, Operand, LH_REGISTERS)}, {LH_REGISTERS((Operand And &H38%) >> &H3%)}"
                     Case &H1% : Instruction &= $" {GetOperand(Code, Position, Operand, XP_REGISTERS)}, {XP_REGISTERS((Operand And &H38%) >> &H3%)}"
                     Case &H2% : Instruction &= $" {LH_REGISTERS((Operand And &H38%) >> &H3%)}, {GetOperand(Code, Position, Operand, LH_REGISTERS)}"
                     Case &H3% : Instruction &= $" {XP_REGISTERS((Operand And &H38%) >> &H3%)}, {GetOperand(Code, Position, Operand, XP_REGISTERS)}"
                     Case &H4% : Instruction &= $" {LH_REGISTERS(ACCUMULATOR_REGISTER)}, {BytesToHexadecimal({Operand}.ToList())}"
                     Case &H5% : Instruction &= $" {XP_REGISTERS(ACCUMULATOR_REGISTER)}, {BytesToHexadecimal({Operand, GetByte(Code, Position)}.ToList())}"
                  End Select
               Case &H6%, &HE%, &H16%, &H1E% : Instruction = $"PUSH {SG_REGISTERS((Opcode And &H18%) >> &H3%)}"
               Case &H7%, &H17%, &H1F% : Instruction = $"POP {SG_REGISTERS((Opcode And &H18%) >> &H3%)}"
               Case &H26%, &H2E%, &H36%, &H3E% : Instruction = SG_REGISTERS((Opcode And &H18%) >> &H3%)
               Case &H40% To &H47% : Instruction = $"INC {XP_REGISTERS(Opcode And &H7%)}"
               Case &H48% To &H4F% : Instruction = $"DEC {XP_REGISTERS(Opcode And &H7%)}"
               Case &H50% To &H57% : Instruction = $"PUSH {XP_REGISTERS(Opcode And &H7%)}"
               Case &H58% To &H5F% : Instruction = $"POP {XP_REGISTERS(Opcode And &H7%)}"
               Case &H64%, &H65% : Instruction = SG_REGISTERS(Opcode And &H7%)
               Case &H70% To &H7F% : Instruction = $"{OPCODES_707F(Opcode And &HF%)} {ShortAddressToHexadecimal(GetBytes(Code, Position, Length:=1), Position)}"
               Case &H80%
                  Operand = GetByte(Code, Position)
                  Instruction = $"{OPCODES_8083((Operand And &H3F%) >> &H3%)} {GetOperand(Code, Position, Operand, LH_REGISTERS)}, {BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}"
               Case &H81%
                  Operand = GetByte(Code, Position)
                  Instruction = $"{OPCODES_8083((Operand And &H3F%) >> &H3%)} {GetOperand(Code, Position, Operand, XP_REGISTERS)}, {BytesToHexadecimal(GetBytes(Code, Position, Length:=2))}"
               Case &H83%
                  Operand = GetByte(Code, Position)
                  Instruction = $"{OPCODES_8083((Operand And &H3F%) >> &H3%)} {If(Operand < &HC0%, " WORD ", Nothing)}{GetOperand(Code, Position, Operand, XP_REGISTERS)}, {If(Operand < &HC0%, " BYTE ", Nothing)}{BytesToHexadecimal(GetBytes(Code, Position, Length:=1), , Signed:=True)}"
               Case &H84%, &H86%
                  If Opcode = &H84% Then Instruction = "TEST "
                  If Opcode = &H86% Then Instruction = "XCHG "

                  Operand = GetByte(Code, Position)
                  Instruction &= $"{GetOperand(Code, Position, Operand, LH_REGISTERS)}, {LH_REGISTERS((Operand And &H38%) >> &H3%)}"
               Case &H85%, &H87%
                  If Opcode = &H85% Then Instruction = "TEST "
                  If Opcode = &H87% Then Instruction = "XCHG "

                  Operand = GetByte(Code, Position)
                  Instruction &= $"{GetOperand(Code, Position, Operand, XP_REGISTERS)}, {XP_REGISTERS((Operand And &H38%) >> &H3%)}"
               Case &H8C%
                  Operand = GetByte(Code, Position)
                  Instruction = $"MOV {GetOperand(Code, Position, Operand, XP_REGISTERS)}, {SG_REGISTERS((Operand And &H38%) >> &H3%)}"
               Case &H8E%
                  Operand = GetByte(Code, Position)
                  Instruction = $"MOV {SG_REGISTERS((Operand And &H38%) >> &H3%)}, {GetOperand(Code, Position, Operand, XP_REGISTERS)}"
               Case &H8F%, &HFF%
                  Operand = GetByte(Code, Position)

                  If Opcode = &H8F% Then Instruction = OPCODES_8F((Operand And &H3F%) >> &H3%)
                  If Opcode = &HFF% AndAlso Operand < &HC0% Then Instruction = OPCODES_FF__00BF((Operand And &H3F%) >> &H3%)
                  If Opcode = &HFF% AndAlso Operand > &HBF% Then Instruction = OPCODES_FF__C0FF((Operand And &H3F%) >> &H3%)
                  If Not Instruction = Nothing Then Instruction &= $" {GetOperand(Code, Position, Operand, XP_REGISTERS)}"
               Case &H91% To &H97% : Instruction = $"XCHG {XP_REGISTERS(ACCUMULATOR_REGISTER)}, {XP_REGISTERS(Opcode And &H7%)}"
               Case &H9A% : Instruction = $"CALL FAR {FarAddressToHexadecimal(GetBytes(Code, Position, Length:=4))}"
               Case &HA0% : Instruction = $"MOV {LH_REGISTERS(ACCUMULATOR_REGISTER)}, [{BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}]"
               Case &HA1% : Instruction = $"MOV {XP_REGISTERS(ACCUMULATOR_REGISTER)}, [{BytesToHexadecimal(GetBytes(Code, Position, Length:=2))}]"
               Case &HA2% : Instruction = $"MOV [{BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}], {LH_REGISTERS(ACCUMULATOR_REGISTER)}"
               Case &HA3% : Instruction = $"MOV [{BytesToHexadecimal(GetBytes(Code, Position, Length:=2))}], {XP_REGISTERS(ACCUMULATOR_REGISTER)}"
               Case &HA8% : Instruction = $"TEST {LH_REGISTERS(ACCUMULATOR_REGISTER)}, {BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}"
               Case &HA9% : Instruction = $"TEST {XP_REGISTERS(ACCUMULATOR_REGISTER)}, {BytesToHexadecimal(GetBytes(Code, Position, Length:=2))}"
               Case &HB0% To &HB7% : Instruction = $"MOV {LH_REGISTERS(Opcode And &H7%)}, {BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}"
               Case &HB8% To &HBF% : Instruction = $"MOV {XP_REGISTERS(Opcode And &H7%)}, {BytesToHexadecimal(GetBytes(Code, Position, Length:=2))}"
               Case &HC2% : Instruction = $"RETN {BytesToHexadecimal(GetBytes(Code, Position, Length:=2))}"
               Case &HC4% To &HC5%, &H8D%
                  If Opcode = &HC4% Then Instruction = "LES "
                  If Opcode = &HC5% Then Instruction = "LDS "
                  If Opcode = &H8D% Then Instruction = "LEA "

                  Operand = GetByte(Code, Position)
                  If Operand < &HC0& Then
                     Instruction &= $"{GetOperand(Code, Position, Operand, XP_REGISTERS)}, {XP_REGISTERS((Operand And &H38%) >> &H3%)}"
                  Else
                     Instruction = Nothing
                  End If
               Case &HC6% To &HC7%
                  Operand = GetByte(Code, Position)
                  Instruction = OPCODES_C6C7((Operand And &H3F%) >> &H3%)

                  If Not Instruction = Nothing Then
                     If Opcode = &HC6% Then
                        Instruction &= $"{If(Operand < &HC0%, " BYTE ", Nothing)}{GetOperand(Code, Position, Operand, LH_REGISTERS)}, {BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}"
                     ElseIf Opcode = &HC7% Then
                        Instruction &= $"{If(Operand < &HC0%, " WORD ", Nothing)}{GetOperand(Code, Position, Operand, XP_REGISTERS)}, {BytesToHexadecimal(GetBytes(Code, Position, Length:=2))}"
                     End If
                  End If
               Case &HCA% : Instruction = $"RETF {BytesToHexadecimal(GetBytes(Code, Position, Length:=2))}"
               Case &HCC% : Instruction = $"INT {HEXADECIMAL_PREFIX}03"
               Case &HCD% : Instruction = $"INT {BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}"
               Case &HC0% To &HC1%, &HD2% To &HD3%
                  Operand = GetByte(Code, Position)
                  Instruction = OPCODES_C0C1_D2D3((Operand And &H3F%) >> &H3%)

                  If Not Instruction = Nothing Then
                     If (Opcode And &H1%) = &H0% Then Instruction &= $" BYTE {GetOperand(Code, Position, Operand, LH_REGISTERS)}"
                     If (Opcode And &H1%) = &H1% Then Instruction &= $" WORD {GetOperand(Code, Position, Operand, XP_REGISTERS)}"

                     If Opcode = &HC0% OrElse Opcode = &HC1% Then Instruction &= $", {HEXADECIMAL_PREFIX}{GetByte(Code, Position):X2}"
                     If Opcode = &HD2% OrElse Opcode = &HD3% Then Instruction &= $", {LH_REGISTERS(COUNTER_REGISTER)}"
                  End If
               Case &HD4% : Instruction = $"AAM {BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}"
               Case &HD5% : Instruction = $"AAD {BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}"
               Case &HE0% : Instruction = $"LOOPNE {ShortAddressToHexadecimal(GetBytes(Code, Position, Length:=1), Position)}"
               Case &HE1% : Instruction = $"LOOPE {ShortAddressToHexadecimal(GetBytes(Code, Position, Length:=1), Position)}"
               Case &HE2% : Instruction = $"LOOP {ShortAddressToHexadecimal(GetBytes(Code, Position, Length:=1), Position)}"
               Case &HE3% : Instruction = $"JCXZ {ShortAddressToHexadecimal(GetBytes(Code, Position, Length:=1), Position)}"
               Case &HE4% : Instruction = $"IN {LH_REGISTERS(ACCUMULATOR_REGISTER)}, {BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}"
               Case &HE5% : Instruction = $"IN {XP_REGISTERS(ACCUMULATOR_REGISTER)}, {BytesToHexadecimal(GetBytes(Code, Position, Length:=2))}"
               Case &HE6% : Instruction = $"OUT {BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}, {LH_REGISTERS(ACCUMULATOR_REGISTER)}"
               Case &HE7% : Instruction = $"OUT {BytesToHexadecimal(GetBytes(Code, Position, Length:=2))}, {XP_REGISTERS(ACCUMULATOR_REGISTER)}"
               Case &HE8% : Instruction = $"CALL NEAR {NearAddressToHexadecimal(GetBytes(Code, Position, Length:=2), Position)}"
               Case &HE9% : Instruction = $"JMP NEAR {NearAddressToHexadecimal(GetBytes(Code, Position, Length:=2), Position)}"
               Case &HEA% : Instruction = $"JMP FAR {FarAddressToHexadecimal(GetBytes(Code, Position, Length:=4))}"
               Case &HEB% : Instruction = $"JMP SHORT {ShortAddressToHexadecimal(GetBytes(Code, Position, Length:=1), Position)}"
               Case &HEC% : Instruction = $"IN {LH_REGISTERS(ACCUMULATOR_REGISTER)}, {XP_REGISTERS(DATA_REGISTER)}"
               Case &HED% : Instruction = $"IN {XP_REGISTERS(ACCUMULATOR_REGISTER)}, {XP_REGISTERS(DATA_REGISTER)}"
               Case &HEE% : Instruction = $"OUT {XP_REGISTERS(DATA_REGISTER)}, {LH_REGISTERS(ACCUMULATOR_REGISTER)}"
               Case &HEF% : Instruction = $"OUT {XP_REGISTERS(DATA_REGISTER)}, {XP_REGISTERS(ACCUMULATOR_REGISTER)}"
               Case &HF6% To &HF7%
                  Operand = GetByte(Code, Position)
                  Instruction = OPCODES_F6F7((Operand And &H3F%) >> &H3%)

                  If Not Instruction = Nothing Then
                     If Opcode = &HF6% Then
                        Instruction &= $" BYTE {GetOperand(Code, Position, Operand, LH_REGISTERS)}"
                        If (Operand And &H3F%) <= &H7% Then Instruction &= $", {BytesToHexadecimal(GetBytes(Code, Position, Length:=1))}"
                     ElseIf Opcode = &HF7% Then
                        Instruction &= $" WORD {GetOperand(Code, Position, Operand, XP_REGISTERS)}"
                        If (Operand And &H3F%) <= &H7% Then Instruction &= $", {BytesToHexadecimal(GetBytes(Code, Position, Length:=2))}"
                     End If
                  End If
               Case &HFE%
                  Operand = GetByte(Code, Position)
                  If Operand < &HC0% Then Instruction = OPCODES_FE__00BF((Operand And &H3F%) >> &H3%)
                  If Operand > &HBF% Then Instruction = OPCODES_FE__C0FF((Operand And &H3F%) >> &H3%)
                  If Not Instruction = Nothing Then Instruction &= $" {GetOperand(Code, Position, Operand, LH_REGISTERS)}"
            End Select

            If Position > Code.Count Then
               Instruction = Nothing
               Position = PreviousPosition
            End If

            If Instruction = Nothing Then
               Instruction = $"DB {BytesToHexadecimal({Opcode}.ToList())}"
               Position = PreviousPosition
            End If

            Return Instruction
         End If

         Return ""
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified far address as a hexadecimal segment and offset word:word (0x0000:0000) representation.
   Private Function FarAddressToHexadecimal(FarAddress As List(Of Byte)) As String
      Try
         Return If(FarAddress.Count < &H4%, Nothing, $"{BytesToHexadecimal({FarAddress(&H2%), FarAddress(&H3%)}.ToList())}:{BytesToHexadecimal({FarAddress(&H0%), FarAddress(&H1%)}.ToList(), NoPrefix:=True)}")
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns a byte from the specified code at the position specified.
   Private Function GetByte(Code As List(Of Byte), ByRef Position As Integer) As Byte
      Try
         Position += &H1%
         Return If(Position > Code.Count, Nothing, Code(Position - &H1%))
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns a string of bytes from the specified code at the position specified.
   Public Function GetBytes(Code As List(Of Byte), ByRef Position As Integer, Length As Integer) As List(Of Byte)
      Try
         Dim Bytes As New List(Of Byte)

         If Position + Length <= Code.Count Then Bytes = Code.GetRange(Position, Length)

         Position += Length

         Return Bytes
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified operand as a string literal.
   Private Function GetOperand(Code As List(Of Byte), ByRef Position As Integer, Operand As Integer, Optional Registers As List(Of String) = Nothing) As String
      Try
         Dim Literal As String = Nothing

         Select Case Operand
            Case Is < &H40%
               Literal = MEMORY_OPERANDS(Operand And &H7%)
               If Literal = MEMORY_OPERANDS(BASE_POINTER) Then Literal = BytesToHexadecimal(GetBytes(Code, Position, Length:=2))
            Case Is < &H80%
               Literal = $"{MEMORY_OPERANDS(Operand And &H7%)} {BytesToHexadecimal(GetBytes(Code, Position, Length:=1), , Signed:=True)}"
            Case Is < &HC0%
               Literal = $"{MEMORY_OPERANDS(Operand And &H7%)} + {BytesToHexadecimal(GetBytes(Code, Position, Length:=2), , Signed:=True)}"
            Case Else
               Literal = Registers(Operand And &H7%)
         End Select

         If Not Literal = Nothing AndAlso Operand < &HC0% Then Literal = $"[{Literal}]"

         Return Literal
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified relative near address as a hexadecimal absolute word (0x0000) representation.
   Private Function NearAddressToHexadecimal(NearAddress As List(Of Byte), Position As Integer) As String
      Try
         Return If(NearAddress.Count < &H2%, Nothing, $"{((Position + ToInt32(BytesToHexadecimal(NearAddress, NoPrefix:=True), fromBase:=16) And &HFFFF%)):X8}")
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified unsigned hexadecimal number as a signed number.
   Public Function SignHexadecimal(Unsigned As String) As String
      Try
         Dim HasPrefix As Boolean = (Unsigned.Substring(0, HEXADECIMAL_PREFIX.Length) = HEXADECIMAL_PREFIX)
         Dim Mask As String = Nothing
         Dim Sign As String = Nothing
         Dim Signed As String = Nothing

         If Not Unsigned.EndsWith("?"c) Then
            If HasPrefix Then
               If Unsigned.Length < &H4% Then Unsigned &= New String("0"c, &H4% - Unsigned.Length)
               Unsigned = Unsigned.Substring(2)
            Else
               If Unsigned.Length < &H2% Then Unsigned &= New String("0"c, &H2% - Unsigned.Length)
            End If

            Mask = $"7F{New String("F"c, Unsigned.Length - &H2%)}"
            If ToInt32(Unsigned, fromBase:=16) > ToInt32(Mask, fromBase:=16) Then
               Sign = "-"c
               Signed = $"{((ToInt32(Mask, fromBase:=16) + &H2%) - (ToInt32(Unsigned, fromBase:=16) - ToInt32(Mask, fromBase:=16))):X}"
            Else
               Sign = "+"c
               Signed = Unsigned
            End If

            If Signed.Length < Unsigned.Length Then Signed = $"{New String("0"c, Unsigned.Length - Signed.Length)}{Signed}"
            If HasPrefix Then Signed = $"{HEXADECIMAL_PREFIX}{Signed}"
         End If

         Return $"{Sign}{Signed}"
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified relative short address as a hexadecimal absolute byte (0x00) representation.
   Private Function ShortAddressToHexadecimal(ShortAddress As List(Of Byte), Position As Integer) As String
      Try
         Return If(ShortAddress.Count < &H1%, Nothing, $"{((Position + ToSByte(BytesToHexadecimal(ShortAddress, NoPrefix:=True), fromBase:=16) And &HFFFF%)):X4}")
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Class
