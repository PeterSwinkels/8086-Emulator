'This class' imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Globalization
Imports System.Linq

'This class contains the 8086/80186 assembler.
Public Class AssemblerClass
   Private Const ACCUMULATOR As Integer = &H0%                        'Defines the accumulator register.
   Private Const ACCUMULATOR_TARGET As Integer = &H4%                 'Indicates that the accumulator register is the target with an immediate numeric value as the source.
   Private Const BASE_POINTER As Integer = &H6%                       'Defines the base pointer register.
   Private Const COUNTER_REGISTER As Integer = &H1%                   'Defines the counter register
   Private Const DEFINE_BYTE As String = "DB"                         'Defines the "define byte" macro instruction.
   Private Const DEFINE_WORD As String = "DW"                         'Defines the "define word" macro instruction.
   Private Const MEMORY_OPERAND_END As Char = "]"c                    'Defines a memory operand's first character.
   Private Const MEMORY_OPERAND_START As Char = "["c                  'Defines a memory operand's last character.
   Private Const MOVE_IMMEDIATE_8 As Integer = &HB0%                  'Defines the base opcode for the MOV instruction with an 8 bit immediate numeric values as a source.
   Private Const MOVE_IMMEDIATE_16 As Integer = &HB8%                 'Defines the base opcode for the MOV instruction with an 16 bit immediate numeric values as a source.
   Private Const MOVE_MEMORY_IMMEDIATE As Integer = &HA0%             'Defines the base opcode for the MOV instruction with a memory immediate/acculumator source/target.
   Private Const MOVE_SEGMENT_SOURCE As Integer = &H8C%               'Defines the base opcode for the MOV instruction with a segment register as a source.
   Private Const MOVE_SEGMENT_TARGET As Integer = &H8E%               'Defines the base opcode for the MOV instruction with a segment register as a target.
   Private Const OPERAND_8_BIT As Integer = &H0%                      'Indicates 8 bit operands are used.
   Private Const OPERAND_16_BIT As Integer = &H1%                     'Indicates 16 bit operands are used.
   Private Const REGISTER_SOURCE As Integer = &H0%                    'Indicates that a general purpose register is the source.
   Private Const REGISTER_TARGET As Integer = &H2%                    'Indicates that a general purpose register is the target.
   Private Const ROTATE_SHIFT_BYTE_BYTE As Byte = &HC0%               'Defines rotate/shift opcodes that take a byte operand and byte target.
   Private Const ROTATE_SHIFT_WORD_BYTE As Byte = &HC1%               'Defines rotate/shift opcodes that take a byte operand and word target.
   Private Const ROTATE_SHIFT_BYTE_CL As Byte = &HD2%                 'Defines rotate/shift opcodes that take CL as an operand and byte target.
   Private Const ROTATE_SHIFT_WORD_CL As Byte = &HD3%                 'Defines rotate/shift opcodes that take CL as an operand and word target.
   Private Const VARIOUS_BYTE As Byte = &H80%                         'Defines the opcode for instructions with a byte target and a byte source.
   Private Const VARIOUS_BYTE_LITERAL As Byte = &HF6%                 'Defines the opcode for instructions with a byte target and a byte source immediate.
   Private Const VARIOUS_REGISTER_8_BIT As Byte = &HFE%               'Defines the opcode for various 8 bit register instructions.
   Private Const VARIOUS_REGISTER_16_BIT As Byte = &HFF%              'Defines the opcode for various 16 bit register instructions.
   Private Const VARIOUS_WORD As Byte = &H81%                         'Defines the opcode for instructions with a word target and a word source.
   Private Const VARIOUS_WORD_LITERAL As Byte = &HF7%                 'Defines the opcode for instructions with a word target and a word source immediate.
   Private Const VARIOUS_WORD_BYTE As Byte = &H83%                    'Defines the opcode for instructions with a word target and a byte source.
   Private Const WORD_BYTE_OPERAND_SOURCE_PREFIX As String = "BYTE"   'Defines the source operand prefix for word/byte instructions.

   'This enumeration lists the supported operand opcodes.
   Private Enum OperandOpcodesE As Integer
      Memory = &H0%                   'A memory operand opcode.
      Memory8BitOImmediate = &H40%    'A memory operand with an 8 bit immediate opcode.
      Memory16BitOImmediate = &H80%   'A memory operand with an 16 bit immediate opcode.
      Registers = &HC0%               'Both operands are registers.
   End Enum

   'This enumeration lists the supported operand types.
   Private Enum OperandTypesE As Integer
      Unknown         'Unknown.
      Immediate       'Numeric immediate value.
      Memory          'Memory location.
      Register8Bit    '8 bit register.
      Register16Bit   '16 bit register.
      Segment         'Segment register.
   End Enum

   'This structure defines an operand and its properties.
   Private Structure OperandStr
      Public Operand As String       'Defines the operand's code.
      Public Type As OperandTypesE   'Defines the operand's type.
   End Structure

   'This dictionary contains the instructions that take the AL register as a source.
   Private ReadOnly AL_SOURCE_OPCODES As New Dictionary(Of String, Byte) From
       {{"OUT", &HE6%}}

   'This dictionary contains the instructions that take the AL register as a source.
   Private ReadOnly AX_SOURCE_OPCODES As New Dictionary(Of String, Byte) From
       {{"OUT", &HE7%}}

   'This dictionary contains the instructions that take a byte operand.
   Private ReadOnly BYTE_OPERAND_OPCODES As New Dictionary(Of String, Byte) From
       {{"AAD", &HD5%},
        {"AAM", &HD4%},
        {"ADD AL", &H4%},
        {"IN AL", &HE4%},
        {"INT", &HCD%},
        {"RETN", &HC2%},
        {"TEST AL", &HA8%}}

   'This dictionary contains the jump instructions that take a two word operands (segment:offset).
   Private ReadOnly FAR_JUMP_OPCODES As New Dictionary(Of String, Byte) From
       {{"CALL FAR", &H9A%},
        {"JMP FAR", &HEA%}}

   'This dictionary contains the instructions that load memory addresses into registers.
   Private ReadOnly LOAD_ADDRESS_OPCODES As New Dictionary(Of String, Byte) From
       {{"LDS", &HC5%},
        {"LEA", &H8D%},
        {"LES", &HC4%}}

   'This dictionary contains the MOV BYTE instruction
   Private ReadOnly MOV_BYTE_OPCODE As New Dictionary(Of String, Byte) From
       {{"MOV BYTE", &HC6%}}

   'This dictionary contains the MOV instruction.
   Private ReadOnly MOV_OPCODE As New Dictionary(Of String, Byte) From
       {{"MOV", &H88%}}

   'This dictionary contains the MOV WORD instruction
   Private ReadOnly MOV_WORD_OPCODE As New Dictionary(Of String, Byte) From
       {{"MOV WORD", &HC7%}}

   'This dictionary contains the jump instructions that take a word operand.
   Private ReadOnly NEAR_JUMP_OPCODES As New Dictionary(Of String, Byte) From
       {{"CALL NEAR", &HE8%},
        {"JMP NEAR", &HE9%}}

   'This dictionary contains the instructions that take no operands.
   Private ReadOnly NO_OPERAND_OPCODES As New Dictionary(Of String, Byte) From
       {{"AAA", &H37%},
        {"AAS", &H3F%},
        {"CBW", &H98%},
        {"CLC", &HF8%},
        {"CLD", &HFC%},
        {"CLI", &HFA%},
        {"CMC", &HF5%},
        {"CMPSB", &HA6%},
        {"CMPSW", &HA7%},
        {"CWD", &H99%},
        {"DAA", &H27%},
        {"DAS", &H2F%},
        {"HLT", &HF4%},
        {"IN AL,DX", &HEC%},
        {"IN AX,DX", &HED%},
        {"INT3", &HCC%},
        {"INTO", &HCE%},
        {"IRET", &HCF%},
        {"LAHF", &H9F%},
        {"LOCK", &HF0%},
        {"LODSB", &HAC%},
        {"LODSW", &HAD%},
        {"MOVSB", &HA4%},
        {"MOVSW", &HA5%},
        {"NOP", &H90%},
        {"OUT DX,AL", &HEE%},
        {"OUT DX,AX", &HEF%},
        {"POPFW", &H9D%},
        {"PUSHFW", &H9C%},
        {"REP", &HF3%},
        {"REPE", &HF3%},
        {"REPNE", &HF2%},
        {"REPNZ", &HF2%},
        {"REPZ", &HF3%},
        {"RETF", &HCB%},
        {"RETN", &HC3%},
        {"SAHF", &H9E%},
        {"SALC", &HD6%},
        {"SCASB", &HAE%},
        {"SCASW", &HAF%},
        {"STC", &HF9%},
        {"STD", &HFD%},
        {"STI", &HFB%},
        {"STOSB", &HAA%},
        {"STOSW", &HAB%},
        {"WAIT", &H9B%},
        {"XLATB", &HD7%}}

   'This dictionary contains the POP memory instruction.
   Private ReadOnly POP_MEMORY_OPCODE As New Dictionary(Of String, Byte) From
       {{"POP", &H8F%}}

   'This dictionary contains the instructions that take a segment register operand.
   Private ReadOnly SEGMENT_OPERAND_OPCODES As New Dictionary(Of String, Byte) From
       {{"POP", &H7%},
        {"PUSH", &H6%}}

   'This dictionary contains the segment prefix instructions.
   Private ReadOnly SEGMENT_PREFIX_OPCODES As New Dictionary(Of String, Byte) From
       {{"CS", &H2E%},
        {"DS", &H3E%},
        {"ES", &H26%},
        {"SS", &H36%}}

   'This dictionary contains the jump instructions that take a byte operand.
   Private ReadOnly SHORT_JUMP_OPCODES As New Dictionary(Of String, Byte) From
       {{"JA", &H77%},
        {"JAE", &H73%},
        {"JBE", &H76%},
        {"JC", &H72%},
        {"JCXZ", &HE3%},
        {"JE", &H74%},
        {"JG", &H7F%},
        {"JGE", &H7D%},
        {"JL", &H7C%},
        {"JLE", &H7E%},
        {"JMP SHORT", &HEB%},
        {"JNA", &H76%},
        {"JNB", &H73%},
        {"JNBE", &H77%},
        {"JNC", &H73%},
        {"JNE", &H75%},
        {"JNG", &H7E%},
        {"JNGE", &H7C%},
        {"JNL", &H7D%},
        {"JNLE", &H7F%},
        {"JNO", &H71%},
        {"JNP", &H7B%},
        {"JNS", &H79%},
        {"JNZ", &H75%},
        {"JO", &H70%},
        {"JP", &H7A%},
        {"JPE", &H7A%},
        {"JPO", &H7B%},
        {"JS", &H78%},
        {"JZ", &H74%},
        {"LOOP", &HE2%},
        {"LOOPE", &HE1%},
        {"LOOPNE", &HE0%},
        {"LOOPNZ", &HE0%},
        {"LOOPZ", &HE1%}}

   'This dictionary contains the instructions that take a source and target operand.
   Private ReadOnly TARGET_SOURCE_OPCODES As New Dictionary(Of String, Byte) From
       {{"TEST", &H84%},
        {"XCHG", &H86%}}

   'This dictionary contains the instructions that take a source and target operand of various types.
   Private ReadOnly VARIOUS_TARGET_SOURCE_OPCODES As New Dictionary(Of String, Byte) From
       {{"ADC", &H10%},
        {"ADD", &H0%},
        {"AND", &H20%},
        {"CMP", &H38%},
        {"OR", &H8%},
        {"SBB", &H18%},
        {"SUB", &H28%},
        {"XOR", &H30%}}

   'This dictionary contains the instructions that take a word operand.
   Private ReadOnly WORD_OPERAND_OPCODES As New Dictionary(Of String, Byte) From
       {{"ADD AX", &H5%},
        {"IN AX", &HE5%},
        {"RETF", &HCA%},
        {"TEST AX", &HA9%}}

   'This dictionary contains the XCHG instruction.
   Private ReadOnly XCHG_OPCODE As New Dictionary(Of String, Byte) From
       {{"XCHG", &H90%}}

   'This dictionary contains the instructions that take a 16 bit register operand.
   Private ReadOnly XP_OPERAND_OPCODES As New Dictionary(Of String, Byte) From
       {{"DEC", &H48%},
        {"INC", &H40%},
        {"POP", &H58%},
        {"PUSH", &H50%}}

   Private ReadOnly IS_16_BIT_REGISTER_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (XP_REGISTERS.Contains(Operand.Trim().ToUpper()))                                                        'Indicates whether the specified operand is an 16 bit register.
   Private ReadOnly IS_8_BIT_IMMEDIATE As Func(Of String, Boolean) = Function(Immediate As String) (ToInt32(Immediate, fromBase:=16) < &H100%)                                                                    'Indicates whether the specified immediate is an 8 bit value.
   Private ReadOnly IS_8_BIT_REGISTER_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (LH_REGISTERS.Contains(Operand.Trim().ToUpper()))                                                         'Indicates whether the specified operand is an 8 bit register.
   Private ReadOnly IS_IMMEDIATE_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (Integer.TryParse(Operand, NumberStyles.HexNumber, Nothing, Nothing))                                          'Indicates whether the specified operand is an immediate numeric value.
   Private ReadOnly IS_MEMORY_IMMEDIATE As Func(Of String, Boolean) = Function(Operand As String) (IS_MEMORY_OPERAND(Operand) AndAlso IS_IMMEDIATE_OPERAND(Operand.Substring(1, Operand.Length - 2)))             'Indicates whether the specified operand is a memory immediate numeric value.
   Private ReadOnly IS_MEMORY_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (Operand.Trim().StartsWith(MEMORY_OPERAND_START) AndAlso Operand.Trim().EndsWith(MEMORY_OPERAND_END))             'Indicates whether the specified operand refers to a memory location.
   Private ReadOnly IS_REGISTER_OPERAND As Func(Of OperandTypesE, Boolean) = Function(OperandType As OperandTypesE) (OperandType = OperandTypesE.Register16Bit OrElse OperandType = OperandTypesE.Register8Bit)   'Indicates whether the specified operand type refers to a register.
   Private ReadOnly IS_SEGMENT_OPERAND As Func(Of String, Boolean) = Function(Operand As String) (SG_REGISTERS.Contains(Operand.Trim().ToUpper()))                                                                'Indicates whether the specified operand is a segment register.
   Private ReadOnly LH_REGISTERS As New List(Of String)({"AL", "CL", "DL", "BL", "AH", "CH", "DH", "BH"})                                                                                                         'Contains the 8086's 8 bit general purpose registers.
   Private ReadOnly MEMORY_IMMEDIATE_OPERANDS As New List(Of String)({"BX+SI+", "BX+DI+", "BP+SI+", "BP+DI+", "SI+", "DI+", "BP+", "BX+"})                                                                        'Contains the 8086's memory oparands that include an immediate numeric value.
   Private ReadOnly MEMORY_OPERANDS As New List(Of String)({"BX+SI", "BX+DI", "BP+SI", "BP+DI", "SI", "DI", "", "BX"})                                                                                            'Contains the 8086's memory oparands.
   Private ReadOnly OPCODES_C0D2 As New List(Of String)({"ROL BYTE", "ROR BYTE", "RCL BYTE", "RCR BYTE", "SHL BYTE", "SHR BYTE", Nothing, "SAR BYTE"})                                                            'Contains the rotate/shift byte instructions.
   Private ReadOnly OPCODES_C1D3 As New List(Of String)({"ROL WORD", "ROR WORD", "RCL WORD", "RCR WORD", "SHL WORD", "SHR WORD", Nothing, "SAR WORD"})                                                            'Contains the rotate/shift word instructions.
   Private ReadOnly OPCODES_80 As New List(Of String)({"ADD BYTE", "OR BYTE", "ADC BYTE", "SBB BYTE", "AND BYTE", "SUB BYTE", "XOR BYTE", "CMP BYTE"})                                                            'Contains various byte instructions.
   Private ReadOnly OPCODES_8183 As New List(Of String)({"ADD WORD", "OR WORD", "ADC WORD", "SBB WORD", "AND WORD", "SUB WORD", "XOR WORD", "CMP WORD"})                                                          'Contains various word instructions.
   Private ReadOnly OPCODES_F6 As New List(Of String)({Nothing, Nothing, "NOT BYTE", "NEG BYTE", Nothing, Nothing, Nothing, Nothing})                                                                             'Contains various byte instructions.
   Private ReadOnly OPCODES_F6_BYTE As New List(Of String)({"TEST BYTE", Nothing, Nothing, Nothing, "MUL BYTE", "IMUL BYTE", "DIV BYTE", "IDIV BYTE"})                                                            'Contains various byte instructions.
   Private ReadOnly OPCODES_F7 As New List(Of String)({Nothing, Nothing, "NOT WORD", "NEG WORD", Nothing, Nothing, Nothing, Nothing})                                                                             'Contains various word instructions.
   Private ReadOnly OPCODES_F7_WORD As New List(Of String)({"TEST WORD", Nothing, Nothing, Nothing, "MUL WORD", "IMUL WORD", "DIV WORD", "IDIV WORD"})                                                            'Contains various word instructions.
   Private ReadOnly OPCODES_FE__00BF As New List(Of String)({"INC BYTE", "DEC BYTE", "CALL WORD NEAR", "CALL WORD FAR", "JMP WORD NEAR", "JMP WORD FAR", "PUSH WORD", Nothing})                                   'Contains various instructions.
   Private ReadOnly OPCODES_FE__C0FF As New List(Of String)({"INC BYTE", "DEC BYTE", "CALL", Nothing, "JMP", Nothing, "PUSH", Nothing})                                                                           'Contains various instructions.
   Private ReadOnly OPCODES_FF__00BF As New List(Of String)({"INC WORD", "DEC WORD", "CALL WORD NEAR", "CALL WORD FAR", "JMP WORD NEAR", "JMP WORD FAR", "PUSH WORD", Nothing})                                   'Contains various instructions.
   Private ReadOnly OPCODES_FF__C0FF As New List(Of String)({Nothing, Nothing, "CALL", Nothing, "JMP", Nothing, Nothing, Nothing})                                                                                'Contains various instructions.
   Private ReadOnly SG_REGISTERS As New List(Of String)({"ES", "CS", "SS", "DS"})                                                                                                                                 'Contains the 8086's segment registers.
   Private ReadOnly XP_REGISTERS As New List(Of String)({"AX", "CX", "DX", "BX", "SP", "BP", "SI", "DI"})                                                                                                         'Contains the 8086's 16 bit general purpose registers.

   Private ReadOnly ILLEGAL_OPCODES As New List(Of Byte)({&HF%})   'This list defines opcodes resulting from an invalid combination of operands and instructions.

   'The definitions for events that can be raised by this class.
   Public Event HandleError(ExceptionO As Exception)

   'This procedure accepts the specified assembly language instruction and returns the appropriate opcodes.
   Public Function Assemble(Instruction As String) As List(Of Byte)
      Try
         Dim ExtraOperand As OperandStr = Nothing
         Dim ImmediateBytes As New List(Of Byte)
         Dim LeftOperand As OperandStr = Nothing
         Dim Offset As New Integer
         Dim Opcode As New Integer
         Dim Opcodes As New List(Of Byte)
         Dim RegisterIndex As New Integer
         Dim RightOperand As OperandStr = Nothing
         Dim Segment As New Integer

         Instruction = RemoveWhiteSpace(Instruction).ToUpper()

         Select Case True
            Case NO_OPERAND_OPCODES.ContainsKey(Instruction)
               Opcodes.Add(NO_OPERAND_OPCODES(Instruction))
            Case Else
               If Instruction.Count(Function(Character As Char) Character = ","c) = 2 Then
                  ExtraOperand = New OperandStr With {.Operand = GetRightMostOperand(Instruction, Delimiter:=","c), .Type = OperandType(.Operand)}
               End If

               If Instruction.Contains(","c) Then
                  RightOperand = New OperandStr With {.Operand = GetRightMostOperand(Instruction, Delimiter:=","c), .Type = OperandType(.Operand)}
               ElseIf Instruction.Contains(" "c) Then
                  RightOperand = New OperandStr With {.Operand = GetRightMostOperand(Instruction, Delimiter:=" "c), .Type = OperandType(.Operand)}
               End If

               Select Case True
                  Case Instruction = DEFINE_BYTE
                     Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=True))
                  Case Instruction = DEFINE_WORD
                     Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=False))
                  Case FAR_JUMP_OPCODES.ContainsKey(Instruction)
                     Opcodes.Add(FAR_JUMP_OPCODES(Instruction))
                     Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand.Split(":"c).Last(), Is8Bit:=False))
                     Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand.Split(":"c).First(), Is8Bit:=False))
                  Case NEAR_JUMP_OPCODES.ContainsKey(Instruction)
                     Opcodes.Add(NEAR_JUMP_OPCODES(Instruction))
                     Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=False))
                  Case OPCODES_F6.Contains(Instruction)
                     Opcodes.Add(VARIOUS_BYTE_LITERAL)

                     If RightOperand.Type = OperandTypesE.Memory Then
                        Opcodes.Add(ToByte(ParseMemoryOperand(RightOperand.Operand, ImmediateBytes) Or (OPCODES_F6.IndexOf(Instruction) << &H3%)))
                        If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                     ElseIf RightOperand.Type = OperandTypesE.Register8Bit Then
                        Opcodes.Add(ToByte((LH_REGISTERS.IndexOf(RightOperand.Operand) Or OperandOpcodesE.Registers) Or (OPCODES_F6.IndexOf(Instruction) << &H3%)))
                     End If
                  Case OPCODES_F7.Contains(Instruction)
                     Opcodes.Add(VARIOUS_WORD_LITERAL)

                     If RightOperand.Type = OperandTypesE.Memory Then
                        Opcodes.Add(ToByte(ParseMemoryOperand(RightOperand.Operand, ImmediateBytes) Or (OPCODES_F7.IndexOf(Instruction) << &H3%)))
                        If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                     ElseIf RightOperand.Type = OperandTypesE.Register16Bit Then
                        Opcodes.Add(ToByte((XP_REGISTERS.IndexOf(RightOperand.Operand) Or OperandOpcodesE.Registers) Or (OPCODES_F7.IndexOf(Instruction) << &H3%)))
                     End If
                  Case SEGMENT_PREFIX_OPCODES.ContainsKey(Instruction)
                     Opcodes.Add(SEGMENT_PREFIX_OPCODES(Instruction))
                  Case SHORT_JUMP_OPCODES.ContainsKey(Instruction)
                     Opcodes.Add(SHORT_JUMP_OPCODES(Instruction))
                     Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=True))
                  Case BYTE_OPERAND_OPCODES.ContainsKey(Instruction) AndAlso RightOperand.Type = OperandTypesE.Immediate
                     Opcodes.Add(BYTE_OPERAND_OPCODES(Instruction))
                     Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=True))
                  Case OPCODES_FE__00BF.Contains(Instruction) AndAlso RightOperand.Type = OperandTypesE.Memory
                     Opcodes.Add(VARIOUS_REGISTER_8_BIT)
                     Opcodes.Add(ToByte(ParseMemoryOperand(RightOperand.Operand, ImmediateBytes) Or (OPCODES_FE__00BF.IndexOf(Instruction) << &H3%)))
                     If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                  Case OPCODES_FE__C0FF.Contains(Instruction) AndAlso RightOperand.Type = OperandTypesE.Register8Bit
                     Opcodes.Add(VARIOUS_REGISTER_8_BIT)
                     Opcodes.Add(ToByte(LH_REGISTERS.IndexOf(RightOperand.Operand) Or (OPCODES_FE__C0FF.IndexOf(Instruction) << &H3%) Or OperandOpcodesE.Registers))
                  Case OPCODES_FF__00BF.Contains(Instruction) AndAlso RightOperand.Type = OperandTypesE.Memory
                     Opcodes.Add(VARIOUS_REGISTER_16_BIT)
                     Opcodes.Add(ToByte(ParseMemoryOperand(RightOperand.Operand, ImmediateBytes) Or (OPCODES_FF__00BF.IndexOf(Instruction) << &H3%)))
                     If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                  Case OPCODES_FF__C0FF.Contains(Instruction) AndAlso RightOperand.Type = OperandTypesE.Register16Bit
                     Opcodes.Add(VARIOUS_REGISTER_16_BIT)
                     Opcodes.Add(ToByte(XP_REGISTERS.IndexOf(RightOperand.Operand) Or (OPCODES_FF__C0FF.IndexOf(Instruction) << &H3%) Or OperandOpcodesE.Registers))
                  Case POP_MEMORY_OPCODE.ContainsKey(Instruction) AndAlso RightOperand.Type = OperandTypesE.Memory
                     Opcodes.Add(POP_MEMORY_OPCODE(Instruction))
                     Opcodes.Add(ToByte(ParseMemoryOperand(RightOperand.Operand, ImmediateBytes)))
                     Opcodes.AddRange(ImmediateBytes)
                  Case WORD_OPERAND_OPCODES.ContainsKey(Instruction) AndAlso RightOperand.Type = OperandTypesE.Immediate
                     Opcodes.Add(WORD_OPERAND_OPCODES(Instruction))
                     Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=False))
                  Case SEGMENT_OPERAND_OPCODES.ContainsKey(Instruction) AndAlso SG_REGISTERS.Contains(RightOperand.Operand)
                     RegisterIndex = SG_REGISTERS.IndexOf(RightOperand.Operand)
                     Opcodes.Add(ToByte(RegisterIndex << &H3%) Or SEGMENT_OPERAND_OPCODES(Instruction))
                  Case XP_OPERAND_OPCODES.ContainsKey(Instruction) AndAlso XP_REGISTERS.Contains(RightOperand.Operand)
                     RegisterIndex = XP_REGISTERS.IndexOf(RightOperand.Operand)
                     If RegisterIndex >= &H0% Then Opcodes.Add(ToByte(XP_OPERAND_OPCODES(Instruction) Or RegisterIndex))
                  Case Else
                     LeftOperand = New OperandStr With {.Operand = GetRightMostOperand(Instruction, Delimiter:=" "c), .Type = OperandType(.Operand)}

                     Select Case True
                        Case RightOperand.Operand = "AL" AndAlso AL_SOURCE_OPCODES.ContainsKey(Instruction)
                           Opcodes.Add(AL_SOURCE_OPCODES(Instruction))
                           Opcodes.AddRange(BytesFromHexadecimal(LeftOperand.Operand, Is8Bit:=True))
                        Case RightOperand.Operand = "AX" AndAlso AX_SOURCE_OPCODES.ContainsKey(Instruction)
                           Opcodes.Add(AX_SOURCE_OPCODES(Instruction))
                           Opcodes.AddRange(BytesFromHexadecimal(LeftOperand.Operand, Is8Bit:=False))
                        Case LOAD_ADDRESS_OPCODES.ContainsKey(Instruction)
                           Opcode = LOAD_ADDRESS_OPCODES(Instruction)

                           If Not (IS_REGISTER_OPERAND(LeftOperand.Type) OrElse RightOperand.Type = OperandTypesE.Register8Bit) Then
                              Opcodes.Add(ToByte(Opcode))
                              Opcodes.AddRange(OperandsToOpcodes(LeftOperand, RightOperand))
                           End If
                        Case MOV_BYTE_OPCODE.ContainsKey(Instruction)
                           Opcodes.Add(MOV_BYTE_OPCODE(Instruction))
                           Opcodes.AddRange(OperandsToOpcodes(LeftOperand, RightOperand, Is8Bit:=True))
                           Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=True))
                        Case MOV_OPCODE.ContainsKey(Instruction)
                           If Not (IS_REGISTER_OPERAND(LeftOperand.Type) AndAlso IS_REGISTER_OPERAND(RightOperand.Type) AndAlso Not LeftOperand.Type = RightOperand.Type) Then
                              If LeftOperand.Type = OperandTypesE.Register8Bit AndAlso RightOperand.Type = OperandTypesE.Immediate Then
                                 Opcode = MOVE_IMMEDIATE_8 Or LH_REGISTERS.IndexOf(LeftOperand.Operand)
                                 Opcodes.Add(ToByte(Opcode))
                                 Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=True))
                              ElseIf LeftOperand.Type = OperandTypesE.Register16Bit AndAlso RightOperand.Type = OperandTypesE.Immediate Then
                                 Opcode = MOVE_IMMEDIATE_16 Or XP_REGISTERS.IndexOf(LeftOperand.Operand)
                                 Opcodes.Add(ToByte(Opcode))
                                 Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=False))
                              Else
                                 If LeftOperand.Type = OperandTypesE.Segment Then
                                    Opcodes.Add(ToByte(MOVE_SEGMENT_TARGET))
                                    Opcodes.AddRange(OperandsToOpcodes(LeftOperand, RightOperand))
                                 ElseIf RightOperand.Type = OperandTypesE.Segment Then
                                    Opcodes.Add(ToByte(MOVE_SEGMENT_SOURCE))
                                    Opcodes.AddRange(OperandsToOpcodes(LeftOperand, RightOperand))
                                 ElseIf (LeftOperand.Operand = LH_REGISTERS(ACCUMULATOR) OrElse LeftOperand.Operand = XP_REGISTERS(ACCUMULATOR)) AndAlso IS_MEMORY_IMMEDIATE(RightOperand.Operand) Then
                                    Opcodes.Add(ToByte(SetOpcodeProperties(MOVE_MEMORY_IMMEDIATE, LeftOperand, RightOperand, SetSourceTargetFlag:=True, FlipSourceTargetFlag:=True)))
                                    Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand.Substring(1, RightOperand.Operand.Length - 2), Is8Bit:=False))
                                 ElseIf (RightOperand.Operand = LH_REGISTERS(ACCUMULATOR) OrElse RightOperand.Operand = XP_REGISTERS(ACCUMULATOR)) AndAlso IS_MEMORY_IMMEDIATE(LeftOperand.Operand) Then
                                    Opcodes.Add(ToByte(SetOpcodeProperties(MOVE_MEMORY_IMMEDIATE, LeftOperand, RightOperand, SetSourceTargetFlag:=True, FlipSourceTargetFlag:=True)))
                                    Opcodes.AddRange(BytesFromHexadecimal(LeftOperand.Operand.Substring(1, LeftOperand.Operand.Length - 2), Is8Bit:=False))
                                 Else
                                    Opcodes.Add(ToByte(SetOpcodeProperties(MOV_OPCODE(Instruction), LeftOperand, RightOperand, SetSourceTargetFlag:=True)))
                                    Opcodes.AddRange(OperandsToOpcodes(LeftOperand, RightOperand))
                                 End If
                              End If
                           End If
                        Case MOV_WORD_OPCODE.ContainsKey(Instruction)
                           Opcodes.Add(MOV_WORD_OPCODE(Instruction))
                           Opcodes.AddRange(OperandsToOpcodes(LeftOperand, RightOperand, Is8Bit:=False))
                           Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=False))
                        Case OPCODES_80.Contains(Instruction)
                           Opcodes.Add(VARIOUS_BYTE)
                           Opcodes.Add(ToByte(CInt(ParseMemoryOperand(LeftOperand.Operand, ImmediateBytes)) Or OPCODES_80.IndexOf(Instruction) << &H3%))
                           If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                           Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=True))
                        Case OPCODES_8183.Contains(Instruction)
                           If RightOperand.Operand.StartsWith(WORD_BYTE_OPERAND_SOURCE_PREFIX) Then
                              RightOperand.Operand = RightOperand.Operand.Substring(WORD_BYTE_OPERAND_SOURCE_PREFIX.Length).Trim()
                              Opcodes.Add(VARIOUS_WORD_BYTE)
                           Else
                              Opcodes.Add(VARIOUS_WORD)
                           End If
                           Opcodes.Add(ToByte(CInt(ParseMemoryOperand(LeftOperand.Operand, ImmediateBytes)) Or OPCODES_8183.IndexOf(Instruction) << &H3%))
                           If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                           Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=False))
                        Case OPCODES_C0D2.Contains(Instruction) AndAlso RightOperand.Type = OperandTypesE.Immediate
                           Opcodes.Add(ROTATE_SHIFT_BYTE_BYTE)

                           If LeftOperand.Type = OperandTypesE.Memory Then
                              Opcodes.Add(ToByte(ParseMemoryOperand(LeftOperand.Operand, ImmediateBytes)))
                              If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                           ElseIf LeftOperand.Type = OperandTypesE.Register8Bit Then
                              Opcodes.Add(ToByte(LH_REGISTERS.IndexOf(LeftOperand.Operand) Or OperandOpcodesE.Registers))
                           End If

                           Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=True))
                        Case OPCODES_C1D3.Contains(Instruction) AndAlso RightOperand.Type = OperandTypesE.Immediate
                           Opcodes.Add(ROTATE_SHIFT_WORD_BYTE)

                           If LeftOperand.Type = OperandTypesE.Memory Then
                              Opcodes.Add(ToByte(ParseMemoryOperand(LeftOperand.Operand, ImmediateBytes)))
                              If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                           ElseIf LeftOperand.Type = OperandTypesE.Register16Bit Then
                              Opcodes.Add(ToByte(XP_REGISTERS.IndexOf(LeftOperand.Operand) Or OperandOpcodesE.Registers))
                           End If

                           Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=True))
                        Case OPCODES_F6_BYTE.Contains(Instruction) AndAlso RightOperand.Type = OperandTypesE.Immediate
                           Opcodes.Add(VARIOUS_BYTE_LITERAL)

                           If LeftOperand.Type = OperandTypesE.Memory Then
                              Opcodes.Add(ToByte(ParseMemoryOperand(LeftOperand.Operand, ImmediateBytes) Or (OPCODES_F6_BYTE.IndexOf(Instruction) << &H3%)))
                              If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                           ElseIf LeftOperand.Type = OperandTypesE.Register8Bit Then
                              Opcodes.Add(ToByte((LH_REGISTERS.IndexOf(LeftOperand.Operand) Or OperandOpcodesE.Registers) Or (OPCODES_F6_BYTE.IndexOf(Instruction) << &H3%)))
                           End If

                           Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=True))
                        Case OPCODES_C0D2.Contains(Instruction) AndAlso RightOperand.Operand = LH_REGISTERS(COUNTER_REGISTER)
                           Opcodes.Add(ROTATE_SHIFT_BYTE_CL)

                           If LeftOperand.Type = OperandTypesE.Memory Then
                              Opcodes.Add(ToByte(ParseMemoryOperand(LeftOperand.Operand, ImmediateBytes)))
                              If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                           ElseIf LeftOperand.Type = OperandTypesE.Register8Bit Then
                              Opcodes.Add(ToByte(LH_REGISTERS.IndexOf(LeftOperand.Operand) Or OperandOpcodesE.Registers))
                           End If
                        Case OPCODES_C1D3.Contains(Instruction) AndAlso RightOperand.Operand = LH_REGISTERS(COUNTER_REGISTER)
                           Opcodes.Add(ROTATE_SHIFT_WORD_CL)

                           If LeftOperand.Type = OperandTypesE.Memory Then
                              Opcodes.Add(ToByte(ParseMemoryOperand(LeftOperand.Operand, ImmediateBytes)))
                              If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                           ElseIf LeftOperand.Type = OperandTypesE.Register16Bit Then
                              Opcodes.Add(ToByte(XP_REGISTERS.IndexOf(LeftOperand.Operand) Or OperandOpcodesE.Registers))
                           End If
                        Case OPCODES_F7_WORD.Contains(Instruction) AndAlso RightOperand.Type = OperandTypesE.Immediate
                           Opcodes.Add(VARIOUS_WORD_LITERAL)

                           If LeftOperand.Type = OperandTypesE.Memory Then
                              Opcodes.Add(ToByte(ParseMemoryOperand(LeftOperand.Operand, ImmediateBytes) Or (OPCODES_F7_WORD.IndexOf(Instruction) << &H3%)))
                              If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)
                           ElseIf LeftOperand.Type = OperandTypesE.Register16Bit Then
                              Opcodes.Add(ToByte((XP_REGISTERS.IndexOf(LeftOperand.Operand) Or OperandOpcodesE.Registers) Or (OPCODES_F7_WORD.IndexOf(Instruction) << &H3%)))
                           End If

                           Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=False))
                        Case TARGET_SOURCE_OPCODES.ContainsKey(Instruction)
                           Opcodes.Add(ToByte(SetOpcodeProperties(TARGET_SOURCE_OPCODES(Instruction), LeftOperand, RightOperand)))
                           Opcodes.AddRange(OperandsToOpcodes(LeftOperand, RightOperand))
                        Case VARIOUS_TARGET_SOURCE_OPCODES.ContainsKey(Instruction)
                           If Not (IS_REGISTER_OPERAND(LeftOperand.Type) AndAlso IS_REGISTER_OPERAND(RightOperand.Type) AndAlso Not LeftOperand.Type = RightOperand.Type) Then
                              Opcode = VARIOUS_TARGET_SOURCE_OPCODES(Instruction)

                              If LeftOperand.Operand = LH_REGISTERS(ACCUMULATOR) AndAlso RightOperand.Type = OperandTypesE.Immediate Then
                                 Opcodes.Add(ToByte(SetOpcodeProperties(Opcode, LeftOperand, RightOperand) Or ACCUMULATOR_TARGET))
                                 Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=True))
                              ElseIf LeftOperand.Operand = XP_REGISTERS(ACCUMULATOR) AndAlso RightOperand.Type = OperandTypesE.Immediate Then
                                 Opcodes.Add(ToByte(SetOpcodeProperties(Opcode, LeftOperand, RightOperand) Or ACCUMULATOR_TARGET))
                                 Opcodes.AddRange(BytesFromHexadecimal(RightOperand.Operand, Is8Bit:=False))
                              Else
                                 Opcodes.Add(ToByte(SetOpcodeProperties(Opcode, LeftOperand, RightOperand, SetSourceTargetFlag:=True)))
                                 Opcodes.AddRange(OperandsToOpcodes(LeftOperand, RightOperand))
                              End If
                           End If
                        Case XCHG_OPCODE.ContainsKey(Instruction) AndAlso LeftOperand.Operand = XP_REGISTERS(ACCUMULATOR) AndAlso XP_REGISTERS.Contains(RightOperand.Operand)
                           Opcodes.Add(ToByte(XCHG_OPCODE(Instruction) Or XP_REGISTERS.IndexOf(RightOperand.Operand)))
                     End Select
               End Select
         End Select

         If Opcodes.Count > &H0% AndAlso ILLEGAL_OPCODES.Contains(Opcodes.First()) Then Opcodes.Clear()

         Return Opcodes
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure converts the specified hexadecimal to bytes and returns these.
   Private Function BytesFromHexadecimal(Number As String, Is8Bit As Boolean) As Byte()
      Return If(Is8Bit, {ToByte(Number.Trim(), fromBase:=16)}, BitConverter.GetBytes(ToUInt16(Number.Trim(), fromBase:=16)))
   End Function

   'This procedure removes the right most operand from the specified instruction using the specified delimiter and returns the results.
   Private Function GetRightMostOperand(ByRef Instruction As String, Delimiter As String) As String
      Dim Operand As String = ""

      If Instruction.Contains(Delimiter) Then
         Operand = Instruction.Substring(Instruction.LastIndexOf(Delimiter)).Trim()
         Instruction = Instruction.Substring(0, Instruction.LastIndexOf(Delimiter)).Trim()
         If Not Delimiter = " "c Then Operand = Operand.Substring(1)
      End If

      Return Operand
   End Function

   'This procedure determines the specified operand's type and returns it.
   Private Function OperandType(Operand As String) As OperandTypesE
      Select Case True
         Case IS_16_BIT_REGISTER_OPERAND(Operand)
            Return OperandTypesE.Register16Bit
         Case IS_8_BIT_REGISTER_OPERAND(Operand)
            Return OperandTypesE.Register8Bit
         Case IS_IMMEDIATE_OPERAND(Operand)
            Return OperandTypesE.Immediate
         Case IS_MEMORY_OPERAND(Operand)
            Return OperandTypesE.Memory
         Case IS_SEGMENT_OPERAND(Operand)
            Return OperandTypesE.Segment
      End Select

      Return OperandTypesE.Unknown
   End Function

   'This procedure converts the specified operands to opcodes and returns these.
   Private Function OperandsToOpcodes(LeftOperand As OperandStr, RightOperand As OperandStr, Optional Is8Bit As Boolean? = Nothing) As List(Of Byte)
      Dim ImmediateBytes As New List(Of Byte)
      Dim ImmediateIs8Bit As New Boolean
      Dim LeftOpcode As New Integer
      Dim Opcodes As New List(Of Byte)
      Dim RightOpcode As New Integer

      If RightOperand.Type = OperandTypesE.Immediate AndAlso Is8Bit IsNot Nothing Then
         If LeftOperand.Type = OperandTypesE.Memory Then
            LeftOpcode = CInt(ParseMemoryOperand(LeftOperand.Operand, ImmediateBytes))
            RightOpcode = &H0%
         End If
      Else
         If LeftOperand.Type = OperandTypesE.Memory Then
            LeftOpcode = CInt(ParseMemoryOperand(LeftOperand.Operand, ImmediateBytes))
            RightOpcode = CInt(RegisterOperandToIndex(RightOperand) << &H3%)
         ElseIf RightOperand.Type = OperandTypesE.Memory Then
            RightOpcode = CInt(ParseMemoryOperand(RightOperand.Operand, ImmediateBytes))
            LeftOpcode = CInt(RegisterOperandToIndex(LeftOperand) << &H3%)
         ElseIf LeftOperand.Type = OperandTypesE.Register16Bit AndAlso RightOperand.Type = OperandTypesE.Segment Then
            LeftOpcode = OperandOpcodesE.Registers Or CInt(RegisterOperandToIndex(LeftOperand))
            RightOpcode = CInt(RegisterOperandToIndex(RightOperand) << &H3%)
         ElseIf LeftOperand.Type = OperandTypesE.Segment AndAlso RightOperand.Type = OperandTypesE.Register16Bit Then
            LeftOpcode = CInt(RegisterOperandToIndex(LeftOperand) << &H3%)
            RightOpcode = OperandOpcodesE.Registers Or CInt(RegisterOperandToIndex(RightOperand))
         Else
            LeftOpcode = OperandOpcodesE.Registers Or CInt(RegisterOperandToIndex(LeftOperand))
            RightOpcode = CInt(RegisterOperandToIndex(RightOperand) << &H3%)
         End If
      End If

      Opcodes.Add(ToByte(LeftOpcode Or RightOpcode))
      If ImmediateBytes.Count > 0 Then Opcodes.AddRange(ImmediateBytes)

      Return Opcodes
   End Function

   'This procedure parses the specified memory operand and returns the resulting opcode and any immediate numeric value present.
   Private Function ParseMemoryOperand(MemoryOperand As String, ByRef ImmediateBytes As List(Of Byte)) As Integer?
      Dim Immediate As String = Nothing
      Dim ImmediateIs8Bit As New Boolean
      Dim Opcode As New Integer?

      MemoryOperand = MemoryOperand.Substring(1, MemoryOperand.Length - 2).Trim()
      If IS_IMMEDIATE_OPERAND(MemoryOperand) Then
         Opcode = BASE_POINTER
         ImmediateBytes.AddRange(BytesFromHexadecimal(MemoryOperand, Is8Bit:=False))
      ElseIf Not MemoryOperand = Nothing Then
         Opcode = MEMORY_OPERANDS.IndexOf(MemoryOperand)
         If Opcode < &H0% Then
            Immediate = MemoryOperand.Substring(MemoryOperand.LastIndexOf("+"c) + 1)
            MemoryOperand = MemoryOperand.Substring(0, MemoryOperand.Length - Immediate.Length)
            Opcode = MEMORY_IMMEDIATE_OPERANDS.IndexOf(MemoryOperand)
            ImmediateIs8Bit = IS_8_BIT_IMMEDIATE(Immediate)
            ImmediateBytes.AddRange(BytesFromHexadecimal(Immediate, ImmediateIs8Bit))
            Opcode = Opcode Or If(ImmediateIs8Bit, OperandOpcodesE.Memory8BitOImmediate, OperandOpcodesE.Memory16BitOImmediate)
         End If
      End If

      Return Opcode
   End Function

   'This procedure returns the specified register operand's index.
   Private Function RegisterOperandToIndex(RegisterOperand As OperandStr) As Integer?
      Try
         Dim Index As New Integer?

         If RegisterOperand.Type = OperandTypesE.Register8Bit Then
            Index = LH_REGISTERS.IndexOf(RegisterOperand.Operand)
         ElseIf RegisterOperand.Type = OperandTypesE.Register16Bit Then
            Index = XP_REGISTERS.IndexOf(RegisterOperand.Operand)
         ElseIf RegisterOperand.Type = OperandTypesE.Segment Then
            Index = SG_REGISTERS.IndexOf(RegisterOperand.Operand)
         End If

         Return Index
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure removes unnecessary white space from the specified instruction and returns it.
   Private Function RemoveWhiteSpace(Instruction As String) As String
      Dim InMemoryOperand As Boolean = False
      Dim Position As Integer = 0

      Instruction = Instruction.Replace(Microsoft.VisualBasic.ControlChars.Tab, " "c)

      Do While Instruction.Contains("  ")
         Instruction = Instruction.Replace("  ", " "c)
      Loop

      Do While Instruction.Contains(" + ")
         Instruction = Instruction.Replace(" + ", "+"c)
      Loop

      Do Until Position >= Instruction.Length
         If Instruction.Chars(Position) = MEMORY_OPERAND_START Then
            InMemoryOperand = True
         ElseIf Instruction.Chars(Position) = MEMORY_OPERAND_END Then
            InMemoryOperand = False
         End If

         If InMemoryOperand AndAlso Instruction.Chars(Position) = " "c Then
            Instruction = Instruction.Remove(Position, 1)
         Else
            Position += 1
         End If
      Loop

      Return Instruction.Replace(", ", ","c).Trim()
   End Function

   'This procedure sets the specified opcode's properties according to the properties of its operands.
   Private Function SetOpcodeProperties(Opcode As Integer, LeftOperand As OperandStr, RightOperand As OperandStr, Optional SetSourceTargetFlag As Boolean = False, Optional FlipSourceTargetFlag As Boolean = False) As Integer
      If LeftOperand.Type = OperandTypesE.Register8Bit OrElse RightOperand.Type = OperandTypesE.Register8Bit Then
         Opcode = Opcode Or OPERAND_8_BIT
      ElseIf LeftOperand.Type = OperandTypesE.Register16Bit OrElse RightOperand.Type = OperandTypesE.Register16Bit Then
         Opcode = Opcode Or OPERAND_16_BIT
      End If

      If SetSourceTargetFlag Then
         If LeftOperand.Type = OperandTypesE.Memory OrElse LeftOperand.Type = OperandTypesE.Segment Then
            Opcode = Opcode Or If(FlipSourceTargetFlag, REGISTER_TARGET, REGISTER_SOURCE)
         ElseIf RightOperand.Type = OperandTypesE.Memory OrElse RightOperand.Type = OperandTypesE.Segment Then
            Opcode = Opcode Or If(FlipSourceTargetFlag, REGISTER_SOURCE, REGISTER_TARGET)
         End If
      End If

      Return Opcode
   End Function
End Class
