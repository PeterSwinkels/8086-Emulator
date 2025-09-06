'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.BitConverter
Imports System.Collections.Generic
Imports System.Linq
Imports System.Math
Imports System.Threading
Imports System.Threading.Tasks

'This class contains the 8086 CPU emulator's procedures.
Public Class CPU8086Class
   'This enumeration lists the type of displacement value operands.
   Private Enum DisplacementTypesE As Byte
      [Byte] = &H1%   'Byte.
      Word = &H2%     'Word.
   End Enum

   'This enumeration lists the flag registers.
   Public Enum FlagRegistersE As Integer
      CF = &H0%           'Carry.
      F1 = &H1%           'Always set to true.
      PF = &H2%           'Parity.
      AF = &H4%           'Auxiliary carry.
      ZF = &H6%           'Zero.
      SF = &H7%           'Sign.
      TF = &H8%           'Trap.
      [IF] = &H9%         'Interrupt.
      DF = &HA%           'Direction.
      [OF] = &HB%         'Overflow.
      All = &HFFFFFFFF%   'All flags.
   End Enum

   'This enumeration lists the memory operands.
   Public Enum MemoryOperandsE As Integer
      LAST = -1    'The last address obtained from a memory operand.
      BX_SI        '[BX + SI].
      BX_DI        '[BX + DI].
      BP_SI        '[BP + SI].
      BP_DI        '[BP + DI].
      SI           '[SI].
      DI           '[DI].
      WORD         '[WORD].
      BX           '[BX].
      BX_SI_BYTE   '[BX + SI + BYTE].
      BX_DI_BYTE   '[BX + DI + BYTE].
      BP_SI_BYTE   '[BP + SI + BYTE].
      BP_DI_BYTE   '[BP + DI + BYTE].
      SI_BYTE      '[SI + BYTE].
      DI_BYTE      '[DI + BYTE].
      BP_BYTE      '[BP + BYTE].
      BX_BYTE      '[BX + BYTE].
      BX_SI_WORD   '[BX + SI + WORD].
      BX_DI_WORD   '[BX + DI + WORD].
      BP_SI_WORD   '[BP + SI + WORD].
      BP_DI_WORD   '[BP + DI + WORD].
      SI_WORD      '[SI + WORD].
      DI_WORD      '[DI + WORD].
      BP_WORD      '[BP + WORD].
      BX_WORD      '[BX + WORD].
   End Enum

   'This enumeration lists the 8086's opcodes.
   Public Enum OpcodesE As Byte
      [LOOP] = &HE2%                  'Loop while CX is zero and decrement CX by one.
      AAA = &H37%                     'ASCII adjust AL for addition.
      AAD = &HD5%                     'ASCII adjust AX for division.
      AAM = &HD4%                     'ASCII adjust AX for multiplication.
      AAS = &H3F%                     'ASCII adjust AL for subtraction.
      ADC_AL_BYTE = &H14%             'Add byte to AL with carry.
      ADC_AX_WORD = &H15%             'Add word to AX with carry.
      ADC_REG16_SRC = &H13%           'Add source to 16 bit register with carry.
      ADC_REG8_SRC = &H12%            'Add source to 8 bit register with carry.
      ADC_TGT_REG16 = &H11%           'Add 16 bit register to target with carry.
      ADC_TGT_REG8 = &H10%            'Add 8 bit register to target with carry.
      ADD_AL_BYTE = &H4%              'Add byte to AL.
      ADD_AX_WORD = &H5%              'Add word to AX.
      ADD_REG16_SRC = &H3%            'Add source to 16 bit register.
      ADD_REG8_SRC = &H2%             'Add source to 8 bit register.
      ADD_TGT_REG16 = &H1%            'Add 16 bit register to target.
      ADD_TGT_REG8 = &H0%             'Add 8 bit register to target.
      AND_AL_BYTE = &H24%             'AND byte with AL.
      AND_AX_WORD = &H25%             'AND word with AX.
      AND_REG16_SRC = &H23%           'AND source with 16 bit register.
      AND_REG8_SRC = &H22%            'AND source with 8 bit register.
      AND_TGT_REG16 = &H21%           'AND 16 bit register with target.
      AND_TGT_REG8 = &H20%            'AND 8 bit register with target.
      CALL_FAR = &H9A%                'Call far.
      CALL_NEAR = &HE8%               'Call near.
      CBW = &H98%                     'Convert byte to word.
      CLC = &HF8%                     'Clear carry flag.
      CLD = &HFC%                     'Clear direction flag.
      CLI = &HFA%                     'Clear interrupt flag.
      CMC = &HF5%                     'Complement carry flag.
      CMP_AL_BYTE = &H3C%             'Compare byte to AL.
      CMP_AX_WORD = &H3D%             'Compare word to AX.
      CMP_REG16_SRC = &H3B%           'Compare source to 16 bit register.
      CMP_REG8_SRC = &H3A%            'Compare source to 8 bit register.
      CMP_TGT_REG16 = &H39%           'Compare 16 bit register to target.
      CMP_TGT_REG8 = &H38%            'Compare 8 bit register to target.
      CMPSB = &HA6%                   'Compare byte string in memory.
      CMPSW = &HA7%                   'Compare word string in memory.
      CS = &H2E%                      'CS segment override.
      CWD = &H99%                     'Convert word to double word.
      DAA = &H27%                     'Decimal adjust for addition.
      DAS = &H2F%                     'Decimal adjust for subtraction.
      DEC_AX = &H48%                  'Decrement AX.
      DEC_BP = &H4D%                  'Decrement BP.
      DEC_BX = &H4B%                  'Decrement BX.
      DEC_CX = &H49%                  'Decrement CX.
      DEC_DI = &H4F%                  'Decrement DI.
      DEC_DX = &H4A%                  'Decrement DX.
      DEC_SI = &H4E%                  'Decrement SI.
      DEC_SP = &H4C%                  'Decrement SP.
      DS = &H3E%                      'DS segment override.
      EXT_INT = &H60%                 'Externally handled interrupt.
      ES = &H26%                      'ES segment override.
      HLT = &HF4%                     'Halt.
      IN_AL_BYTE = &HE4%              'Read from port to AL.
      IN_AL_DX = &HEC%                'Read from port to AL.
      IN_AX_DX = &HED%                'Read from port to AX.
      IN_AX_WORD = &HE5%              'Read from port to AX.
      INC_AX = &H40%                  'Increment AX.
      INC_BP = &H45%                  'Increment BP.
      INC_BX = &H43%                  'Increment BX.
      INC_CX = &H41%                  'Increment CX.
      INC_DI = &H47%                  'Increment DI.
      INC_DX = &H42%                  'Increment DX.
      INC_SI = &H46%                  'Increment SI.
      INC_SP = &H44%                  'Increment SP.
      INT = &HCD%                     'Interrupt.
      INT3 = &HCC%                    'Interrupt 0x03.
      INTO = &HCE%                    'Interrupt 0x04 if OF is true.
      IRET = &HCF%                    'Interrupt return.
      JA = &H77%                      'Jump on no carry and no zero.
      JC = &H72%                      'Jump on carry.
      JCXZ = &HE3%                    'Jump if CX equals zero.
      JG = &H7F%                      'Jump on no zero or if the sign and overflow flags match.
      JL = &H7C%                      'Jump if the sign and overflow flags do not match.
      JMP_FAR = &HEA%                 'Jump far.
      JMP_NEAR = &HE9%                'Jump near.
      JMP_SHORT = &HEB%               'Jump short.
      JNA = &H76%                     'Jump on carry or zero.
      JNC = &H73%                     'Jump on no carry.
      JNG = &H7E%                     'Jump on zero or if the sign and overflow flags do not match.
      JNL = &H7D%                     'Jump if the sign and overflow flags match.
      JNO = &H71%                     'Jump on no overflow.
      JNS = &H79%                     'Jump on no sign.
      JNZ = &H75%                     'Jump on no zero.
      JO = &H70%                      'Jump on overflow.
      JPE = &H7A%                     'Jump on parity equal.
      JPO = &H7B%                     'Jump on parity odd.
      JS = &H78%                      'Jump on sign.
      JZ = &H74%                      'Jump on zero.
      LAHF = &H9F%                    'Load AF, CF, PF, SF and ZF flags into AH.
      LDS = &HC5%                     'Load pointer from source into DS and target.
      LEA = &H8D%                     'Load pointer indicated by source into target.
      LES = &HC4%                     'Load pointer from source into ES and target.
      LOCK = &HF0%                    'Lock the bus.
      LODSB = &HAC%                   'Load byte string from DS:SI to AL.
      LODSW = &HAD%                   'Load word string from DS:SI to AX.
      LOOPE = &HE1%                   'Loop if CX is not zero and decrement CX while the zero flag equals true.
      LOOPNE = &HE0%                  'Loop if CX is not zero and decrement CX while the zero flag equals false.
      MOV_AH_BYTE = &HB4%             'Move byte to AH.
      MOV_AL_BYTE = &HB0%             'Move byte to AL.
      MOV_AL_MEM = &HA0%              'Copy byte from memory to AL.
      MOV_AX_MEM = &HA1%              'Copy word from memory to AX.
      MOV_AX_WORD = &HB8%             'Move word to AX.
      MOV_BH_BYTE = &HB7%             'Move byte to BH.
      MOV_BL_BYTE = &HB3%             'Move byte to BL.
      MOV_BP_WORD = &HBD%             'Move word to BP.
      MOV_BX_WORD = &HBB%             'Move word to BX.
      MOV_CH_BYTE = &HB5%             'Move byte to CH.
      MOV_CL_BYTE = &HB1%             'Move byte to CL.
      MOV_CX_WORD = &HB9%             'Move word to CX.
      MOV_DH_BYTE = &HB6%             'Move byte to DH.
      MOV_DI_WORD = &HBF%             'Move word to DI.
      MOV_DL_BYTE = &HB2%             'Move byte to DL.
      MOV_DX_WORD = &HBA%             'Move word to DX.
      MOV_MEM_AL = &HA2%              'Copy byte from AL to memory.
      MOV_MEM_AX = &HA3%              'Copy word from AX to memory.
      MOV_REG16_SRC = &H8B%           'Move source to 16 bit register.
      MOV_REG8_SRC = &H8A%            'Move source to 8 bit register.
      MOV_SG_SRC = &H8E%              'Move source to segment register.
      MOV_SI_WORD = &HBE%             'Move word to SI.
      MOV_SP_WORD = &HBC%             'Move word to SP.
      MOV_TGT_BYTE = &HC6%            'Move byte to target. 
      MOV_TGT_REG16 = &H89%           'Move 16 bit register to target.
      MOV_TGT_REG8 = &H88%            'Move 8 bit register to target.
      MOV_TGT_SG = &H8C%              'Move segment to target.
      MOV_TGT_WORD = &HC7%            'Move byte to target. 
      MOVSB = &HA4%                   'Moves byte string from DS:SI to ES:DI.
      MOVSW = &HA5%                   'Moves word string from DS:SI to ES:DI.
      MULTI1_TGT_BYTE = &H80%         'Multiple byte operations.
      MULTI1_TGT_WORD = &H81%         'Multiple word operations.
      MULTI1_TGT16_BYTE = &H83%       'Multiple byte operations with a 16 bit target.
      MULTI2_TGT_BYTE = &H82%         'Multiple byte operations.
      MULTI2_TGT_WORD = &HF7%         'Multiple word operations.
      MULTI3_TGT_BYTE = &HF6%         'Multiple byte operations.
      MULTI3_TGT_WORD = &HFF%         'Multiple word operations.
      MULTI4_TGT_BYTE = &HFE%         'Multiple byte operations.
      NOP = &H90%                     'No operation.
      OR_AL_BYTE = &HC%               'OR byte with AL.
      OR_AX_WORD = &HD%               'OR word with AX.
      OR_REG16_SRC = &HB%             'OR source with 16 bit register.
      OR_REG8_SRC = &HA%              'OR source with 8 bit register.
      OR_TGT_REG16 = &H9%             'OR 16 bit register with target.
      OR_TGT_REG8 = &H8%              'OR 8 bit register with target.
      OUT_BYTE_AL = &HE6%             'Write AL to port indicated by byte.
      OUT_DX_AL = &HEE%               'Write AL to port indicated by DX.
      OUT_DX_AX = &HEF%               'Write AX to port indicated by DX.
      OUT_WORD_AX = &HE7%             'Write AX to port indicated by word.
      POP_AX = &H58%                  'Pop AX.
      POP_BP = &H5D%                  'Pop BP.
      POP_BX = &H5B%                  'Pop BX.
      POP_CX = &H59%                  'Pop CX.
      POP_DI = &H5F%                  'Pop DI.
      POP_DS = &H1F%                  'Pop DS.
      POP_DX = &H5A%                  'Pop DX.
      POP_ES = &H7%                   'Pop ES.
      POP_SI = &H5E%                  'Pop SI.
      POP_SP = &H5C%                  'Pop SP.
      POP_SS = &H17%                  'Pop SS.
      POP_TGT = &H8F%                 'Pop to target.
      POPF = &H9D%                    'Pop flags.
      PUSH_AX = &H50%                 'Push AX.
      PUSH_BP = &H55%                 'Push BP.
      PUSH_BX = &H53%                 'Push BX.
      PUSH_CS = &HE%                  'Push CS.
      PUSH_CX = &H51%                 'Push CX.
      PUSH_DI = &H57%                 'Push DI.
      PUSH_DS = &H1E%                 'Push DS.
      PUSH_DX = &H52%                 'Push DX.
      PUSH_ES = &H6%                  'Push ES.
      PUSH_SI = &H56%                 'Push SI.
      PUSH_SP = &H54%                 'Push SP.
      PUSH_SS = &H16%                 'Push SS.
      PUSHF = &H9C%                   'Push flags.
      REPNE = &HF2%                   'Repeat instruction until CX reads zero while the zero flag reads false.
      REPZ = &HF3%                    'Repeat instruction until CX reads zero while the zero flag reads true.
      RETF = &HCB%                    'Return far.
      RETF_BYTES = &HCA%              'Return far and release bytes.
      RETN = &HC3%                    'Return near.
      RETN_BYTES = &HC2%              'Return near and release bytes.
      RSBITS_BYTE_1 = &HD0%           'Rotate or shift a byte's bits by one.
      RSBITS_BYTE_CL = &HD2%          'Rotate or shift a byte's bits by CL.
      RSBITS_WORD_1 = &HD1%           'Rotate or shift a word's bits by one.
      RSBITS_WORD_CL = &HD3%          'Rotate or shift a word's bits by CL.
      SAHF = &H9E%                    'Save AH to AF, CF, PF, SF and ZF flags.
      SBB_AL_BYTE = &H1C%             'Subtract byte from AL with borrow.
      SBB_AX_WORD = &H1D%             'Subtract word from AX with borrow.
      SBB_REG16_SRC = &H1B%           'Subtract source from 16 bit register with borrow.
      SBB_REG8_SRC = &H1A%            'Subtract source from 8 bit register with borrow.
      SBB_TGT_REG16 = &H19%           'Subtract 16 bit register from target with borrow.
      SBB_TGT_REG8 = &H18%            'Subtract 8 bit register from target with borrow.
      SCASB = &HAE%                   'Compare AL to string at ES:DI.
      SCASW = &HAF%                   'Compare AX to string at ES:DI.
      SS = &H36%                      'SS segment override.
      STC = &HF9%                     'Set carry flag.
      STD = &HFD%                     'Set direction flag.
      STI = &HFB%                     'Set interrupt flag.
      STOSB = &HAA%                   'Store AL in string at ES:DI.
      STOSW = &HAB%                   'Store AX in string at ES:DI.
      SUB_AL_BYTE = &H2C%             'Subtract byte from AL.
      SUB_AX_WORD = &H2D%             'Subtract word from AX.
      SUB_REG16_SRC = &H2B%           'Subtract source from 16 bit register.
      SUB_REG8_SRC = &H2A%            'Subtract source from 8 bit register.
      SUB_TGT_REG16 = &H29%           'Subtract 16 bit register from target.
      SUB_TGT_REG8 = &H28%            'Subtract 8 bit register from target.
      TEST_AL_BYTE = &HA8%            'Test AL against byte.
      TEST_AX_WORD = &HA9%            'Test AX against word.
      TEST_SRC_REG16 = &H85%          'Test 16 bit source against target. 
      TEST_SRC_REG8 = &H84%           'Test 8 bit source against target.
      WAIT = &H9B%                    'Wait for coprocessor.
      XCHG_AX_BP = &H95%              'Exchange AX and BP.
      XCHG_AX_BX = &H93%              'Exchange AX and BX.
      XCHG_AX_CX = &H91%              'Exchange AX and CX.
      XCHG_AX_DI = &H97%              'Exchange AX and DI.
      XCHG_AX_DX = &H92%              'Exchange AX and DX.
      XCHG_AX_SI = &H96%              'Exchange AX and SI.
      XCHG_AX_SP = &H94%              'Exchange AX and SP.
      XCHG_REG16 = &H87%              'Exchange words.
      XCHG_REG8 = &H86%               'Exchange bytes.
      XLAT = &HD7%                    'Replace AL with a byte from a table pointed at by BX.
      XOR_AL_BYTE = &H34%             'XOR byte with AL.
      XOR_AX_WORD = &H35%             'XOR word with AX.
      XOR_REG16_SRC = &H33%           'XOR source with 16 bit register.
      XOR_REG8_SRC = &H32%            'XOR source with 8 bit register.
      XOR_TGT_REG16 = &H31%           'XOR 16 bit register with target.
      XOR_TGT_REG8 = &H30%            'XOR 8 bit register with target.
   End Enum

   'This enumeration lists the supported operand pairs.
   Private Enum OperandPairsE As Byte
      MemReg_Reg_8     '8 bit memory/register and register.
      MemReg_Reg_16    '16 bit memory/register and register.
      Reg_MemReg_8     '8 bit register and memory/register.
      Reg_MemReg_16    '16 bit register and memory/register.
      AL_Byte          'AL register and byte.
      AX_Word          'AL register and word.
      MemReg_Byte      'Memory/register and byte.
      MemReg_Word      'Memory/register and word.
      MemReg_16_Byte   '16 bit Memory/register and byte.
   End Enum

   'This enumeration lists the operations for the 0x80, 0x81 and 0x83 opcodes.
   Private Enum Operations80_83E As Byte
      ADD     'Add.
      [OR]    'OR.
      ADC     'Add with carry.
      SBB     'Subtract with borrow.
      [AND]   'AND.
      [SUB]   'Subtract.
      [XOR]   'XOR.
      [CMP]   'Compare.
   End Enum

   'This enumeration lists the operations for the 0xF6 and 0xF7 opcodes.
   Private Enum OperationsF6_F7E As Byte
      TEST    'Test.
      None    'No operation.
      [NOT]   'NOT.
      NEG     'Negate.
      MUL     'Multiply.
      IMUL    'Signed multiply.
      DIV     'Divide.
      IDIV    'Signed divide.
   End Enum

   'This enumeration lists the operations for the 0xFE:0x00-0xFE:0xBF and 0xFF:0x00-0xFF:0xBF opcodes.
   Private Enum OperationsFEFF00_FEFFBFE As Byte
      INC              'Increment byte/word.
      DEC              'Decrement byte/word.
      CALL_WORD_NEAR   'Call near.
      CALL_DWORD_FAR   'Call far.
      JMP_WORD_NEAR    'Jump near.
      JMP_DWORD_FAR    'Jump far.
      PUSH             'Push.
      None             'No operation.
   End Enum

   'This enumeration lists the operations for the 0xFE:0x00-0xFE:0xBF and 0xFF:0xC0-0xFF:0xFF opcodes.
   Private Enum OperationsFEFFC0_FEFFFFE As Byte
      INC      'Increment.
      DEC      'Decrement.
      [CALL]   'Call.
      None1    'No operation.
      JMP      'Jump.
      None2    'No operation.
      PUSH     'Push.
      None3    'No operation.
   End Enum

   'This enumeration lists the 16 bit registers.
   Public Enum Registers16BitE As Integer
      AX   'Accumulator.
      CX   'Counter.
      DX   'Data.
      BX   'Base.
      SP   'Stack pointer.
      BP   'Base pointer.
      SI   'Source index.
      DI   'Destination index.
      IP   'Instruction pointer.
   End Enum

   'This enumeration lists the bit rotate/shift operations.
   Private Enum RotateAndShiftOperationsE As Byte
      ROL    'Rotate left into carry flag.
      ROR    'Rotate right into carry flag.
      RCL    'Rotate left into carry flag.
      RCR    'Rotate right into carry flag.
      SHL    'Shift left into carry flag.
      SHR    'Shift right into carry flag.
      None   'No operation.
      SAR    'Shift right into carry and set left most bit to equal the sign flag.
   End Enum

   'This enumeration lists the segment registers.
   Public Enum SegmentRegistersE As Integer
      ES      'Extra segment.
      CS      'Code segment.
      SS      'Stack segment.
      DS      'Data segment.
   End Enum

   'This enumeration lists the 8 bit sub registers.
   Public Enum SubRegisters8BitE As Integer
      AL   'Low accumulator.
      CL   'Low counter.
      DL   'Low data.
      BL   'Low base.
      AH   'High accumulator.
      CH   'High counter.
      DH   'High data.
      BH   'High base.
   End Enum

   'This structure defines a relative address and absolute flat address.
   Public Structure AddressesStr
      Public Address As Integer?       'Defines a relative address.
      Public FlatAddress As Integer?   'Defines a flat address.
   End Structure

   'This structure defines a source and target.
   Private Structure OperandPairStr
      Public Address As Integer?             'Defines a relative address for a memory operand.
      Public Displacement As Integer?        'Defines an optional displacement value that is part of one of the operands.
      Public DisplacementIs8Bit As Boolean   'Indicates whether the displacement is 8 or 16 bits.
      Public Is8Bit As Boolean               'Indicates whether the source and target are 8 or 16 bits.
      Public FlatAddress As Integer?         'Defines an absolute flat address for a memory operand.
      Public NewValue As Integer             'Defines the result of an operation performed using the first and second values.
      Public Operand1 As Object              'Defines the first operand.
      Public Operand2 As Object              'Defines the second operand.
      Public Value1 As Integer               'Defines the value referred to be the first operand.
      Public Value2 As Integer               'Defines the value referred to be the second operand.
   End Structure

   Public Const KEYBOARD As Integer = &H9%               'Defines the keyboard hardware interrupt number.
   Public Const SYSTEM_TIMER As Integer = &H8%           'Defines the system timer's interrupt number.

   Private Const BREAK_POINT As Integer = &H3%           'Defines the break point interrupt number.
   Private Const DIVIDE_BY_ZERO As Integer = &H0%        'Defines the divide by zero interrupt number.
   Private Const LOW_FLAG_BITS As Integer = &HD7%        'Defines the bits used in the flag register's lower byte.
   Private Const OVERFLOW_TRAP As Integer = &H4%         'Defines the overflow trap interrupt number. 
   Private Const SINGLE_STEP As Integer = &H1%           'Defines the single step interrupt number.

   Public Const ADDRESS_MASK As Integer = &HFFFFF%      'Defines the 20 bits used to address memory.
   Public Const INVALID_OPCODE As Integer = &H6%        'Defines the invalid opcode interrupt number.

   Public Clock As New Task(AddressOf Execute)                                     'Contains the CPU clock.
   Public ClockToken As New CancellationTokenSource                                'Indicates whether or not to stop the CPU.
   Public HardwareInterrupts As New List(Of Integer)                               'Contains a list of pending interrupts triggered by emulated hardware.
   Public Memory() As Byte = Enumerable.Repeat(CByte(&H0%), &H100000%).ToArray()   'Contains the memory used by the emulated 8086 CPU.
   Public Tracing As Boolean = False                                               'Indicates whether or not tracing is enabled.

   Public Event Halt()                                                                   'Defines the halt event.
   Public Event Interrupt(Number As Integer, AH As Integer)                              'Defines the interrupt event.
   Public Event ReadIOPort(Port As Integer, ByRef Value As Integer, Is8Bit As Boolean)   'Defines the IO port read event.
   Public Event Trace(FlatCSIP As Integer)                                               'Defines the trace event.
   Public Event WriteIOPort(Port As Integer, Value As Integer, Is8Bit As Boolean)        'Defines the IO port write event.

   Public ReadOnly GET_FLAT_CS_IP As Func(Of Integer) = Function() (CInt(Registers(CPU8086Class.SegmentRegistersE.CS)) << &H4%) + CInt(Registers(Registers16BitE.IP)) And ADDRESS_MASK  'Returns the flat memory address for the emulated CPU's CS:IP registers.
   Public ReadOnly GET_WORD As Func(Of Integer, Integer) = Function(Address As Integer) (ToUInt16(Memory, Address And ADDRESS_MASK))                                                    'Returns the word at the specified address.

   'This procedure initializes the CPU.
   Public Sub New()
      For Each Register As Registers16BitE In [Enum].GetValues(GetType(Registers16BitE))
         Registers(Register, NewValue:=&H0%)
      Next Register

      For Each Register As FlagRegistersE In [Enum].GetValues(GetType(FlagRegistersE))
         Registers(Register, NewValue:=False)
      Next Register

      For Each Register As SegmentRegistersE In [Enum].GetValues(GetType(SegmentRegistersE))
         Registers(Register, NewValue:=&H0%)
      Next Register

      Registers(FlagRegistersE.IF, NewValue:=True)

      Memory = Enumerable.Repeat(CByte(&H0%), &H100000%).ToArray()
   End Sub

   'This procedure returns the memory address indicated by the specified operand and current data segment.
   Public Function AddressFromOperand(Operand As MemoryOperandsE, Optional Displacement As Integer? = Nothing, Optional DisplacementIs8Bit As Boolean = False) As AddressesStr
      Dim Addresses As New AddressesStr
      Dim Override As New SegmentRegistersE?
      Dim Segment As SegmentRegistersE = SegmentRegistersE.DS
      Static LastAddresses As New AddressesStr

      If Operand = MemoryOperandsE.LAST Then
         Addresses = LastAddresses
         LastAddresses.Address = Nothing
         LastAddresses.FlatAddress = Nothing
      Else
         Override = SegmentOverride()

         Select Case Operand
            Case MemoryOperandsE.BX_SI,
         MemoryOperandsE.BX_SI_BYTE,
         MemoryOperandsE.BX_SI_WORD
               Addresses.Address = CInt(Registers(Registers16BitE.BX)) + CInt(Registers(Registers16BitE.SI))
            Case MemoryOperandsE.BX_DI,
         MemoryOperandsE.BX_DI_BYTE,
         MemoryOperandsE.BX_DI_WORD
               Addresses.Address = CInt(Registers(Registers16BitE.BX)) + CInt(Registers(Registers16BitE.DI))
            Case MemoryOperandsE.BP_SI,
         MemoryOperandsE.BP_SI_BYTE,
         MemoryOperandsE.BP_SI_WORD
               Addresses.Address = CInt(Registers(Registers16BitE.BP)) + CInt(Registers(Registers16BitE.SI))
               Segment = SegmentRegistersE.SS
            Case MemoryOperandsE.BP_DI,
         MemoryOperandsE.BP_DI_BYTE,
         MemoryOperandsE.BP_DI_WORD
               Addresses.Address = CInt(Registers(Registers16BitE.BP)) + CInt(Registers(Registers16BitE.DI))
               Segment = SegmentRegistersE.SS
            Case MemoryOperandsE.SI,
         MemoryOperandsE.SI_BYTE,
         MemoryOperandsE.SI_WORD
               Addresses.Address = CInt(Registers(Registers16BitE.SI))
            Case MemoryOperandsE.DI,
         MemoryOperandsE.DI_BYTE,
         MemoryOperandsE.DI_WORD
               Addresses.Address = CInt(Registers(Registers16BitE.DI))
            Case MemoryOperandsE.WORD
               Addresses.Address = GetWordCSIP()
            Case MemoryOperandsE.BP_BYTE,
         MemoryOperandsE.BP_WORD
               Addresses.Address = CInt(Registers(Registers16BitE.BP))
               Segment = SegmentRegistersE.SS
            Case MemoryOperandsE.BX,
         MemoryOperandsE.BX_BYTE,
         MemoryOperandsE.BX_WORD
               Addresses.Address = CInt(Registers(Registers16BitE.BX))
         End Select

         If Displacement IsNot Nothing Then
            If DisplacementIs8Bit Then
               Displacement = ConvertWidening(Displacement.Value, Is8Bit:=True)
            End If

            Addresses.Address = (Addresses.Address + Displacement.Value) And &HFFFF%
         End If

         Addresses.FlatAddress = ((CInt(If(Override Is Nothing, Registers(Segment), Registers(Override))) << &H4%) + Addresses.Address) And ADDRESS_MASK
         LastAddresses = Addresses
      End If

      Return Addresses
   End Function

   'This procedure adjusts the flags register based on the specified values.
   Private Sub AdjustFlags(OldValue As Integer, NewValue As Integer, Optional Is8Bit As Boolean = True, Optional NewCarryFlag As Integer? = Nothing, Optional PreserveCarryFlag As Boolean = False, Optional NewOverflowFlag As Integer? = Nothing, Optional Subtraction As Boolean = True)
      Dim SignMask As Integer = If(Is8Bit, &H80%, &H8000%)

      Registers(FlagRegistersE.PF, NewValue:=(BitCount(NewValue And &HFF%) Mod &H2%) = &H0%)
      Registers(FlagRegistersE.SF, NewValue:=(NewValue And SignMask) = SignMask)
      Registers(FlagRegistersE.ZF, NewValue:=(NewValue And If(Is8Bit, &HFF%, &HFFFF%)) = &H0%)

      If NewOverflowFlag Is Nothing Then
         If Subtraction Then
            Registers(FlagRegistersE.OF, NewValue:=((NewValue < &H0%) AndAlso (Abs(NewValue) >= SignMask)))
         Else
            Registers(FlagRegistersE.OF, NewValue:=((NewValue >= &H0%) AndAlso ((NewValue And If(Is8Bit, &HFF%, &HFFFF%)) >= SignMask)))
         End If
      Else
         Registers(FlagRegistersE.OF, NewValue:=NewOverflowFlag)
      End If

      If Not PreserveCarryFlag Then
         If NewCarryFlag Is Nothing Then
            Registers(FlagRegistersE.CF, NewValue:=NewValue > If(Is8Bit, &HFF%, &HFFFF%) OrElse NewValue < &H0%)
         Else
            Registers(FlagRegistersE.CF, NewValue:=NewCarryFlag)
         End If
      End If
   End Sub

   'This procedure returns the number bits in the specified value.
   Private Function BitCount(Value As Integer) As Integer
      Dim Count As Integer = 0

      Do While Value > &H0%
         Count += Value And &H1%
         Value >>= &H1%
      Loop

      Return Count
   End Function

   'This procedure shifts/rotates bits once in the specified direction and returns the last/first bit shifted out.
   Private Function BitRotateOrShift(Bits As Integer, Is8Bit As Boolean, Optional Right As Boolean = True, Optional Rotate As Boolean = True, Optional ReplicateSign As Boolean = False) As Integer
      Dim NewBits As Integer = Bits And If(Is8Bit, &HFF%, &HFFFF%)
      Dim SignBit As Integer = NewBits And If(Is8Bit, &H80%, &H8000%)
      Dim NewCarryFlag As Integer = If(Right, NewBits And &H1%, SignBit)

      If Not Right Then NewCarryFlag >>= If(Is8Bit, &H7%, &HF%)
      NewBits = If(Right, NewBits >> &H1%, NewBits << &H1%)
      If Rotate Then NewBits = NewBits Or If(Right, NewCarryFlag << If(Is8Bit, &H7%, &HF%), NewCarryFlag)
      If ReplicateSign Then NewBits = NewBits Or (SignBit >> If(Is8Bit, &H7%, &HF%))

      AdjustFlags(Bits, NewBits, Is8Bit, NewCarryFlag)

      Return NewBits And If(Is8Bit, &HFF%, &HFFFF%)
   End Function

   'This procedure converts the specified byte/word to a word/dword and returns the result.
   Public Function ConvertWidening(Value As Integer, Is8Bit As Boolean) As Integer
      Value = Value And If(Is8Bit, &HFF%, &HFFFF%)

      If Is8Bit AndAlso Value > &H7F% Then
         Value -= &H100%
      ElseIf Not Is8Bit AndAlso Value > &H7FFF% Then
         Value -= &H10000%
      End If

      Return Value And If(Is8Bit, &HFFFF%, &HFFFFFFFF%)
   End Function

   'This procedure gives the CPU the command to execute instructions and returns the address of the next instruction to be executed.
   Public Function Execute() As Task(Of Integer)
      Try
         Do Until ClockToken.Token.IsCancellationRequested
            If Tracing Then RaiseEvent Trace(GET_FLAT_CS_IP())

            Try
               Do While HardwareInterrupts.Any
                  ExecuteInterrupt(OpcodesE.INT, Number:=HardwareInterrupts.First)
                  HardwareInterrupts.RemoveAt(0)
               Loop
            Catch
            End Try

            If Not ExecuteOpcode() Then
               ExecuteInterrupt(OpcodesE.INT, Number:=INVALID_OPCODE)
            End If

            If Tracing Then RaiseEvent Trace(Nothing)
         Loop

         Return Task.FromResult(GET_FLAT_CS_IP())
      Catch ExceptionO As Exception
         SyncLock Synchronizer
            CPUEvent.Append($"{ExceptionO.Message}{Environment.NewLine}")
         End SyncLock
      End Try

      Return Nothing
   End Function

   'This procedure executes control flow instructions and returns whether or not it succeeded.
   Private Function ExecuteControlFlowOpcode(Opcode As OpcodesE) As Boolean
      Dim Offset As New Integer
      Dim Operand As New Integer
      Dim Segment As New Integer

      Select Case Opcode
         Case OpcodesE.CALL_FAR
            Offset = GetWordCSIP()
            Segment = GetWordCSIP()
            Stack(Push:=CInt(Registers(SegmentRegistersE.CS)))
            Stack(Push:=CInt(Registers(Registers16BitE.IP)))
            Registers(SegmentRegistersE.CS, NewValue:=Segment)
            Registers(Registers16BitE.IP, NewValue:=Offset)
         Case OpcodesE.CALL_NEAR
            Operand = GetWordCSIP()
            Stack(Push:=CInt(Registers(Registers16BitE.IP)))
            Registers(Registers16BitE.IP, NewValue:=CInt(Registers(Registers16BitE.IP)) + Operand)
         Case OpcodesE.JMP_FAR
            Offset = GetWordCSIP()
            Segment = GetWordCSIP()
            Registers(SegmentRegistersE.CS, NewValue:=Segment)
            Registers(Registers16BitE.IP, NewValue:=Offset)
         Case OpcodesE.JMP_NEAR
            Operand = GetWordCSIP()
            If Operand > &H7FFF% Then Operand -= &H10000%
            Registers(Registers16BitE.IP, NewValue:=CInt(Registers(Registers16BitE.IP)) + Operand)
         Case OpcodesE.JCXZ, OpcodesE.JMP_SHORT, OpcodesE.JO To OpcodesE.JG, OpcodesE.LOOP, OpcodesE.LOOPE, OpcodesE.LOOPNE
            Operand = GetByteCSIP()
            If Operand > &H7F% Then Operand -= &H100%
            Offset = CInt(Registers(Registers16BitE.IP)) + Operand

            Select Case Opcode
               Case OpcodesE.JA
                  If Not (CBool(Registers(FlagRegistersE.CF)) OrElse CBool(Registers(FlagRegistersE.ZF))) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JC
                  If CBool(Registers(FlagRegistersE.CF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JCXZ
                  If CInt(Registers(Registers16BitE.CX)) = &H0% Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JG
                  If CBool(Registers(FlagRegistersE.OF)) = CBool(Registers(FlagRegistersE.SF)) AndAlso (Not CBool(Registers(FlagRegistersE.ZF))) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JL
                  If Not CBool(Registers(FlagRegistersE.OF)) = CBool(Registers(FlagRegistersE.SF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JNA
                  If CBool(Registers(FlagRegistersE.CF)) OrElse CBool(Registers(FlagRegistersE.ZF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JNC
                  If Not CBool(Registers(FlagRegistersE.CF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JNG
                  If Not (CBool(Registers(FlagRegistersE.OF)) = CBool(Registers(FlagRegistersE.SF)) AndAlso (Not CBool(Registers(FlagRegistersE.ZF)))) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JNL
                  If CBool(Registers(FlagRegistersE.OF)) = CBool(Registers(FlagRegistersE.SF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JNO
                  If Not CBool(Registers(FlagRegistersE.OF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JNS
                  If Not CBool(Registers(FlagRegistersE.SF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JNZ
                  If Not CBool(Registers(FlagRegistersE.ZF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JO
                  If CBool(Registers(FlagRegistersE.OF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JPE
                  If CBool(Registers(FlagRegistersE.PF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JPO
                  If Not CBool(Registers(FlagRegistersE.PF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JS
                  If CBool(Registers(FlagRegistersE.SF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JZ
                  If CBool(Registers(FlagRegistersE.ZF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.JMP_SHORT
                  Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.LOOP
                  Registers(Registers16BitE.CX, NewValue:=CInt(Registers(Registers16BitE.CX)) - &H1%)
                  If CInt(Registers(Registers16BitE.CX)) > &H0% Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.LOOPE
                  Registers(Registers16BitE.CX, NewValue:=CInt(Registers(Registers16BitE.CX)) - &H1%)
                  If CInt(Registers(Registers16BitE.CX)) > &H0% AndAlso CBool(Registers(FlagRegistersE.ZF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
               Case OpcodesE.LOOPNE
                  Registers(Registers16BitE.CX, NewValue:=CInt(Registers(Registers16BitE.CX)) - &H1%)
                  If CInt(Registers(Registers16BitE.CX)) > &H0% AndAlso Not CBool(Registers(FlagRegistersE.ZF)) Then Registers(Registers16BitE.IP, NewValue:=Offset)
            End Select
         Case OpcodesE.RETF
            Registers(Registers16BitE.IP, NewValue:=Stack())
            Registers(SegmentRegistersE.CS, NewValue:=Stack())
         Case OpcodesE.RETF_BYTES
            Operand = GetByteCSIP()
            Registers(Registers16BitE.IP, NewValue:=Stack())
            Registers(SegmentRegistersE.CS, NewValue:=Stack())
            Registers(Registers16BitE.SP, NewValue:=CInt(Registers(Registers16BitE.SP)) + Operand)
         Case OpcodesE.RETN
            Registers(Registers16BitE.IP, NewValue:=Stack())
         Case OpcodesE.RETN_BYTES
            Operand = GetByteCSIP()
            Registers(Registers16BitE.IP, NewValue:=Stack())
            Registers(Registers16BitE.SP, NewValue:=CInt(Registers(Registers16BitE.SP)) + Operand)
         Case Else
            Return False
      End Select

      Return True
   End Function

   'This procedure executes the specified interrupt call.
   Public Sub ExecuteInterrupt(Opcode As OpcodesE, Optional Number As Integer? = Nothing)
      Dim Address As New Integer

      If (Opcode = OpcodesE.INTO AndAlso CBool(Registers(FlagRegistersE.OF))) OrElse (Not Opcode = OpcodesE.INTO) Then
         Stack(Push:=CInt(Registers(FlagRegistersE.All)))
         Registers(FlagRegistersE.IF, NewValue:=False)
         Registers(FlagRegistersE.TF, NewValue:=False)

         Select Case Opcode
            Case OpcodesE.INT
               If Number Is Nothing Then Number = GetByteCSIP()
            Case OpcodesE.INT3
               Number = BREAK_POINT
            Case OpcodesE.INTO
               Number = OVERFLOW_TRAP
         End Select

         Stack(Push:=CInt(Registers(SegmentRegistersE.CS)))
         Stack(Push:=CInt(Registers(Registers16BitE.IP)))
         Address = CInt(Number) * &H4%
         Registers(SegmentRegistersE.CS, NewValue:=GET_WORD(Address + &H2%))
         Registers(Registers16BitE.IP, NewValue:=GET_WORD(Address))
      End If
   End Sub

   'This procedure executes the opcode located at CS:IP and returns whether or not it succeeded.
   Private Function ExecuteMultiOpcode(Opcode As OpcodesE) As Boolean
      Dim AL As New Integer
      Dim AX As New Integer
      Dim DX As New Integer
      Dim LargeValue As New Long
      Dim Operand As New Integer
      Dim OperandPair As New OperandPairStr
      Dim Operation As New Object
      Dim Segment As New Integer
      Dim Value As New Integer

      Select Case Opcode
         Case OpcodesE.MULTI1_TGT_BYTE, OpcodesE.MULTI1_TGT_WORD, OpcodesE.MULTI1_TGT16_BYTE, OpcodesE.MULTI2_TGT_BYTE
            If Opcode = OpcodesE.MULTI2_TGT_BYTE Then Opcode = OpcodesE.MULTI1_TGT_BYTE

            Operand = GetByteCSIP()
            Operation = (Operand And &H3F%) >> &H3%

            If Opcode = OpcodesE.MULTI1_TGT16_BYTE Then
               OperandPair = GetValues(GetOperandPair(CByte((Opcode >> &H2%) Xor &H6%), CByte(Operand), IsMemReg16Byte:=True))
            Else
               OperandPair = GetValues(GetOperandPair(CByte(Opcode Xor &H6%), CByte(Operand)))
            End If

            With OperandPair
               Select Case DirectCast(CByte(Operation), Operations80_83E)
                  Case Operations80_83E.ADC
                     .NewValue = .Value1 + .Value2
                     If CBool(Registers(FlagRegistersE.CF)) Then .NewValue += &H1%
                     AdjustFlags(.Value1, .NewValue, .Is8Bit,,,, Subtraction:=False)
                  Case Operations80_83E.ADD
                     .NewValue = .Value1 + .Value2
                     AdjustFlags(.Value1, .NewValue, .Is8Bit,,,, Subtraction:=False)
                  Case Operations80_83E.AND
                     .NewValue = .Value1 And .Value2
                     AdjustFlags(.Value1, .NewValue, .Is8Bit)
                  Case Operations80_83E.CMP, Operations80_83E.SUB
                     .NewValue = .Value1 - .Value2
                     AdjustFlags(.Value1, .NewValue, .Is8Bit)
                  Case Operations80_83E.OR
                     .NewValue = .Value1 Or .Value2
                     AdjustFlags(.Value1, .NewValue, .Is8Bit)
                  Case Operations80_83E.SBB
                     .NewValue = .Value1 - .Value2
                     If CBool(Registers(FlagRegistersE.CF)) Then .NewValue -= &H1%
                     AdjustFlags(.Value1, .NewValue, .Is8Bit)
                  Case Operations80_83E.XOR
                     .NewValue = .Value1 Xor .Value2
                     AdjustFlags(.Value1, .NewValue, .Is8Bit)
               End Select

               Select Case DirectCast(CByte(Operation), Operations80_83E)
                  Case Operations80_83E.ADC,
                  Operations80_83E.ADD,
                  Operations80_83E.AND,
                  Operations80_83E.OR,
                  Operations80_83E.SBB,
                  Operations80_83E.SUB,
                  Operations80_83E.XOR
                     SetNewValue(OperandPair)
               End Select
            End With
         Case OpcodesE.MULTI3_TGT_BYTE, OpcodesE.MULTI2_TGT_WORD
            Operand = GetByteCSIP()
            Operation = (Operand And &H3F%) >> &H3%
            OperandPair = GetValues(GetOperandPair(CByte(Opcode Xor &H8%), CByte(Operand)))

            With OperandPair
               Select Case DirectCast(CByte(Operation), OperationsF6_F7E)
                  Case OperationsF6_F7E.DIV,
                  OperationsF6_F7E.IDIV,
                  OperationsF6_F7E.IMUL,
                  OperationsF6_F7E.MUL,
                  OperationsF6_F7E.NEG,
                  OperationsF6_F7E.None,
                  OperationsF6_F7E.NOT
                     Registers(Registers16BitE.IP, NewValue:=CInt(Registers(Registers16BitE.IP)) - If(.Is8Bit, &H1%, &H2%))
               End Select

               Select Case DirectCast(CByte(Operation), OperationsF6_F7E)
                  Case OperationsF6_F7E.DIV
                     If .Value1 = &H0% Then
                        ExecuteInterrupt(OpcodesE.INT, DIVIDE_BY_ZERO)
                     Else
                        If .Is8Bit Then
                           Value = CInt(Registers(Registers16BitE.AX))
                           .NewValue = CInt(Floor(Value / .Value1))
                           Registers(SubRegisters8BitE.AL, NewValue:= .NewValue)
                           Registers(SubRegisters8BitE.AH, NewValue:=Value Mod .Value1)
                        Else
                           LargeValue = (CLng(Registers(Registers16BitE.DX)) << &H10%) Or CLng(Registers(Registers16BitE.AX))
                           .NewValue = CInt(Floor(LargeValue / .Value1))
                           Registers(Registers16BitE.AX, NewValue:= .NewValue)
                           Registers(Registers16BitE.DX, NewValue:=LargeValue Mod .Value1)
                        End If
                     End If
                  Case OperationsF6_F7E.IDIV
                     If .Value1 = &H0% Then
                        ExecuteInterrupt(OpcodesE.INT, DIVIDE_BY_ZERO)
                     Else
                        If .Is8Bit Then
                           AX = CInt(Registers(Registers16BitE.AX))
                           If .Value1 > &H7F% Then .Value1 -= &H100%
                           .NewValue = CInt(Truncate(AX / .Value1))

                           If .NewValue < SByte.MinValue OrElse .NewValue > SByte.MaxValue Then
                              ExecuteInterrupt(OpcodesE.INT, DIVIDE_BY_ZERO)
                           Else
                              Registers(SubRegisters8BitE.AL, NewValue:= .NewValue)
                              Registers(SubRegisters8BitE.AH, NewValue:=CInt(LargeValue Mod .Value1))
                           End If
                        Else
                           LargeValue = (CLng(Registers(Registers16BitE.DX)) << &H10%) Or CLng(Registers(Registers16BitE.AX))
                           If .Value1 > &H7FFF% Then .Value1 -= &H10000%
                           .NewValue = CInt(Truncate(LargeValue / .Value1))

                           If .NewValue < Short.MinValue OrElse .NewValue > Short.MaxValue Then
                              ExecuteInterrupt(OpcodesE.INT, DIVIDE_BY_ZERO)
                           Else
                              Registers(Registers16BitE.AX, NewValue:= .NewValue)
                              Registers(Registers16BitE.DX, NewValue:=CInt(LargeValue Mod .Value1))
                           End If
                        End If
                     End If
                  Case OperationsF6_F7E.IMUL
                     If .Is8Bit Then
                        AL = CInt(Registers(SubRegisters8BitE.AL))
                        If AL > &H7F% Then AL -= &H100%
                        If .Value1 > &H7F% Then .Value1 -= &H100%

                        .NewValue = AL * .Value1
                        If .NewValue > &H7F% Then .NewValue -= &H100%
                        Registers(Registers16BitE.AX, NewValue:= .NewValue)

                        Registers(FlagRegistersE.CF, NewValue:=CInt(.NewValue > &H7FFF&))
                        Registers(FlagRegistersE.OF, NewValue:=CInt(.NewValue > &H7FFF&))
                     Else
                        AX = CInt(Registers(Registers16BitE.AX))
                        If AX > &H7FFF% Then AX -= &H10000%
                        If .Value1 > &H7FFF% Then .Value1 -= &H10000%

                        LargeValue = CLng(AX) * .Value1

                        AX = CInt(LargeValue And &HFFFF%)
                        DX = CInt(LargeValue >> &H10%)
                        If AX > &H7FFF% Then AX -= &H10000%
                        If DX > &H7FFF% Then DX -= &H10000%
                        Registers(Registers16BitE.AX, NewValue:=AX)
                        Registers(Registers16BitE.DX, NewValue:=DX)

                        Registers(FlagRegistersE.CF, NewValue:=CInt(LargeValue > &HFFFFFFFF&))
                        Registers(FlagRegistersE.OF, NewValue:=CInt(LargeValue > &HFFFFFFFF&))
                     End If
                  Case OperationsF6_F7E.MUL
                     If .Is8Bit Then
                        .NewValue = CInt(Registers(SubRegisters8BitE.AL)) * .Value1
                        Registers(Registers16BitE.AX, NewValue:= .NewValue)

                        Registers(FlagRegistersE.CF, NewValue:=CInt(.NewValue > &HFFFF&))
                        Registers(FlagRegistersE.OF, NewValue:=CInt(.NewValue > &HFFFF&))
                     Else
                        LargeValue = CLng(CInt(Registers(Registers16BitE.AX))) * .Value1

                        Registers(Registers16BitE.AX, NewValue:=LargeValue And &HFFFF%)
                        Registers(Registers16BitE.DX, NewValue:=LargeValue >> &H10%)

                        Registers(FlagRegistersE.CF, NewValue:=CInt(LargeValue > &HFFFFFFFF&))
                        Registers(FlagRegistersE.OF, NewValue:=CInt(LargeValue > &HFFFFFFFF&))
                     End If
                  Case OperationsF6_F7E.NEG
                     .NewValue = &H0% - .Value1

                     AdjustFlags(.Value1, .NewValue, Is8Bit:= .Is8Bit)
                  Case OperationsF6_F7E.None
                     Return False
                  Case OperationsF6_F7E.NOT
                     .NewValue = Not .Value1
                  Case OperationsF6_F7E.TEST
                     .NewValue = .Value1 And .Value2

                     AdjustFlags(.Value1, .NewValue, Is8Bit:= .Is8Bit, NewCarryFlag:=&H0%,, NewOverflowFlag:=&H0%)
               End Select

               Select Case DirectCast(CByte(Operation), OperationsF6_F7E)
                  Case OperationsF6_F7E.NEG, OperationsF6_F7E.NOT
                     SetNewValue(OperandPair)
               End Select
            End With
         Case OpcodesE.MULTI4_TGT_BYTE, OpcodesE.MULTI3_TGT_WORD
            Operand = GetByteCSIP()
            Operation = (Operand And &H3F%) >> &H3%
            OperandPair = GetValues(GetOperandPair(CByte(Opcode Xor &H6%), CByte(Operand)))
            With OperandPair
               If Operand < &HC0% Then
                  Select Case DirectCast(CByte(Operation), OperationsFEFF00_FEFFBFE)
                     Case OperationsFEFF00_FEFFBFE.CALL_DWORD_FAR
                        Segment = GET_WORD(CInt(.FlatAddress) + &H2%)
                        Stack(Push:=CInt(Registers(SegmentRegistersE.CS)))
                        Stack(Push:=CInt(Registers(Registers16BitE.IP)))
                        Registers(SegmentRegistersE.CS, NewValue:=Segment)
                        Registers(Registers16BitE.IP, NewValue:= .Value1)
                     Case OperationsFEFF00_FEFFBFE.CALL_WORD_NEAR
                        Stack(Push:=CInt(Registers(Registers16BitE.IP)))
                        Registers(Registers16BitE.IP, NewValue:= .Value1)
                     Case OperationsFEFF00_FEFFBFE.DEC
                        .NewValue = .Value1 - &H1%
                        AdjustFlags(.Value1, .NewValue, .Is8Bit,, PreserveCarryFlag:=True)
                     Case OperationsFEFF00_FEFFBFE.INC
                        .NewValue = .Value1 + &H1%
                        AdjustFlags(.Value1, .NewValue, .Is8Bit,, PreserveCarryFlag:=True,, Subtraction:=False)
                     Case OperationsFEFF00_FEFFBFE.JMP_DWORD_FAR
                        Registers(SegmentRegistersE.CS, NewValue:=GET_WORD(CInt(.FlatAddress) + &H2%))
                        Registers(Registers16BitE.IP, NewValue:= .Value1)
                     Case OperationsFEFF00_FEFFBFE.JMP_WORD_NEAR
                        Registers(Registers16BitE.IP, NewValue:= .Value1)
                     Case OperationsFEFF00_FEFFBFE.PUSH
                        Stack(Push:= .Value1)
                     Case Else
                        Return False
                  End Select
               Else
                  Select Case DirectCast(CByte(Operation), OperationsFEFFC0_FEFFFFE)
                     Case OperationsFEFFC0_FEFFFFE.CALL
                        Stack(Push:=CInt(Registers(Registers16BitE.IP)))
                        Registers(Registers16BitE.IP, NewValue:= .Value1)
                     Case OperationsFEFFC0_FEFFFFE.DEC
                        .NewValue = .Value1 - &H1%
                        AdjustFlags(.Value1, .NewValue, .Is8Bit,, PreserveCarryFlag:=True)
                     Case OperationsFEFFC0_FEFFFFE.INC
                        .NewValue = .Value1 + &H1%
                        AdjustFlags(.Value1, .NewValue, .Is8Bit,, PreserveCarryFlag:=True,, Subtraction:=False)
                     Case OperationsFEFFC0_FEFFFFE.JMP
                        Registers(Registers16BitE.IP, NewValue:= .Value1)
                     Case OperationsFEFFC0_FEFFFFE.PUSH
                        Stack(Push:= .Value1)
                     Case Else
                        Return False
                  End Select
               End If

               Select Case DirectCast(CByte(Operation), OperationsFEFFC0_FEFFFFE)
                  Case OperationsFEFFC0_FEFFFFE.DEC,
                     OperationsFEFFC0_FEFFFFE.INC
                     SetNewValue(OperandPair)
               End Select
            End With
         Case Else
            Return False
      End Select

      Return True
   End Function

   'This procedure executes the specified opcode and returns whether or not it succeeded..
   Public Function ExecuteOpcode(Optional Opcode As OpcodesE = Nothing) As Boolean
      Dim CFStopValue As New Boolean
      Dim LowOctet As New Integer
      Dim NewValue As New Integer
      Dim Operand As New Integer
      Dim OperandPair As New OperandPairStr
      Dim Override As New SegmentRegistersE?
      Dim Value As New Integer

      If Opcode = Nothing Then Opcode = DirectCast(GetByteCSIP(), OpcodesE)

      Select Case Opcode
         Case OpcodesE.AAA, OpcodesE.AAS
            If (CInt(Registers(SubRegisters8BitE.AL)) And &HF%) > &H9% OrElse CBool(Registers(FlagRegistersE.AF)) Then
               If Opcode = OpcodesE.AAA Then
                  NewValue = &H106%
               ElseIf Opcode = OpcodesE.AAS Then
                  NewValue = -&H106%
               End If

               Registers(Registers16BitE.AX, NewValue:=CInt(Registers(Registers16BitE.AX)) + NewValue)
               Registers(FlagRegistersE.AF, NewValue:=True)
               Registers(FlagRegistersE.CF, NewValue:=True)
            Else
               Registers(FlagRegistersE.AF, NewValue:=False)
               Registers(FlagRegistersE.CF, NewValue:=False)
            End If
         Case OpcodesE.AAD
            Operand = GetByteCSIP()
            If Operand = &HA% Then
               NewValue = (CInt(Registers(SubRegisters8BitE.AL)) + (CInt(Registers(SubRegisters8BitE.AH)) * Operand)) And &HFF%
               AdjustFlags(CInt(Registers(Registers16BitE.AX)), NewValue)
               Registers(SubRegisters8BitE.AL, NewValue:=NewValue)
               Registers(SubRegisters8BitE.AH, NewValue:=&H0%)
               Registers(FlagRegistersE.AF, NewValue:=False)
            Else
               Return False
            End If
         Case OpcodesE.AAM
            Value = CInt(Registers(Registers16BitE.AX))
            NewValue = CInt(Registers(SubRegisters8BitE.AL))
            Operand = GetByteCSIP()
            If Operand = &HA% Then
               Registers(SubRegisters8BitE.AH, NewValue:=CInt(Floor(NewValue / Operand)))
               Registers(SubRegisters8BitE.AL, NewValue:=NewValue Mod Operand)
               Registers(FlagRegistersE.AF, NewValue:=False)
               AdjustFlags(Value, CInt(Registers(Registers16BitE.AX)), Is8Bit:=False)
            Else
               Return False
            End If
         Case OpcodesE.ADC_TGT_REG8 To OpcodesE.ADC_AX_WORD,
            OpcodesE.ADD_TGT_REG8 To OpcodesE.ADD_AX_WORD,
            OpcodesE.AND_TGT_REG8 To OpcodesE.AND_AX_WORD,
            OpcodesE.CMP_TGT_REG8 To OpcodesE.CMP_AX_WORD,
            OpcodesE.MOV_TGT_REG8 To OpcodesE.MOV_REG16_SRC,
            OpcodesE.OR_TGT_REG8 To OpcodesE.OR_AX_WORD,
            OpcodesE.SBB_TGT_REG8 To OpcodesE.SBB_AX_WORD,
            OpcodesE.SUB_TGT_REG8 To OpcodesE.SUB_AX_WORD,
            OpcodesE.XOR_TGT_REG8 To OpcodesE.XOR_AX_WORD

            Operand = GetByteCSIP()
            OperandPair = GetValues(GetOperandPair(Opcode, CByte(Operand)))
            With OperandPair
               Select Case Opcode
                  Case OpcodesE.ADC_TGT_REG8 To OpcodesE.ADC_AX_WORD
                     .NewValue = .Value1 + .Value2
                     If CBool(Registers(FlagRegistersE.CF)) Then .NewValue += &H1%
                     AdjustFlags(.Value1, .NewValue, .Is8Bit,,,, Subtraction:=False)
                  Case OpcodesE.ADD_TGT_REG8 To OpcodesE.ADD_AX_WORD
                     .NewValue = .Value1 + .Value2
                     AdjustFlags(.Value1, .NewValue, .Is8Bit,,,, Subtraction:=False)
                  Case OpcodesE.AND_TGT_REG8 To OpcodesE.AND_AX_WORD
                     .NewValue = .Value1 And .Value2
                     AdjustFlags(.Value1, .NewValue, .Is8Bit)
                  Case OpcodesE.CMP_TGT_REG8 To OpcodesE.CMP_AX_WORD, OpcodesE.SUB_TGT_REG8 To OpcodesE.SUB_AX_WORD
                     .NewValue = .Value1 - .Value2
                     AdjustFlags(.Value1, .NewValue, .Is8Bit)
                  Case OpcodesE.MOV_TGT_REG8 To OpcodesE.MOV_REG16_SRC
                     .NewValue = .Value2
                  Case OpcodesE.OR_TGT_REG8 To OpcodesE.OR_AX_WORD
                     .NewValue = .Value1 Or .Value2
                     AdjustFlags(.Value1, .NewValue, .Is8Bit)
                  Case OpcodesE.SBB_TGT_REG8 To OpcodesE.SBB_AX_WORD
                     .NewValue = .Value1 - .Value2
                     If CBool(Registers(FlagRegistersE.CF)) Then .NewValue -= &H1%
                     AdjustFlags(.Value1, .NewValue, .Is8Bit)
                  Case OpcodesE.SUB_TGT_REG8 To OpcodesE.SUB_AX_WORD
                     .NewValue = .Value1 - .Value2
                     AdjustFlags(.Value1, .NewValue, .Is8Bit)
                  Case OpcodesE.XOR_TGT_REG8 To OpcodesE.XOR_AX_WORD
                     .NewValue = .Value1 Xor .Value2
                     AdjustFlags(.Value1, .NewValue, .Is8Bit)
               End Select

               Select Case Opcode
                  Case OpcodesE.ADC_TGT_REG8 To OpcodesE.ADC_AX_WORD,
               OpcodesE.ADD_TGT_REG8 To OpcodesE.ADD_AX_WORD,
               OpcodesE.AND_TGT_REG8 To OpcodesE.AND_AX_WORD,
               OpcodesE.MOV_TGT_REG8 To OpcodesE.MOV_REG16_SRC,
               OpcodesE.OR_TGT_REG8 To OpcodesE.OR_AX_WORD,
               OpcodesE.SBB_TGT_REG8 To OpcodesE.SBB_AX_WORD,
               OpcodesE.SUB_TGT_REG8 To OpcodesE.SUB_AX_WORD,
               OpcodesE.XOR_TGT_REG8 To OpcodesE.XOR_AX_WORD

                     SetNewValue(OperandPair)
               End Select
            End With
         Case OpcodesE.CBW
            Registers(Registers16BitE.AX, ConvertWidening(CInt(Registers(SubRegisters8BitE.AL)), Is8Bit:=True))
         Case OpcodesE.CLC, OpcodesE.STC
            Registers(FlagRegistersE.CF, NewValue:=CBool(Opcode And &H1%))
         Case OpcodesE.CLD, OpcodesE.STD
            Registers(FlagRegistersE.DF, NewValue:=CBool(Opcode And &H1%))
         Case OpcodesE.CLI, OpcodesE.STI
            Registers(FlagRegistersE.IF, NewValue:=CBool(Opcode And &H1%))
         Case OpcodesE.CMC
            Registers(FlagRegistersE.CF, NewValue:=Not CBool(Registers(FlagRegistersE.CF)))
         Case OpcodesE.CS, OpcodesE.DS, OpcodesE.ES, OpcodesE.SS
            SegmentOverride(NewOverride:=DirectCast((CInt(Opcode) >> &H3%) And &H3%, SegmentRegistersE))
         Case OpcodesE.CWD
            NewValue = ConvertWidening(CInt(Registers(Registers16BitE.AX)), Is8Bit:=False)
            Registers(Registers16BitE.AX, NewValue:=NewValue)
            Registers(Registers16BitE.DX, NewValue:=NewValue >> &H8%)
         Case OpcodesE.DAA
            NewValue = CInt(Registers(SubRegisters8BitE.AL))
            LowOctet = NewValue And &HF%
            If (LowOctet > &H9% AndAlso LowOctet < &H10%) OrElse CBool(Registers(FlagRegistersE.AF)) Then
               NewValue += &H6%
               Registers(FlagRegistersE.AF, NewValue:=False)
            End If
            If NewValue > &H99% Then NewValue -= &HA0%
            AdjustFlags(CInt(Registers(SubRegisters8BitE.AL)), NewValue, Is8Bit:=True)
            Registers(SubRegisters8BitE.AL, NewValue:=NewValue)
         Case OpcodesE.DAS
            NewValue = CInt(Registers(SubRegisters8BitE.AL))
            If NewValue > &H99% Then NewValue -= &H60%
            LowOctet = NewValue And &HF%
            If (LowOctet > &H9% AndAlso LowOctet < &H10%) OrElse CBool(Registers(FlagRegistersE.AF)) Then
               NewValue -= &H6%
               Registers(FlagRegistersE.AF, NewValue:=False)
            End If
            AdjustFlags(CInt(Registers(SubRegisters8BitE.AL)), NewValue, Is8Bit:=True)
            Registers(SubRegisters8BitE.AL, NewValue:=NewValue)
         Case OpcodesE.DEC_AX To OpcodesE.DEC_DI, OpcodesE.INC_AX To OpcodesE.INC_DI
            Operand = Opcode And &H7%
            Value = CInt(Registers(DirectCast(Operand, Registers16BitE)))
            Select Case Opcode
               Case OpcodesE.DEC_AX To OpcodesE.DEC_DI
                  NewValue = Value - &H1%
               Case OpcodesE.INC_AX To OpcodesE.INC_DI
                  NewValue = Value + &H1%
            End Select
            Registers(DirectCast(Operand, Registers16BitE), NewValue:=NewValue)
            AdjustFlags(Value, NewValue, Is8Bit:=False,, PreserveCarryFlag:=True)
         Case OpcodesE.HLT
            RaiseEvent Halt()
         Case OpcodesE.IN_AL_BYTE
            RaiseEvent ReadIOPort(GetByteCSIP(), NewValue, Is8Bit:=True)
            Registers(SubRegisters8BitE.AL, NewValue:=NewValue)
         Case OpcodesE.IN_AL_DX
            RaiseEvent ReadIOPort(CInt(Registers(Registers16BitE.DX)), NewValue, Is8Bit:=True)
            Registers(SubRegisters8BitE.AL, NewValue:=NewValue)
         Case OpcodesE.IN_AX_DX
            RaiseEvent ReadIOPort(CInt(Registers(Registers16BitE.DX)), NewValue, Is8Bit:=False)
            Registers(Registers16BitE.AX, NewValue:=NewValue)
         Case OpcodesE.IN_AX_WORD
            RaiseEvent ReadIOPort(GetWordCSIP(), NewValue, Is8Bit:=False)
            Registers(Registers16BitE.AX, NewValue:=NewValue)
         Case OpcodesE.IRET
            Registers(Registers16BitE.IP, NewValue:=Stack())
            Registers(SegmentRegistersE.CS, NewValue:=Stack())
            Registers(FlagRegistersE.All, NewValue:=Stack())
         Case OpcodesE.INT, OpcodesE.INT3, OpcodesE.INTO
            ExecuteInterrupt(Opcode)
         Case OpcodesE.LAHF
            Registers(SubRegisters8BitE.AH, NewValue:=CInt(Registers(FlagRegistersE.All)) And LOW_FLAG_BITS)
         Case OpcodesE.LDS, OpcodesE.LES
            Operand = GetByteCSIP()
            If Operand < &HC0% Then
               OperandPair = GetValues(GetOperandPair(CByte(Opcode Xor &H6%), CByte(Operand)))

               With OperandPair
                  Registers(DirectCast(.Operand1, Registers16BitE), NewValue:=GET_WORD(CInt(.FlatAddress)))

                  Select Case Opcode
                     Case OpcodesE.LDS
                        Registers(SegmentRegistersE.DS, NewValue:=GET_WORD(CInt(.FlatAddress) + &H2%))
                     Case OpcodesE.LES
                        Registers(SegmentRegistersE.ES, NewValue:=GET_WORD(CInt(.FlatAddress) + &H2%))
                  End Select
               End With
            Else
               Return False
            End If
         Case OpcodesE.LEA
            Operand = GetByteCSIP()
            If Operand < &HC0% Then
               OperandPair = GetValues(GetOperandPair(CByte(Opcode Xor &H6%), CByte(Operand)))
               Registers(DirectCast(OperandPair.Operand1, Registers16BitE), NewValue:=OperandPair.Address)
            Else
               Return False
            End If
         Case OpcodesE.MOV_AL_BYTE To OpcodesE.MOV_BH_BYTE
            Registers(DirectCast(Opcode And &H7%, SubRegisters8BitE), NewValue:=GetByteCSIP())
         Case OpcodesE.MOV_AL_MEM
            Override = SegmentOverride()
            Registers(SubRegisters8BitE.AL, NewValue:=Memory((CInt(If(Override Is Nothing, Registers(SegmentRegistersE.DS), Registers(Override))) << &H4%) + GetWordCSIP()))
         Case OpcodesE.MOV_AX_MEM
            Override = SegmentOverride()
            Registers(Registers16BitE.AX, NewValue:=GET_WORD((CInt(If(Override Is Nothing, Registers(SegmentRegistersE.DS), Registers(Override))) << &H4%) + GetWordCSIP()))
         Case OpcodesE.MOV_AX_WORD To OpcodesE.MOV_DI_WORD
            Registers(DirectCast(Opcode And &H7%, Registers16BitE), NewValue:=GetWordCSIP())
         Case OpcodesE.MOV_MEM_AL
            Override = SegmentOverride()
            Memory((CInt(If(Override Is Nothing, Registers(SegmentRegistersE.DS), Registers(Override))) << &H4%) + GetWordCSIP()) = CByte(Registers(SubRegisters8BitE.AL))
         Case OpcodesE.MOV_MEM_AX
            Override = SegmentOverride()
            PutWord((CInt(If(Override Is Nothing, Registers(SegmentRegistersE.DS), Registers(Override))) << &H4%) + GetWordCSIP(), CInt(Registers(Registers16BitE.AX)))
         Case OpcodesE.MOV_SG_SRC
            With GetSegmentAndTarget(Opcode, GetByteCSIP())
               If TypeOf .Operand2 Is MemoryOperandsE Then NewValue = GET_WORD(CInt(AddressFromOperand(DirectCast(.Operand2, MemoryOperandsE), .Displacement, .DisplacementIs8Bit).FlatAddress)) Else NewValue = CInt(Registers(.Operand2))
               Registers(.Operand1, NewValue:=NewValue)
            End With
         Case OpcodesE.MOV_TGT_BYTE
            With GetOperandPair(Opcode, GetByteCSIP())
               If TypeOf .Operand1 Is MemoryOperandsE Then Memory(CInt(.FlatAddress)) = CByte(.Operand2) Else Registers(.Operand1, NewValue:= .Operand2)
            End With
         Case OpcodesE.MOV_TGT_SG
            With GetSegmentAndTarget(Opcode, GetByteCSIP())
               NewValue = CInt(Registers(.Operand1))
               If TypeOf .Operand2 Is MemoryOperandsE Then PutWord(CInt(AddressFromOperand(DirectCast(.Operand2, MemoryOperandsE), .Displacement, .DisplacementIs8Bit).FlatAddress), NewValue) Else Registers(.Operand2, NewValue:=NewValue)
            End With
         Case OpcodesE.MOV_TGT_WORD
            With GetOperandPair(Opcode, GetByteCSIP())
               If TypeOf .Operand1 Is MemoryOperandsE Then PutWord(CInt(.FlatAddress), CInt(.Operand2)) Else Registers(.Operand1, NewValue:= .Operand2)
            End With
         Case OpcodesE.OUT_BYTE_AL
            RaiseEvent WriteIOPort(GetByteCSIP(), Value:=CInt(Registers(SubRegisters8BitE.AL)), Is8Bit:=True)
         Case OpcodesE.OUT_DX_AL
            RaiseEvent WriteIOPort(Port:=CInt(Registers(Registers16BitE.DX)), Value:=CInt(Registers(SubRegisters8BitE.AL)), Is8Bit:=True)
         Case OpcodesE.OUT_DX_AX
            RaiseEvent WriteIOPort(Port:=CInt(Registers(Registers16BitE.DX)), Value:=CInt(Registers(Registers16BitE.AX)), Is8Bit:=False)
         Case OpcodesE.OUT_WORD_AX
            RaiseEvent WriteIOPort(GetByteCSIP(), Value:=CInt(Registers(Registers16BitE.AX)), Is8Bit:=False)
         Case OpcodesE.REPNE, OpcodesE.REPZ
            If Opcode = OpcodesE.REPNE Then
               CFStopValue = True
            ElseIf Opcode = OpcodesE.REPZ Then
               CFStopValue = False
            End If

            Registers(FlagRegistersE.ZF, NewValue:=(Not CFStopValue))

            Opcode = DirectCast(GetByteCSIP(), OpcodesE)
            Do Until CBool(Registers(FlagRegistersE.ZF)) = CFStopValue OrElse CInt(Registers(Registers16BitE.CX)) = &H0%
               ExecuteStringOpcode(Opcode)
               Registers(Registers16BitE.CX, NewValue:=CInt(Registers(Registers16BitE.CX)) - &H1%)
            Loop
         Case OpcodesE.RSBITS_BYTE_1 To OpcodesE.RSBITS_WORD_1, OpcodesE.RSBITS_BYTE_CL To OpcodesE.RSBITS_WORD_CL
            Operand = GetByteCSIP()
            OperandPair = GetOperandPair(CByte(Opcode And &H1%), CByte(Operand))
            With OperandPair
               If TypeOf .Operand1 Is MemoryOperandsE Then
                  .Value1 = If(.Is8Bit, Memory(CInt(.FlatAddress)), GET_WORD(CInt(.FlatAddress)))
               ElseIf TypeOf .Operand1 Is Registers16BitE OrElse TypeOf .Operand1 Is SubRegisters8BitE Then
                  .Value1 = CInt(Registers(.Operand1))
               End If

               Select Case Opcode
                  Case OpcodesE.RSBITS_BYTE_1, OpcodesE.RSBITS_WORD_1
                     .Value2 = &H1%
                  Case OpcodesE.RSBITS_BYTE_CL, OpcodesE.RSBITS_WORD_CL
                     .Value2 = CInt(Registers(SubRegisters8BitE.CL))
               End Select

               .NewValue = .Value1
               For Shift As Integer = &H1% To .Value2
                  Select Case DirectCast(CByte((Operand And &H3F%) >> &H3%), RotateAndShiftOperationsE)
                     Case RotateAndShiftOperationsE.None
                        Return False
                     Case RotateAndShiftOperationsE.RCL
                        .NewValue = BitRotateOrShift(.NewValue, .Is8Bit, Right:=False)
                     Case RotateAndShiftOperationsE.RCR
                        .NewValue = BitRotateOrShift(.NewValue, .Is8Bit)
                     Case RotateAndShiftOperationsE.ROL
                        .NewValue = BitRotateOrShift(.NewValue, .Is8Bit, Right:=False)
                     Case RotateAndShiftOperationsE.ROR
                        .NewValue = BitRotateOrShift(.NewValue, .Is8Bit,)
                     Case RotateAndShiftOperationsE.SAR
                        .NewValue = BitRotateOrShift(.NewValue, .Is8Bit,, Rotate:=False, ReplicateSign:=True)
                     Case RotateAndShiftOperationsE.SHL
                        .NewValue = BitRotateOrShift(.NewValue, .Is8Bit, Right:=False, Rotate:=False)
                     Case RotateAndShiftOperationsE.SHR
                        .NewValue = BitRotateOrShift(.NewValue, .Is8Bit,, Rotate:=False)
                  End Select
               Next Shift

               SetNewValue(OperandPair)
            End With
         Case OpcodesE.SAHF
            Registers(FlagRegistersE.All, NewValue:=(CInt(Registers(FlagRegistersE.All)) And &HFF00%) Or (CInt(Registers(SubRegisters8BitE.AH)) And LOW_FLAG_BITS))
         Case OpcodesE.TEST_AL_BYTE
            Value = CInt(Registers(SubRegisters8BitE.AL))
            AdjustFlags(Value, NewValue:=Value And GetByteCSIP(),, NewCarryFlag:=&H0%,, NewOverflowFlag:=&H0%)
         Case OpcodesE.TEST_AX_WORD
            Value = CInt(Registers(Registers16BitE.AX))
            AdjustFlags(Value, NewValue:=Value And GetWordCSIP(), Is8Bit:=False, NewCarryFlag:=&H0%,, NewOverflowFlag:=&H0%)
         Case OpcodesE.TEST_SRC_REG16
            Operand = GetByteCSIP()
            OperandPair = GetOperandPair(CByte(Opcode And &H1%), CByte(Operand))

            With OperandPair
               If TypeOf .Operand1 Is MemoryOperandsE Then .Value1 = GET_WORD(CInt(.FlatAddress)) Else .Value1 = CInt(Registers(.Operand1))
               .Value2 = CInt(Registers(.Operand2))
               AdjustFlags(.Value1, NewValue:= .Value1 And .Value2, Is8Bit:=False, NewCarryFlag:=&H0%,, NewOverflowFlag:=&H0%)
            End With
         Case OpcodesE.TEST_SRC_REG8
            Operand = GetByteCSIP()
            OperandPair = GetOperandPair(CByte(Opcode And &H1%), CByte(Operand))

            With OperandPair
               If TypeOf .Operand1 Is MemoryOperandsE Then .Value1 = Memory(CInt(.FlatAddress)) Else .Value1 = CInt(Registers(.Operand1))
               .Value2 = CInt(Registers(.Operand2))
               AdjustFlags(.Value1, NewValue:= .Value1 And .Value2,, NewCarryFlag:=&H0%,, NewOverflowFlag:=&H0%)
            End With
         Case OpcodesE.XCHG_AX_CX To OpcodesE.XCHG_AX_DI
            Value = CInt(Registers(Registers16BitE.AX))
            Operand = Opcode And &H7%
            Registers(Registers16BitE.AX, NewValue:=CInt(Registers(DirectCast(Operand, Registers16BitE))))
            Registers(DirectCast(Operand, Registers16BitE), NewValue:=Value)
         Case OpcodesE.XCHG_REG16, OpcodesE.XCHG_REG8
            Operand = GetByteCSIP()
            OperandPair = GetOperandPair(CByte(Opcode And &H1%), CByte(Operand))

            With OperandPair
               .Value1 = CInt(Registers(.Operand2))
               If TypeOf .Operand1 Is MemoryOperandsE Then
                  .Value2 = If(.Is8Bit, Memory(CInt(.FlatAddress)), GET_WORD(CInt(.FlatAddress)))
               ElseIf TypeOf .Operand1 Is Registers16BitE OrElse TypeOf .Operand1 Is SubRegisters8BitE Then
                  .Value2 = CInt(Registers(.Operand1))
               End If

               Registers(.Operand2, NewValue:= .Value2)
               .NewValue = .Value1

               SetNewValue(OperandPair)
            End With
         Case OpcodesE.XLAT
            Registers(SubRegisters8BitE.AL, NewValue:=Memory(CInt(Registers(Registers16BitE.BX)) + CInt(Registers(SubRegisters8BitE.AL))))
         Case OpcodesE.EXT_INT
            RaiseEvent Interrupt(GetByteCSIP(), CInt(Registers(SubRegisters8BitE.AH)))
         Case Else
            If Array.IndexOf({OpcodesE.LOCK, OpcodesE.NOP, OpcodesE.WAIT}, Opcode) < 0 Then
               Return ExecuteControlFlowOpcode(Opcode) OrElse ExecuteMultiOpcode(Opcode) OrElse ExecuteStackOpcode(Opcode) OrElse ExecuteStringOpcode(Opcode)
            End If
      End Select

      Return True
   End Function

   'This procedure executes a stack instruction and returns whether or not it succeeded..
   Private Function ExecuteStackOpcode(Opcode As OpcodesE) As Boolean
      Select Case Opcode
         Case OpcodesE.POP_AX To OpcodesE.POP_DI
            Registers(DirectCast(Opcode And &H7%, Registers16BitE), NewValue:=Stack())
         Case OpcodesE.POP_ES, OpcodesE.POP_SS, OpcodesE.POP_DS
            Registers(DirectCast((Opcode And &H18%) >> &H3%, SegmentRegistersE), NewValue:=Stack())
         Case OpcodesE.POP_TGT
            With GetOperandPair(Opcode, GetByteCSIP(), HasNoDisplacement:=True)
               If TypeOf .Operand1 Is MemoryOperandsE Then PutWord(CInt(.FlatAddress), Stack()) Else Registers(.Operand1, NewValue:=Stack())
            End With
         Case OpcodesE.POPF
            Registers(Register:=FlagRegistersE.All, NewValue:=Stack())
            If CBool(Registers(FlagRegistersE.TF)) Then
               If Not ExecuteOpcode() Then
                  ExecuteInterrupt(OpcodesE.INT, Number:=INVALID_OPCODE)
               End If
               ExecuteInterrupt(OpcodesE.INT, Number:=SINGLE_STEP)
               Registers(FlagRegistersE.TF, NewValue:=False)
            End If
         Case OpcodesE.PUSH_AX To OpcodesE.PUSH_DI
            Stack(Push:=CInt(Registers(DirectCast(Opcode And &H7%, Registers16BitE))))
         Case OpcodesE.PUSH_ES, OpcodesE.PUSH_CS, OpcodesE.PUSH_SS, OpcodesE.PUSH_DS
            Stack(Push:=CInt(Registers(DirectCast((Opcode And &H18%) >> &H3%, SegmentRegistersE))))
         Case OpcodesE.PUSHF
            Stack(Push:=CInt(Registers(Register:=FlagRegistersE.All)))
         Case Else
            Return False
      End Select

      Return True
   End Function

   'This procedure executes the specified string opcode and returns whether or not it succeeded.
   Private Function ExecuteStringOpcode(Opcode As OpcodesE) As Boolean
      Dim Is8Bit As Boolean = ((Opcode And &H1%) = &H0%)
      Dim NewValue As New Integer
      Dim SourceValue As New Integer
      Dim StepSize As Integer = If(Is8Bit, &H1%, &H2%)
      Dim TargetValue As New Integer

      If CBool(Registers(FlagRegistersE.DF)) Then StepSize = -StepSize

      Select Case Opcode
         Case OpcodesE.CMPSB
            SourceValue = Memory((CInt(Registers(SegmentRegistersE.DS)) << &H4%) + CInt(Registers(Registers16BitE.SI)))
            TargetValue = Memory((CInt(Registers(SegmentRegistersE.ES)) << &H4%) + CInt(Registers(Registers16BitE.DI)))
            NewValue = TargetValue - SourceValue
            AdjustFlags(TargetValue, NewValue)
         Case OpcodesE.CMPSW
            SourceValue = GET_WORD((CInt(Registers(SegmentRegistersE.DS)) << &H4%) + CInt(Registers(Registers16BitE.SI)))
            TargetValue = GET_WORD((CInt(Registers(SegmentRegistersE.ES)) << &H4%) + CInt(Registers(Registers16BitE.DI)))
            NewValue = TargetValue - SourceValue
            AdjustFlags(TargetValue, NewValue, Is8Bit:=False)
         Case OpcodesE.LODSB
            Registers(SubRegisters8BitE.AL, NewValue:=Memory((CInt(Registers(SegmentRegistersE.DS)) << &H4%) + CInt(Registers(Registers16BitE.SI))))
         Case OpcodesE.LODSW
            Registers(Registers16BitE.AX, NewValue:=GET_WORD((CInt(Registers(SegmentRegistersE.DS)) << &H4%) + CInt(Registers(Registers16BitE.SI))))
         Case OpcodesE.MOVSB
            Memory((CInt(Registers(SegmentRegistersE.ES)) << &H4%) + CInt(Registers(Registers16BitE.DI))) = Memory((CInt(Registers(SegmentRegistersE.DS)) << &H4%) + CInt(Registers(Registers16BitE.SI)))
         Case OpcodesE.MOVSW
            PutWord((CInt(Registers(SegmentRegistersE.ES)) << &H4%) + CInt(Registers(Registers16BitE.DI)), GET_WORD((CInt(Registers(SegmentRegistersE.DS)) << &H4%) + CInt(Registers(Registers16BitE.SI))))
         Case OpcodesE.SCASB
            SourceValue = CInt(Registers(SubRegisters8BitE.AL))
            TargetValue = Memory((CInt(Registers(SegmentRegistersE.ES)) << &H4%) + CInt(Registers(Registers16BitE.DI)))
            NewValue = TargetValue - SourceValue
            AdjustFlags(TargetValue, NewValue)
         Case OpcodesE.SCASW
            SourceValue = CInt(Registers(Registers16BitE.AX))
            TargetValue = GET_WORD((CInt(Registers(SegmentRegistersE.ES)) << &H4%) + CInt(Registers(Registers16BitE.DI)))
            NewValue = TargetValue - SourceValue
            AdjustFlags(TargetValue, NewValue, Is8Bit:=False)
         Case OpcodesE.STOSB
            Memory((CInt(Registers(SegmentRegistersE.ES)) << &H4%) + CInt(Registers(Registers16BitE.DI))) = CByte(Registers(SubRegisters8BitE.AL))
         Case OpcodesE.STOSW
            PutWord((CInt(Registers(SegmentRegistersE.ES)) << &H4%) + CInt(Registers(Registers16BitE.DI)), CInt(Registers(Registers16BitE.AX)))
         Case Else
            Return False
      End Select

      Select Case Opcode
         Case OpcodesE.CMPSB, OpcodesE.CMPSW, OpcodesE.MOVSB, OpcodesE.MOVSW
            Registers(Registers16BitE.DI, NewValue:=CInt(Registers(Registers16BitE.DI)) + StepSize)
            Registers(Registers16BitE.SI, NewValue:=CInt(Registers(Registers16BitE.SI)) + StepSize)
         Case OpcodesE.LODSB, OpcodesE.LODSW
            Registers(Registers16BitE.SI, NewValue:=CInt(Registers(Registers16BitE.SI)) + StepSize)
         Case OpcodesE.SCASB, OpcodesE.SCASW, OpcodesE.STOSB, OpcodesE.STOSW
            Registers(Registers16BitE.DI, NewValue:=CInt(Registers(Registers16BitE.DI)) + StepSize)
         Case Else
            Return False
      End Select

      Return True
   End Function

   'This procedure returns the byte located at CS:IP and adjusts the IP register.
   Private Function GetByteCSIP() As Byte
      Dim IP As Integer = CInt(Registers(Registers16BitE.IP))
      Dim [Byte] As Byte = Memory(((CInt(Registers(SegmentRegistersE.CS)) << &H4%) + IP) And ADDRESS_MASK)

      Registers(Registers16BitE.IP, NewValue:=(IP + &H1%) And &HFFFF%)

      Return [Byte]
   End Function

   'This procedure returns an operand's displacement value, if present.
   Private Function GetOperandDisplacement(Operand As Integer, ByRef DisplacementIs8Bit As Boolean) As Integer?
      Dim Displacement As New Integer

      Select Case Operand >> &H3%
         Case DisplacementTypesE.Byte
            DisplacementIs8Bit = True
            Displacement = GetByteCSIP()
         Case DisplacementTypesE.Word
            DisplacementIs8Bit = False
            Displacement = GetWordCSIP()
      End Select

      Return Displacement
   End Function

   'This procedure returns the operand pair indicated by the specified opcode and operand byte.
   Private Function GetOperandPair(Opcode As Byte, Operand As Byte, Optional IsMemReg16Byte As Boolean = False, Optional HasNoDisplacement As Boolean = False) As OperandPairStr
      Dim Addresses As New AddressesStr
      Dim MemoryOperand As MemoryOperandsE = DirectCast(((Operand And &HC0%) >> &H3%) + (Operand And &H7%), MemoryOperandsE)
      Dim OperandPair As New OperandPairStr With {.Is8Bit = Not CBool(Opcode And &H1%), .Displacement = Nothing, .Operand1 = Nothing, .Operand2 = Nothing}
      Dim OperandPairType As OperandPairsE = If(IsMemReg16Byte, OperandPairsE.MemReg_16_Byte, DirectCast(CByte(Opcode And &H7%), OperandPairsE))
      Dim Register1 As Integer = (Operand And &H38%) >> &H3%
      Dim Register2 As Integer = Operand And &H7%

      With OperandPair
         Select Case OperandPairType
            Case OperandPairsE.MemReg_Reg_8 To OperandPairsE.Reg_MemReg_16, OperandPairsE.MemReg_16_Byte, OperandPairsE.MemReg_Byte To OperandPairsE.MemReg_Word
               If MemoryOperand >= MemoryOperandsE.BX_SI_BYTE Then .Displacement = GetOperandDisplacement(MemoryOperand, .DisplacementIs8Bit)
         End Select

         Select Case OperandPairType
            Case OperandPairsE.MemReg_Reg_8
               If Operand < &HC0% Then .Operand1 = MemoryOperand Else .Operand1 = DirectCast(Register2, SubRegisters8BitE)
               .Operand2 = DirectCast(Register1, SubRegisters8BitE)
            Case OperandPairsE.MemReg_Reg_16
               If Operand < &HC0% Then .Operand1 = MemoryOperand Else .Operand1 = DirectCast(Register2, Registers16BitE)
               .Operand2 = DirectCast(Register1, Registers16BitE)
            Case OperandPairsE.Reg_MemReg_8
               .Operand1 = DirectCast(Register1, SubRegisters8BitE)
               If Operand < &HC0% Then .Operand2 = MemoryOperand Else .Operand2 = DirectCast(Register2, SubRegisters8BitE)
            Case OperandPairsE.Reg_MemReg_16
               .Operand1 = DirectCast(Register1, Registers16BitE)
               If Operand < &HC0% Then .Operand2 = MemoryOperand Else .Operand2 = DirectCast(Register2, Registers16BitE)
            Case OperandPairsE.AL_Byte
               .Operand1 = SubRegisters8BitE.AL
               .Operand2 = Operand
            Case OperandPairsE.AX_Word
               .Operand1 = Registers16BitE.AX
               .Operand2 = Operand Or (CInt(GetByteCSIP()) << &H8%)
            Case OperandPairsE.MemReg_Byte
               If Operand < &HC0% Then .Operand1 = MemoryOperand Else .Operand1 = DirectCast(Register2, SubRegisters8BitE)
            Case OperandPairsE.MemReg_Word
               If Operand < &HC0% Then .Operand1 = MemoryOperand Else .Operand1 = DirectCast(Register2, Registers16BitE)
            Case OperandPairsE.MemReg_16_Byte
               .Is8Bit = False
               If Operand < &HC0% Then .Operand1 = MemoryOperand Else .Operand1 = DirectCast(Register2, Registers16BitE)
         End Select

         If TypeOf .Operand1 Is MemoryOperandsE Then
            Addresses = AddressFromOperand(DirectCast(.Operand1, MemoryOperandsE), .Displacement, .DisplacementIs8Bit)
            .Address = CInt(Addresses.Address)
            .FlatAddress = CInt(Addresses.FlatAddress)
         ElseIf TypeOf .Operand2 Is MemoryOperandsE Then
            Addresses = AddressFromOperand(DirectCast(.Operand2, MemoryOperandsE), .Displacement, .DisplacementIs8Bit)
            .Address = CInt(Addresses.Address)
            .FlatAddress = CInt(Addresses.FlatAddress)
         End If

         If Not HasNoDisplacement Then
            Select Case OperandPairType
               Case OperandPairsE.MemReg_Byte, OperandPairsE.MemReg_16_Byte
                  .Operand2 = GetByteCSIP()
               Case OperandPairsE.MemReg_Word
                  .Operand2 = GetWordCSIP()
            End Select
         End If
      End With

      Return OperandPair
   End Function

   'This procedure returns the source segment register and target indicated by the specified opcode and operand.
   Private Function GetSegmentAndTarget(Opcode As Byte, Operand As Byte) As OperandPairStr
      Dim MemoryOperand As MemoryOperandsE = DirectCast(((Operand And &HC0%) >> &H3%) + (Operand And &H7%), MemoryOperandsE)
      Dim Register1 As Integer = (Operand And &H38%) >> &H3%
      Dim OperandPair As New OperandPairStr With {.Is8Bit = False, .Operand1 = DirectCast(Register1, SegmentRegistersE)}
      Dim Register2 As Integer = Operand And &H7%

      If Operand < &HC0% Then OperandPair.Operand2 = MemoryOperand Else OperandPair.Operand2 = DirectCast(Register2, Registers16BitE)

      Select Case DirectCast(OperandPair.Operand2, MemoryOperandsE)
         Case MemoryOperandsE.BX_SI_BYTE,
              MemoryOperandsE.BX_DI_BYTE,
              MemoryOperandsE.BP_SI_BYTE,
              MemoryOperandsE.BP_DI_BYTE,
              MemoryOperandsE.SI_BYTE,
              MemoryOperandsE.DI_BYTE,
              MemoryOperandsE.BP_BYTE,
              MemoryOperandsE.BX_BYTE,
              MemoryOperandsE.BX_SI_WORD,
              MemoryOperandsE.BX_DI_WORD,
              MemoryOperandsE.BP_SI_WORD,
              MemoryOperandsE.BP_DI_WORD,
              MemoryOperandsE.SI_WORD,
              MemoryOperandsE.DI_WORD,
              MemoryOperandsE.BP_WORD,
              MemoryOperandsE.BX_WORD
            OperandPair.Displacement = GetOperandDisplacement(MemoryOperand, OperandPair.DisplacementIs8Bit)
      End Select

      Return OperandPair
   End Function

   'This procedure returns the values from the specified locations.
   Private Function GetValues(OperandPair As OperandPairStr) As OperandPairStr
      With OperandPair
         If TypeOf .Operand1 Is MemoryOperandsE Then
            .Value1 = If(.Is8Bit, Memory(CInt(.FlatAddress)), GET_WORD(CInt(.FlatAddress)))
         ElseIf TypeOf .Operand1 Is Registers16BitE OrElse TypeOf .Operand1 Is SubRegisters8BitE Then
            .Value1 = CInt(Registers(.Operand1))
         End If

         If TypeOf .Operand2 Is MemoryOperandsE Then
            .Value2 = If(.Is8Bit, Memory(CInt(.FlatAddress)), GET_WORD(CInt(.FlatAddress)))
         ElseIf TypeOf .Operand2 Is Registers16BitE OrElse TypeOf .Operand2 Is SubRegisters8BitE Then
            .Value2 = CInt(Registers(.Operand2))
         Else
            .Value2 = CInt(.Operand2)
         End If
      End With

      Return OperandPair
   End Function

   'This procedure returns the word located at CS:IP and adjusts the IP register.
   Private Function GetWordCSIP() As Integer
      Dim IP As Integer = CInt(Registers(Registers16BitE.IP))
      Dim Word As Integer = GET_WORD((CInt(Registers(SegmentRegistersE.CS)) << &H4%) + IP)

      Registers(Registers16BitE.IP, NewValue:=(IP + &H2%) And &HFFFF%)

      Return Word
   End Function

   'This procedure writes the specified word to the specified address.
   Public Sub PutWord(Address As Integer, Word As Integer)
      Memory(Address And ADDRESS_MASK) = CByte(Word And &HFF%)
      Memory((Address + &H1%) And ADDRESS_MASK) = CByte((Word And &HFF00%) >> &H8%)
   End Sub

   'This procedure sets and/or returns the specified register's value.
   Public Function Registers(Register As Object, Optional NewValue As Object = Nothing) As Object
      Dim Index As New Integer
      Dim Value As New Object
      Static FlagsRegister As Integer = &H0%
      Static Registers16Bit(Registers16BitE.AX To Registers16BitE.IP) As Integer
      Static SegmentRegisters(SegmentRegistersE.ES To SegmentRegistersE.DS) As Integer

      If TypeOf Register Is FlagRegistersE Then
         If DirectCast(Register, FlagRegistersE) = FlagRegistersE.All Then
            If NewValue IsNot Nothing Then FlagsRegister = CInt(NewValue)
            FlagsRegister = FlagsRegister Or (&H1% << DirectCast(FlagRegistersE.F1, Integer))
            Value = FlagsRegister
         Else
            Index = DirectCast(Register, Integer)

            If NewValue IsNot Nothing Then
               If CBool(NewValue) Then
                  FlagsRegister = FlagsRegister Or (&H1% << Index)
               Else
                  FlagsRegister = FlagsRegister Xor (FlagsRegister And (&H1% << Index))
               End If
            End If

            FlagsRegister = FlagsRegister Or (&H1% << DirectCast(FlagRegistersE.F1, Integer))

            Value = CBool(FlagsRegister And (&H1% << Index))
         End If
      ElseIf TypeOf Register Is Registers16BitE Then
         Index = DirectCast(Register, Integer)
         If NewValue IsNot Nothing Then Registers16Bit(Index) = CInt(NewValue) And &HFFFF%
         Value = Registers16Bit(Index) And &HFFFF%
      ElseIf TypeOf Register Is SegmentRegistersE Then
         Index = DirectCast(Register, Integer)
         If NewValue IsNot Nothing Then SegmentRegisters(Index) = CInt(NewValue) And &HFFFF%
         Value = SegmentRegisters(Index) And &HFFFF%
      ElseIf TypeOf Register Is SubRegisters8BitE Then
         Index = DirectCast(Register, Integer) And &H3%
         If NewValue IsNot Nothing Then Registers16Bit(Index) = If(DirectCast(Register, SubRegisters8BitE) >= SubRegisters8BitE.AH, (Registers16Bit(Index) And &HFF%) Or ((CInt(NewValue) And &HFF%) << &H8%), (Registers16Bit(Index) And &HFF00%) Or (CInt(NewValue) And &HFF%))
         Value = If(DirectCast(Register, SubRegisters8BitE) >= SubRegisters8BitE.AH, (Registers16Bit(Index) And &HFF00%) >> &H8%, Registers16Bit(Index) And &HFF%)
      End If

      Return Value
   End Function

   'This procedure sets and/or returns the current segment override.
   Public Function SegmentOverride(Optional NewOverride As SegmentRegistersE? = Nothing, Optional Preserve As Boolean = False) As SegmentRegistersE?
      Dim Override As SegmentRegistersE? = Nothing
      Static CurrentOverride As SegmentRegistersE? = Nothing

      If NewOverride IsNot Nothing Then
         CurrentOverride = NewOverride
         Override = CurrentOverride
      Else
         Override = CurrentOverride
         If Not Preserve Then CurrentOverride = Nothing
      End If

      Return Override
   End Function

   'This procedure sets the specified new value at the location specified by the specified operands.
   Private Sub SetNewValue(OperandPair As OperandPairStr)
      With OperandPair
         If TypeOf .Operand1 Is MemoryOperandsE Then
            If .Is8Bit Then Memory(CInt(.FlatAddress)) = CByte(.NewValue And &HFF%) Else PutWord(CInt(.FlatAddress), .NewValue)
         ElseIf TypeOf .Operand1 Is Registers16BitE OrElse TypeOf .Operand1 Is SubRegisters8BitE Then
            Registers(.Operand1, NewValue:= .NewValue)
         End If
      End With
   End Sub

   'This procedure manages the stack and returns any value popped.
   Public Function Stack(Optional Push As Integer? = Nothing) As Integer
      Dim SP As Integer = CInt(Registers(Registers16BitE.SP))
      Dim Word As New Integer

      If Push IsNot Nothing Then
         SP = (SP - &H2%) And &HFFFF%
         PutWord((CInt(Registers(SegmentRegistersE.SS)) << &H4%) + SP, Push.Value)
         Registers(Registers16BitE.SP, NewValue:=SP)
      Else
         Word = GET_WORD((CInt(Registers(SegmentRegistersE.SS)) << &H4%) + SP)
         SP = (SP + &H2%) And &HFFFF%
         Registers(Registers16BitE.SP, NewValue:=SP)

         Return Word
      End If

      Return Nothing
   End Function
End Class
