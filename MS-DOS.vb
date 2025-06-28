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
Imports System.Linq
Imports System.Math
Imports System.Windows.Forms

'This module handles MS-DOS related functions.
Public Module MSDOSModule
   'This enumeration lists the STD file handles.
   Private Enum STDFileHandlesE As Integer
      STDIN    'Input.
      STDOUT   'Output.
      STDERR   'Error.
      STDAUX   'Auxiliary.
      STDPRN   'Printer.
   End Enum

   Private Const CARRY_FLAG_INDEX As Integer = &H0%                     'Defines the carry flag's bit index.
   Private Const ERROR_INSUFFICIENT_MEMORY As Integer = &H8%            'Defines the insufficient memory error code.
   Private Const ERROR_INVALID_MEMORY_BLOCK_ADDRESS As Integer = &H9%   'Defines the invalid memory block address error code.
   Private Const EXE_HEADER_SIZE As Integer = &H8%                      'Defines where an executable's header size is stored.
   Private Const EXE_INITIAL_CS As Integer = &H16%                      'Defines where an executable's initial code segment is stored.
   Private Const EXE_INITIAL_IP As Integer = &H14%                      'Defines where an executable's initial instruction pointer is stored.
   Private Const EXE_INITIAL_SP As Integer = &H10%                      'Defines where an executable's initial stack pointer is stored.
   Private Const EXE_INITIAL_SS As Integer = &HE%                       'Defines where an executable's initial stack segment is stored.
   Private Const EXE_RELOCATION_ITEM_COUNT As Integer = &H6%            'Defines where an executable's number of relocation items is stored.
   Private Const EXE_RELOCATION_ITEM_TABLE As Integer = &H18%           'Defines where an executable's number of relocation item table offset is stored.
   Private Const HIGHEST_ADDRESS As Integer = &HA000%                   'Defines the highest address that can be allocated.
   Private Const LOWEST_ADDRESS As Integer = &H600%                     'Defines the lowest address that can be allocated.
   Private Const MS_DOS As Integer = &HFF00%                            'Defines a value indicating that the operating system is MS-DOS.
   Private Const PSP_ENVIRONMENT_SEGMENT As Integer = &H2C%             'Defines the segment of the MS-DOS environment in a PSP.
   Private Const PSP_MEMORY_TOP As Integer = &H2%                       'Defines the offset of the top of memory value in a PSP.
   Private Const PSP_SIZE As Integer = &H100%                           'Defines a PSP's size.
   Private Const VERSION As Integer = &H1606%                           'Defines the emulated MS-DOS version as 6.22.

   Private ReadOnly ENVIRONMENT As String = $"COMSPEC=C:\COMMAND.COM{ToChar(&H0%)}PATH={ToChar(&H0%)}"   'Defines the MS-DOS environment.
   Private ReadOnly ENVIRONMENT_SEGMENT As Integer = LOWEST_ADDRESS                                      'Defines the MS-DOS environment's segment.
   Private ReadOnly EXE_MZ_SIGNATURE() As Byte = {&H4D%, &H5A%}                                          'Defines the signature of an MZ exectuable.

   Private Allocations As New List(Of Tuple(Of Integer, Integer))   'Contains the memory allocations.
   Private ProcessSegments As New Stack(Of Integer)                 'Contains the segments allocated to processes.

   'This procedure attempts to allocate the specified amount of memory and returns an address if successful.
   Private Function AllocateMemory(Size As Integer) As Integer?
      Try
         Dim AllocatedAddress As New Integer?
         Dim PreviousEndAddress As Integer = LOWEST_ADDRESS

         Allocations.Sort(Function(a, b) a.Item1.CompareTo(b.Item1))

         If Allocations.Count = 0 Then
            AllocatedAddress = LOWEST_ADDRESS
         Else
            For Each Allocation As Tuple(Of Integer, Integer) In Allocations
               If Allocation.Item1 - PreviousEndAddress >= Size Then
                  AllocatedAddress = PreviousEndAddress
                  Exit For
               End If

               PreviousEndAddress = Allocation.Item2 + &H1%
            Next Allocation

            If AllocatedAddress Is Nothing AndAlso HIGHEST_ADDRESS - PreviousEndAddress >= Size Then
               AllocatedAddress = PreviousEndAddress
            End If
         End If

         If AllocatedAddress IsNot Nothing Then
            If AllocatedAddress + Size <= HIGHEST_ADDRESS Then
               Allocations.Add(Tuple.Create(CInt(AllocatedAddress), CInt(AllocatedAddress) + Size - &H1%))
            Else
               AllocatedAddress = New Integer?
            End If
         End If

         Return AllocatedAddress
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure attempts to free the specified allocated memory address and returns whether or not it succeeded.
   Private Function FreeAllocatedMemory(StartAddress As Integer) As Boolean
      Try
         Dim Index As Integer = Allocations.FindIndex(Function(Allocation) Allocation.Item1 = StartAddress)
         Dim Found As Boolean = (Index >= 0)

         If Found Then Allocations.RemoveAt(Index)

         Return Found
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns the current time.
   Private Sub GetCurrentTime()
      Try
         Dim Counter As Integer = (CPU.GET_WORD(AddressesE.Clock + &H2%) << &H10%) Or (CPU.GET_WORD(AddressesE.Clock))
         Dim Hour As New Integer
         Dim Hundreth As New Integer
         Dim Minute As New Integer
         Dim Second As New Integer
         Dim TotalSeconds As New Double

         TotalSeconds = Counter / 18.2065
         Hour = CInt(Floor(TotalSeconds / 3600))
         TotalSeconds -= Hour * 3600
         Minute = CInt(Floor(TotalSeconds / 60))
         TotalSeconds -= Minute * 60
         Second = CInt(Floor(TotalSeconds))
         Hundreth = CInt(Floor((TotalSeconds - Second) * 100))

         CPU.Registers(CPU8086Class.SubRegisters8BitE.CH, NewValue:=Hour)
         CPU.Registers(CPU8086Class.SubRegisters8BitE.CL, NewValue:=Minute)
         CPU.Registers(CPU8086Class.SubRegisters8BitE.DH, NewValue:=Second)
         CPU.Registers(CPU8086Class.SubRegisters8BitE.DL, NewValue:=Hundreth)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure handles the specified MS-DOS interrupt and returns whether or not is succeeded.
   Public Function HandleMSDOSInterrupt(Number As Integer, AH As Integer, ByRef Flags As Integer) As Boolean
      Try
         Dim Address As New Integer
         Dim Count As New Integer
         Dim KeyCode As New Integer
         Dim Position As New Integer
         Dim Result As New Integer?
         Dim Success As Boolean = False
         Static ExtendedKeyCode As New Integer

         Select Case Number
            Case &H20%
               TerminateProgram($"Program terminated.{NewLine}")
               Success = True
            Case &H21%
               Select Case AH
                  Case &H0%
                     TerminateProgram($"Program terminated.{NewLine}")
                     Success = True
                  Case &H8%
                     If ExtendedKeyCode = Nothing Then

                        Do
                           Application.DoEvents()
                           KeyCode = LastBIOSKeyCode()
                           If (KeyCode And &HFF%) = &H0% Then ExtendedKeyCode = (KeyCode >> &H8%)
                           KeyCode = (KeyCode And &HFF%)
                        Loop While (KeyCode = Nothing) AndAlso (ExtendedKeyCode = Nothing) AndAlso (Not CPU.ClockToken.IsCancellationRequested)

                        LastBIOSKeyCode(, Clear:=True)

                        CPU.Registers(CPU8086Class.SubRegisters8BitE.AL, NewValue:=KeyCode)
                     Else
                        CPU.Registers(CPU8086Class.SubRegisters8BitE.AL, NewValue:=ExtendedKeyCode)
                        ExtendedKeyCode = New Integer
                     End If

                     Success = True
                  Case &H9%
                     Position = (CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H4%) + CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX))
                     Do Until ToChar(CPU.Memory(Position And CPU8086Class.ADDRESS_MASK)) = "$"c
                        TeleType(CPU.Memory(Position And CPU8086Class.ADDRESS_MASK))
                        Position += &H1%
                     Loop

                     Success = True
                  Case &H25%
                     Address = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)) * &H4%
                     CPU.PutWord(Address + &H2%, CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)))
                     CPU.PutWord(Address, CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX)))
                     Success = True
                     ''--->>>
                  Case &H2C%
                     GetCurrentTime()
                  Case &H30%
                     CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=VERSION)
                     CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=MS_DOS)
                     Success = True
                  Case &H35%
                     Address = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)) * &H4%
                     CPU.Registers(CPU8086Class.SegmentRegistersE.ES, NewValue:=CPU.GET_WORD(Address + &H2%))
                     CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=CPU.GET_WORD(Address))
                     Success = True
                  Case &H40%
                     Select Case DirectCast(CPU.Registers(CPU8086Class.Registers16BitE.BX), STDFileHandlesE)
                        Case STDFileHandlesE.STDOUT, STDFileHandlesE.STDERR
                           Count = CInt(CPU.Registers(CPU8086Class.Registers16BitE.CX))
                           Position = (CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H4%) + CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX))
                           For Character As Integer = &H0% To Count - &H1%
                              TeleType(CPU.Memory(Position And CPU8086Class.ADDRESS_MASK))
                              Position += &H1%
                           Next Character

                           CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=Count)

                           Success = True
                     End Select
                  Case &H48%
                     Result = AllocateMemory(CInt(CPU.Registers(CPU8086Class.Registers16BitE.BX)) << &H4%)
                     CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=(LargestFreeMemoryBlock() >> &H4%))
                     Flags = SetBit(Flags, (Result Is Nothing), CARRY_FLAG_INDEX)

                     If Result Is Nothing Then
                        CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=ERROR_INSUFFICIENT_MEMORY)
                     Else
                        CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=CInt(Result) >> &H4%)
                     End If
                     Success = True
                  Case &H49%
                     If FreeAllocatedMemory(CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.ES)) << &H4%) Then
                        CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=&H0%)
                        Flags = SetBit(Flags, False, CARRY_FLAG_INDEX)
                     Else
                        CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=ERROR_INVALID_MEMORY_BLOCK_ADDRESS)
                        Flags = SetBit(Flags, True, CARRY_FLAG_INDEX)
                     End If
                     Success = True
                  Case &H4A%
                     Result = ModifyAllocatedMemory(CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.ES)) << &H4%, CInt(CPU.Registers(CPU8086Class.Registers16BitE.BX)) << &H4%)
                     CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=(LargestFreeMemoryBlock() >> &H4%))
                     Flags = SetBit(Flags, (Result IsNot Nothing), CARRY_FLAG_INDEX)

                     If Result IsNot Nothing Then
                        CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=CInt(Result))
                     End If
                     Success = True
                  Case &H4C%
                     TerminateProgram($"Program terminated with return code: {CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)):X2}.{NewLine}")
                     Success = True
               End Select
         End Select

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function

   ' This function returns the size of the largest amount of memory that can be allocated.
   Private Function LargestFreeMemoryBlock() As Integer
      Try
         Dim FreeBlockSize As New Integer
         Dim LargestBlock As Integer = &H0%
         Dim LastFreeBlock As New Integer
         Dim PreviousEndAddress As Integer = LOWEST_ADDRESS

         Allocations.Sort(Function(a, b) a.Item1.CompareTo(b.Item1))

         If Allocations.Count = 0 Then
            LargestBlock = HIGHEST_ADDRESS - LOWEST_ADDRESS
         Else
            For Each Allocation As Tuple(Of Integer, Integer) In Allocations
               FreeBlockSize = Allocation.Item1 - PreviousEndAddress
               If FreeBlockSize > LargestBlock Then
                  LargestBlock = FreeBlockSize
               End If

               PreviousEndAddress = Allocation.Item2 + 1
            Next Allocation

            LastFreeBlock = HIGHEST_ADDRESS - PreviousEndAddress

            If LastFreeBlock > LargestBlock Then
               LargestBlock = LastFreeBlock
            End If
         End If

         Return LargestBlock
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure loads "MS-DOS" into memory.
   Public Sub LoadMSDOS()
      Try
         WriteStringToMemory(ENVIRONMENT, ENVIRONMENT_SEGMENT << &H4%)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure attempts to modify the specified allocated memory address and returns whether or not it succeeded.
   Private Function ModifyAllocatedMemory(StartAddress As Integer, NewSize As Integer) As Integer?
      Try
         Dim ErrorCode As New Integer?
         Dim Index As Integer = Allocations.FindIndex(Function(Allocation) Allocation.Item1 = StartAddress)
         Dim ModifiedAllocation As Tuple(Of Integer, Integer) = Nothing
         Dim NewEnd As New Integer

         If Index >= 0 Then
            ModifiedAllocation = Allocations(Index)
            NewEnd = ModifiedAllocation.Item1 + NewSize

            If NewEnd <= HIGHEST_ADDRESS OrElse Index = Allocations.Count - 1 OrElse NewEnd < Allocations(Index + 1).Item1 Then
               Allocations(Index) = Tuple.Create(ModifiedAllocation.Item1, NewEnd)
            Else
               ErrorCode = ERROR_INSUFFICIENT_MEMORY
            End If
         Else
            ErrorCode = ERROR_INVALID_MEMORY_BLOCK_ADDRESS
         End If

         Return ErrorCode
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure loads the specified MS-DOS program into the emulated CPU's memory after processing its header.
   Public Sub LoadMSDOSProgram(FileName As String)
      Try
         Dim Executable As New List(Of Byte)(File.ReadAllBytes(FileName))
         Dim LoadAddress As Integer = CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.CS)) << &H4%

         If Executable.Count >= &H2% AndAlso Executable.GetRange(&H0%, EXE_MZ_SIGNATURE.Length).SequenceEqual(EXE_MZ_SIGNATURE) Then
            Executable = LoadMZEXE(Executable, FileName, LoadAddress)
         Else
            Output.AppendText($"Loading the compact binary executable ""{FileName}"" at address {LoadAddress:X8}.{NewLine}")

            CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=&HFFFF%)
            CPU.Registers(CPU8086Class.Registers16BitE.CX, NewValue:=Executable.Count)
            CPU.Registers(CPU8086Class.Registers16BitE.DX, NewValue:=&H0%)
            CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=&H0%)
            CPU.Registers(CPU8086Class.SegmentRegistersE.DS, NewValue:=CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.CS)))
            CPU.Registers(CPU8086Class.SegmentRegistersE.ES, NewValue:=CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)))
            CPU.Registers(CPU8086Class.Registers16BitE.IP, NewValue:=PSP_SIZE)
            CPU.Registers(CPU8086Class.Registers16BitE.BP, NewValue:=&H0%)
            CPU.Registers(CPU8086Class.Registers16BitE.SP, NewValue:=&HFFFF%)
            CPU.Registers(CPU8086Class.SegmentRegistersE.SS, NewValue:=CPU.Registers(CPU8086Class.SegmentRegistersE.CS))

            CPU.PutWord(LoadAddress + PSP_ENVIRONMENT_SEGMENT, ENVIRONMENT_SEGMENT)
            CPU.PutWord(LoadAddress + PSP_MEMORY_TOP, LargestFreeMemoryBlock())

            Executable.CopyTo(CPU.Memory, LoadAddress + PSP_SIZE)
         End If

         If Allocations.FindIndex(Function(Allocation) Allocation.Item1 = CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H4%) < 0 Then
            Allocations.Add(Tuple.Create(CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H4%, ((Executable.Count >> &H4%) + &H1%) << &H4%))
            ProcessSegments.Push(Allocations.Last.Item1)
         Else
            SyncLock Synchronizer
               CPUEvent.Append($"Memory allocation failure.{NewLine}")
            End SyncLock
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure loads the specified MZ executable.
   Private Function LoadMZEXE(Executable As List(Of Byte), FileName As String, LoadAddress As Integer) As List(Of Byte)
      Try
         Dim HeaderSize As Integer = BitConverter.ToUInt16(Executable.ToArray(), EXE_HEADER_SIZE) << &H4%
         Dim InitialSP As Integer = BitConverter.ToUInt16(Executable.ToArray(), EXE_INITIAL_SP)
         Dim InitialSS As Integer = BitConverter.ToUInt16(Executable.ToArray(), EXE_INITIAL_SS)
         Dim Position As New Integer
         Dim RelocatedCS As Integer = (LoadAddress >> &H4%) + BitConverter.ToUInt16(Executable.ToArray(), EXE_INITIAL_CS)
         Dim RelocationItem As New Integer
         Dim RelocationItemFlatAddress As New Integer
         Dim RelocationItemOffset As New Integer
         Dim RelocationItemSegment As New Integer
         Dim RelocationTable As Integer = BitConverter.ToUInt16(Executable.ToArray(), EXE_RELOCATION_ITEM_TABLE)
         Dim RelocationTableSize As Integer = BitConverter.ToUInt16(Executable.ToArray(), EXE_RELOCATION_ITEM_COUNT) * &H4%

         If LoadAddress + Executable.Count <= CPU.Memory.Length Then
            Output.AppendText($"Loading the MZ-executable ""{FileName}"" at address {LoadAddress:X8}.{NewLine}")

            CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=&H0%)
            CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=(Executable.Count - HeaderSize) >> &H10%)
            CPU.Registers(CPU8086Class.Registers16BitE.CX, NewValue:=(Executable.Count - HeaderSize) And &HFFFF%)
            CPU.Registers(CPU8086Class.Registers16BitE.DX, NewValue:=&H0%)
            CPU.Registers(CPU8086Class.SegmentRegistersE.CS, NewValue:=RelocatedCS)
            CPU.Registers(CPU8086Class.Registers16BitE.IP, NewValue:=BitConverter.ToUInt16(Executable.ToArray(), EXE_INITIAL_IP))
            CPU.Registers(CPU8086Class.SegmentRegistersE.DS, NewValue:=((LoadAddress - PSP_SIZE) >> &H4%))
            CPU.Registers(CPU8086Class.SegmentRegistersE.ES, NewValue:=CPU.Registers(CPU8086Class.SegmentRegistersE.DS))
            CPU.Registers(CPU8086Class.SegmentRegistersE.SS, NewValue:=(If(InitialSS = Nothing, RelocatedCS, RelocatedCS + InitialSS)))
            CPU.Registers(CPU8086Class.Registers16BitE.SP, NewValue:=If(InitialSP = Nothing, &HFFFF%, InitialSP))

            If RelocationTableSize > &H0% Then
               Position = RelocationTable - &H1%
               Do
                  RelocationItemOffset = BitConverter.ToUInt16({Executable(Position + &H1%), Executable(Position)}, &H0%)
                  RelocationItemSegment = BitConverter.ToUInt16({Executable(Position + &H3%), Executable(Position + &H2%)}, &H0%)

                  RelocationItemFlatAddress = HeaderSize + ((RelocationItemSegment << &H4%) + RelocationItemOffset)

                  RelocationItem = (BitConverter.ToUInt16(Executable.ToArray(), RelocationItemFlatAddress) + (LoadAddress >> &H4%)) And &HFFFF%

                  Executable(RelocationItemFlatAddress) = ToByte(RelocationItem And &HFF%)
                  Executable(RelocationItemFlatAddress + &H1%) = ToByte(RelocationItem >> &H8%)

                  Position += &H4%
               Loop Until Position >= (RelocationTable + RelocationTableSize) - &H1%
            End If

            CPU.PutWord((LoadAddress - PSP_SIZE) + PSP_ENVIRONMENT_SEGMENT, ENVIRONMENT_SEGMENT)
            CPU.PutWord((LoadAddress - PSP_SIZE) + PSP_MEMORY_TOP, LargestFreeMemoryBlock())

            Executable = Executable.GetRange(HeaderSize, Executable.Count - HeaderSize)

            Executable.CopyTo(CPU.Memory, LoadAddress)
         Else
            Output.AppendText($"""{FileName}"" does not fit inside the emulated memory.{NewLine}")
         End If

         Return Executable
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return New List(Of Byte)
   End Function

   'This procedure resets the MS-DOS environment.
   Public Sub ResetMSDOS()
      Try
         Allocations.Clear()
         ProcessSegments.Clear()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure terminates the currently running MS-DOS program.
   Private Sub TerminateProgram(Message As String)
      Try
         CPU.ClockToken.Cancel()

         FreeAllocatedMemory(ProcessSegments.Pop())

         SyncLock Synchronizer
            CPUEvent.Append(Message)
         End SyncLock
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Module
