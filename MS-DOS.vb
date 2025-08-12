'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.DateTime
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Environment
Imports System.IO
Imports System.Linq
Imports System.Math
Imports System.Security
Imports System.Text
Imports System.Windows.Forms

'This module handles MS-DOS related functions.
Public Module MSDOSModule
   'This enumeration lists addresses used inside a DTA block.
   Private Enum DTAE As Integer
      Attribute = &H15%   'File attribute.
      FileTime = &H16%    'File time.
      FileDate = &H18%    'File date.
      FileSize = &H1A%    'File size.
      FileName = &H1E%    'Null terminated file name.
   End Enum

   'This enumeration lists the STD file handles.
   Private Enum STDFileHandlesE As Integer
      STDIN    'Input.
      STDOUT   'Output.
      STDERR   'Error.
      STDAUX   'Auxiliary.
      STDPRN   'Printer.
   End Enum

   Public Const COMMAND_TAIL_MAXIMUM_LENGTH As Integer = &H7E%          'Defines the maximum length of a command tail in a PSP.
   Private Const CARRY_FLAG_INDEX As Integer = &H0%                     'Defines the carry flag's bit index.
   Private Const ERROR_ACCESS_DENIED As Integer = &H5%                  'Defines the access denied error code.
   Private Const ERROR_CRC As Integer = &H17%                           'Defines the CRC error code.
   Private Const ERROR_DISK_FULL As Integer = &H70%                     'Defines the disk full error code.
   Private Const ERROR_FILE_NOT_FOUND As Integer = &H2%                 'Defines the file not found error code.
   Private Const ERROR_FILENAME_EXCED_RANGE As Integer = &HC3%          'Defines the filename exceeds range code.
   Private Const ERROR_GEN_FAILURE As Integer = &H1F%                   'Defines the general failure error code.
   Private Const ERROR_HANDLE_EOF As Integer = &H26%                    'Defines the handle EOF error code.
   Private Const ERROR_INVALID_HANDLE As Integer = &H6%                 'Defines the invalid handle error code.
   Private Const ERROR_INVALID_MEMORY_BLOCK_ADDRESS As Integer = &H9%   'Defines the invalid memory block address error code.
   Private Const ERROR_INSUFFICIENT_MEMORY As Integer = &H8%            'Defines the insufficient memory error code.
   Private Const ERROR_INVALID_NAME As Integer = &H7B%                  'Defines the invalid name error code.
   Private Const ERROR_LOCK_VIOLATION As Integer = &H21%                'Defines the lock violation error code.
   Private Const ERROR_NOT_READY As Integer = &HF%                      'Defines the not ready error code.
   Private Const ERROR_NOT_SUPPORTED As Integer = &H32%                 'Defines the not supported error code.
   Private Const ERROR_PATH_NOT_FOUND As Integer = &H3%                 'Defines the path not found error code.
   Private Const ERROR_SHARING_VIOLATION As Integer = &H20%             'Defines the sharing violation error code.
   Private Const EXE_HEADER_SIZE As Integer = &H8%                      'Defines where an executable's header size is stored.
   Private Const EXE_INITIAL_CS As Integer = &H16%                      'Defines where an executable's initial code segment is stored.
   Private Const EXE_INITIAL_IP As Integer = &H14%                      'Defines where an executable's initial instruction pointer is stored.
   Private Const EXE_INITIAL_SP As Integer = &H10%                      'Defines where an executable's initial stack pointer is stored.
   Private Const EXE_INITIAL_SS As Integer = &HE%                       'Defines where an executable's initial stack segment is stored.
   Private Const EXE_RELOCATION_ITEM_COUNT As Integer = &H6%            'Defines where an executable's number of relocation items is stored.
   Private Const EXE_RELOCATION_ITEM_TABLE As Integer = &H18%           'Defines where an executable's number of relocation item table offset is stored.
   Private Const FILE_ACCESS_RW_MASK As Integer = &H3%                  'Defines the read/write bits for file access.
   Private Const HIGHEST_ADDRESS As Integer = &HA0000%                  'Defines the highest address that can be allocated.
   Private Const INT_20H As Integer = &H20CD%                           'Defines the INT 20h instruction.
   Private Const LOWEST_ADDRESS As Integer = &H600%                     'Defines the lowest address that can be allocated.
   Private Const MS_DOS As Integer = &HFF00%                            'Defines a value indicating that the operating system is MS-DOS.
   Private Const PSP_BYTES_AVAILABLE As Integer = &H6%                  'Defines the number of bytes available in a block of memory less the space used by the PSP.
   Private Const PSP_COMMAND_TAIL As Integer = &H80%                    'Defines the offset of the command tail in a PSP.
   Private Const PSP_ENVIRONMENT_SEGMENT As Integer = &H2C%             'Defines the segment of the MS-DOS environment in a PSP.
   Private Const PSP_INT_20H As Integer = &H0%                          'Defines a call to the INT 20h handler in a PSP.
   Private Const PSP_INT_21H As Integer = &H50%                         'Defines a call to the INT 21h handler in a PSP.
   Private Const PSP_INT_22H As Integer = &HA%                          'Defines the offset of the INT 22h handler in a PSP.
   Private Const PSP_INT_23H As Integer = &HE%                          'Defines the offset of the INT 23h handler in a PSP.
   Private Const PSP_INT_24H As Integer = &H12%                         'Defines the offset of the INT 24h handler in a PSP.
   Private Const PSP_MEMORY_TOP As Integer = &H2%                       'Defines the offset of the top of memory value in a PSP.
   Private Const PSP_PREVIOUS_PSP As Integer = &H38%                    'Defines the offset of the previous PSP in a PSP.
   Private Const PSP_SIZE As Integer = &H100%                           'Defines a PSP's size.
   Private Const PSP_SSSP As Integer = &H2E%                            'Defines the SS:SP values in a PSP.
   Private Const VERSION As Integer = &H1606%                           'Defines the emulated MS-DOS version as 6.22.

   Private ReadOnly DATE_TO_MSDOS_DATE As Func(Of Date, Integer) = Function([Date] As Date) ([Date].Day And &H1F%) Or (([Date].Month And &HF) << &H5%) Or ((If([Date].Year - 1980 >= &H0% AndAlso [Date].Year - 1980 < &H7F%, [Date].Year - 1980, Nothing) And &H7F%) << &H9%)   'Converts the specified date to a value suitable for MS-DOS and returns the result.
   Private ReadOnly ENVIRONMENT_SEGMENT As Integer = LOWEST_ADDRESS                                           'Defines the MS-DOS environment's segment.
   Private ReadOnly ENVIRONMENT_TEXT As String = $"COMSPEC=C:\COMMAND.COM{ToChar(&H0%)}PATH={ToChar(&H0%)}"   'Defines the MS-DOS environment.
   Private ReadOnly EXE_MZ_SIGNATURE() As Byte = {&H4D%, &H5A%}                                               'Defines the signature of an MZ executable.
   Private ReadOnly INT_21H_RETF() As Byte = {&HCD%, &H21%, &HCB%}                                            'Defines the INT 21h and RETF instructions.
   Private ReadOnly LOWEST_FILE_HANDLE As Integer = STDFileHandlesE.STDPRN + &H1%                             'Defines the lowest possible file handle.
   Private ReadOnly TIME_TO_MSDOS_TIME As Func(Of Date, Integer) = Function([Date] As Date) (([Date].Second \ &H2%) And &H1F%) Or (([Date].Minute And &H3F) << &H5%) Or (([Date].Hour And &H1F) << &HB%)   'Converts the specified time to a value suitable for MS-DOS and returns the result.

   Public CommandTail As String = ""                                   'Contains the command tail used in a new PSP.
   Private Allocations As New List(Of Tuple(Of Integer, Integer))      'Contains the memory allocations.
   Private AvailableDevices As Boolean = True                          'Contains the AVAILDEV flag.
   Private DTA As New Integer                                          'Contains the Disk Transfer Address.
   Private OpenFiles As New List(Of Tuple(Of FileStream, Integer))     'Contains the open file streams and their handles.
   Private PrinterBuffer As New StringBuilder                          'Contains the printer buffer.
   Private ProcessSegments As New Stack(Of Integer)                    'Contains the segments allocated to processes.
   Private SwitchCharacter As Char = "-"c                              'Contains the switch character.

   Private WithEvents PrinterDocumentO As New PrintDocument   'Contains the document with output to STDPRN to be printed.

   'This procedure attempts to allocate the specified amount of memory and returns an address if successful.
   Private Function AllocateMemory(Size As Integer) As Integer?
      Try
         Dim AllocatedAddress As New Integer?
         Dim PreviousEndAddress As Integer = LOWEST_ADDRESS

         Allocations.Sort(Function(Allocation1, Allocation2) Allocation1.Item1.CompareTo(Allocation2.Item1))

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

   'This procedure attempts to close the specified file handle and returns whether or not it succeeded.
   Private Function CloseFileHandle(ByRef Flags As Integer) As Boolean
      Try
         Dim OpenFileToBeClosed As Tuple(Of FileStream, Integer) = OpenFiles.FirstOrDefault(Function(OpenedFile) OpenedFile.Item2 = CInt(CPU.Registers(CPU8086Class.Registers16BitE.BX)))
         Dim Success As Boolean = False

         If OpenFileToBeClosed IsNot Nothing Then
            Try
               OpenFileToBeClosed.Item1.Close()
               Flags = SET_BIT(Flags, False, CARRY_FLAG_INDEX)
            Catch MSDOSException As Exception
               CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=GetMSDOSErrorCode(MSDOSException))
               Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
            End Try

            OpenFiles.Remove(OpenFileToBeClosed)
            Success = True
         End If

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function

   'This procedure creates a file.
   Private Sub CreateFile(ByRef Flags As Integer)
      Try
         Dim FileName As String = GetStringZ(CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)), CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX)))
         Dim FileStreamO As FileStream = Nothing
         Dim NextHandle As New Integer?

         Try
            FileStreamO = New FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Write)
            File.SetAttributes(FileName, DirectCast(CPU.Registers(CPU8086Class.Registers16BitE.CX), FileAttributes))
            NextHandle = GetNextFreeFileHandle()
            Flags = SET_BIT(Flags, False, CARRY_FLAG_INDEX)
         Catch MSDOSException As Exception
            CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=GetMSDOSErrorCode(MSDOSException))
            Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
         End Try

         If NextHandle IsNot Nothing Then
            CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=CInt(NextHandle))
            OpenFiles.Add(New Tuple(Of FileStream, Integer)(FileStreamO, CInt(NextHandle)))
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure creates a PSP.
   Private Sub CreatePSP(Address As Integer)
      Try
         CPU.PutWord(Address + PSP_INT_20H, INT_20H)
         CPU.PutWord(Address + PSP_MEMORY_TOP, LargestFreeMemoryBlock())
         CPU.PutWord(Address + PSP_BYTES_AVAILABLE, &HFEF0%)
         CPU.PutWord(Address + PSP_INT_22H, CPU.GET_WORD(&H8A%))
         CPU.PutWord(Address + PSP_INT_22H + &H2%, CPU.GET_WORD(&H88%))
         CPU.PutWord(Address + PSP_INT_23H, CPU.GET_WORD(&H8E%))
         CPU.PutWord(Address + PSP_INT_23H + &H2%, CPU.GET_WORD(&H8C%))
         CPU.PutWord(Address + PSP_INT_24H, CPU.GET_WORD(&H91%))
         CPU.PutWord(Address + PSP_INT_24H + &H2%, CPU.GET_WORD(&H90%))
         CPU.PutWord(Address + PSP_SSSP, CInt(CPU.Registers(CPU8086Class.Registers16BitE.SP)))
         CPU.PutWord(Address + PSP_SSSP + &H2%, CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.SS)))
         CPU.PutWord(Address + PSP_ENVIRONMENT_SEGMENT, ENVIRONMENT_SEGMENT)
         CPU.PutWord(Address + PSP_PREVIOUS_PSP, &HFFFF%)
         CPU.PutWord(Address + PSP_PREVIOUS_PSP + &H2%, &HFFFF%)
         WriteBytesToMemory(INT_21H_RETF, Address + PSP_INT_21H)
         CPU.Memory(Address + PSP_COMMAND_TAIL) = ToByte(CommandTail.Length)
         WriteStringToMemory($"{CommandTail}{ToChar(&HD%)}", Address + PSP_COMMAND_TAIL + &H1%)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub


   'This procedure deletes a file.
   Private Sub DeleteFile(ByRef Flags As Integer)
      Try
         Try
            File.Delete(GetStringZ(CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)), CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX))))
            Flags = SET_BIT(Flags, False, CARRY_FLAG_INDEX)
         Catch MSDOSException As Exception
            CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=GetMSDOSErrorCode(MSDOSException))
            Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
         End Try
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure attempts to find files matching a given pattern.
   Private Sub FindFile(ByRef Flags As Integer, Optional IsFirst As Boolean = False)
      Try
         Dim Attributes As New FileAttributes
         Dim FileName As String = Nothing
         Static Files As New Stack(Of String)

         Try
            If IsFirst Then
               FileName = GetStringZ(CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)), CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX)))
               Attributes = DirectCast(CPU.Registers(CPU8086Class.Registers16BitE.CX), FileAttributes)
               Files = New Stack(Of String)(From Item In Directory.GetFiles(CurrentDirectory(), FileName) Where (File.GetAttributes(Item) And Attributes) = Attributes Select Path.GetFileName(Item))
               WriteDTA(Files.Pop)
            Else
               If Files.Count = 0 Then
                  CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=ERROR_FILE_NOT_FOUND)
                  Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
               Else
                  WriteDTA(Files.Pop)
               End If
            End If
         Catch MSDOSException As Exception
            CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=GetMSDOSErrorCode(MSDOSException))
            Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
         End Try
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

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
         Dim Hundredth As New Integer
         Dim Minute As New Integer
         Dim Second As New Integer
         Dim TotalSeconds As Double = Counter / 18.2065

         Hour = CInt(Floor(TotalSeconds / 3600))
         TotalSeconds -= Hour * 3600
         Minute = CInt(Floor(TotalSeconds / 60))
         TotalSeconds -= Minute * 60
         Second = CInt(Floor(TotalSeconds))
         Hundredth = CInt(Floor((TotalSeconds - Second) * 100))

         CPU.Registers(CPU8086Class.SubRegisters8BitE.CH, NewValue:=Hour)
         CPU.Registers(CPU8086Class.SubRegisters8BitE.CL, NewValue:=Minute)
         CPU.Registers(CPU8086Class.SubRegisters8BitE.DH, NewValue:=Second)
         CPU.Registers(CPU8086Class.SubRegisters8BitE.DL, NewValue:=Hundredth)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure returns a file's date and time.
   Private Sub GetFileDateTime(ByRef Flags As Integer)
      Try
         Dim OpenFileToBeChecked As Tuple(Of FileStream, Integer) = Nothing

         Select Case DirectCast(CPU.Registers(CPU8086Class.Registers16BitE.BX), STDFileHandlesE)
            Case STDFileHandlesE.STDAUX, STDFileHandlesE.STDERR, STDFileHandlesE.STDIN, STDFileHandlesE.STDOUT, STDFileHandlesE.STDPRN
               CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=ERROR_INVALID_HANDLE)
               Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
            Case Else
               OpenFileToBeChecked = OpenFiles.FirstOrDefault(Function(OpenedFile) OpenedFile.Item2 = CInt(CPU.Registers(CPU8086Class.Registers16BitE.BX)))

               Try
                  CPU.Registers(CPU8086Class.Registers16BitE.CX, NewValue:=TIME_TO_MSDOS_TIME(File.GetCreationTime(OpenFileToBeChecked.Item1.Name)))
                  CPU.Registers(CPU8086Class.Registers16BitE.DX, NewValue:=DATE_TO_MSDOS_DATE(File.GetCreationTime(OpenFileToBeChecked.Item1.Name)))
                  Flags = SET_BIT(Flags, False, CARRY_FLAG_INDEX)
               Catch MSDOSException As Exception
                  CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=GetMSDOSErrorCode(MSDOSException))
                  Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
               End Try
         End Select
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure returns the MS-DOS error code for the specified exception.
   Private Function GetMSDOSErrorCode(MSDOSException As Exception) As Integer
      Try
         Dim IOExceptionO As New IOException
         Dim MSDOSErrorCode As Integer = ERROR_GEN_FAILURE

         Select Case True
            Case TypeOf MSDOSException Is DirectoryNotFoundException
               MSDOSErrorCode = ERROR_PATH_NOT_FOUND
            Case TypeOf MSDOSException Is DriveNotFoundException
               MSDOSErrorCode = ERROR_NOT_READY
            Case TypeOf MSDOSException Is EndOfStreamException
               MSDOSErrorCode = ERROR_HANDLE_EOF
            Case TypeOf MSDOSException Is FileNotFoundException
               MSDOSErrorCode = ERROR_FILE_NOT_FOUND
            Case TypeOf MSDOSException Is IOException
               IOExceptionO = DirectCast(MSDOSException, IOException)

               Select Case IOExceptionO.HResult And &HFFFF%
                  Case ERROR_DISK_FULL, ERROR_INVALID_NAME, ERROR_LOCK_VIOLATION, ERROR_SHARING_VIOLATION, ERROR_CRC
                     MSDOSErrorCode = (IOExceptionO.HResult And &HFFFF%)
                  Case Else
                     MSDOSErrorCode = ERROR_GEN_FAILURE
               End Select
            Case TypeOf MSDOSException Is NotSupportedException
               MSDOSErrorCode = ERROR_NOT_SUPPORTED
            Case TypeOf MSDOSException Is ObjectDisposedException
               MSDOSErrorCode = ERROR_GEN_FAILURE
            Case TypeOf MSDOSException Is PathTooLongException
               MSDOSErrorCode = ERROR_FILENAME_EXCED_RANGE
            Case TypeOf MSDOSException Is UnauthorizedAccessException
               MSDOSErrorCode = ERROR_ACCESS_DENIED
            Case TypeOf MSDOSException Is SecurityException
               MSDOSErrorCode = ERROR_ACCESS_DENIED
            Case Else
               MSDOSErrorCode = ERROR_GEN_FAILURE
         End Select

         Return MSDOSErrorCode
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure returns the next free file handle.
   Private Function GetNextFreeFileHandle() As Integer
      Try
         Dim NextHandle As Integer = LOWEST_FILE_HANDLE
         Dim UsedHandles As HashSet(Of Integer) = OpenFiles.Select(Function(Handle) Handle.Item2).ToHashSet()

         While UsedHandles.Contains(NextHandle)
            NextHandle += 1
         End While

         Return NextHandle
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure reads a key with echo and returns the result.
   Private Function GetKeyWithEcho() As Integer
      Try
         Dim KeyCode As New Integer

         Do
            Application.DoEvents()
            KeyCode = LastBIOSKeyCode() And &HFF%
         Loop While (KeyCode = Nothing) AndAlso (Not CPU.ClockToken.IsCancellationRequested)

         LastBIOSKeyCode(, Clear:=True)
         TeleType(CByte(KeyCode))
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure handles the specified MS-DOS interrupt and returns whether or not it succeeded.
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
            Case &H20%, &H22%
               TerminateProgram($"Program terminated.{NewLine}")
               Success = True
            Case &H21%
               Select Case AH
                  Case &H0%
                     TerminateProgram($"Program terminated.{NewLine}")
                     Success = True
                  Case &H1%
                     CPU.Registers(CPU8086Class.SubRegisters8BitE.AL, NewValue:=GetKeyWithEcho())
                     Success = True
                  Case &H2%
                     TeleType(CByte(CPU.Registers(CPU8086Class.SubRegisters8BitE.DL)))
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
                  Case &H1A%
                     DTA = (CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H10%) Or CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX))
                     Success = True
                  Case &H25%
                     Address = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)) * &H4%
                     CPU.PutWord(Address + &H2%, CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)))
                     CPU.PutWord(Address, CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX)))
                     Success = True
                  Case &H2A%
                     CPU.Registers(CPU8086Class.SubRegisters8BitE.AL, NewValue:=Now.DayOfWeek)
                     CPU.Registers(CPU8086Class.Registers16BitE.CX, NewValue:=Now.Year)
                     CPU.Registers(CPU8086Class.SubRegisters8BitE.DH, NewValue:=Now.Month)
                     CPU.Registers(CPU8086Class.SubRegisters8BitE.DL, NewValue:=Now.Day)
                     Success = True
                  Case &H2B%
                     Success = True
                  Case &H2C%
                     GetCurrentTime()
                     Success = True
                  Case &H2F%
                     CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=DTA And &HFFFF%)
                     CPU.Registers(CPU8086Class.SegmentRegistersE.ES, NewValue:=DTA >> &H10%)
                     Success = True
                  Case &H30%
                     CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=VERSION)
                     CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=MS_DOS)
                     Success = True
                  Case &H35%
                     Address = CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)) * &H4%
                     CPU.Registers(CPU8086Class.SegmentRegistersE.ES, NewValue:=CPU.GET_WORD(Address + &H2%))
                     CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=CPU.GET_WORD(Address))
                     Success = True
                  Case &H37%
                     Select Case CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL))
                        Case &H0%
                           CPU.Registers(CPU8086Class.SubRegisters8BitE.DL, NewValue:=ToByte(SwitchCharacter))
                        Case &H1%
                           SwitchCharacter = ToChar(CPU.Registers(CPU8086Class.SubRegisters8BitE.DL))
                        Case &H2%
                           CPU.Registers(CPU8086Class.SubRegisters8BitE.DL, NewValue:=Abs(CInt(AvailableDevices)))
                        Case &H3%
                           AvailableDevices = CBool(CPU.Registers(CPU8086Class.SubRegisters8BitE.DL))
                        Case Else
                           CPU.Registers(CPU8086Class.SubRegisters8BitE.AL, NewValue:=&HFF%)
                     End Select
                     Success = True
                  Case &H3C%
                     CreateFile(Flags)
                     Success = True
                  Case &H3D%
                     OpenFile(Flags)
                     Success = True
                  Case &H3E%
                     CloseFileHandle(Flags)
                     Success = True
                  Case &H3F%
                     ReadFile(Flags)
                     Success = True
                  Case &H40%
                     WriteFile(Flags)
                     Success = True
                  Case &H41%
                     DeleteFile(Flags)
                     Success = True
                  Case &H42%
                     SeekFile(Flags)
                     Success = True
                  Case &H44%
                     Select Case CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL))
                        Case &H0%
                           Select Case DirectCast(CPU.Registers(CPU8086Class.Registers16BitE.BX), STDFileHandlesE)
                              Case STDFileHandlesE.STDAUX, STDFileHandlesE.STDERR, STDFileHandlesE.STDIN, STDFileHandlesE.STDOUT, STDFileHandlesE.STDPRN
                                 CPU.Registers(CPU8086Class.Registers16BitE.DX, NewValue:=&H80%)
                                 Success = True
                           End Select
                     End Select
                  Case &H48%
                     Result = AllocateMemory(CInt(CPU.Registers(CPU8086Class.Registers16BitE.BX)) << &H4%)
                     CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=(LargestFreeMemoryBlock() >> &H4%))
                     Flags = SET_BIT(Flags, (Result Is Nothing), CARRY_FLAG_INDEX)

                     If Result Is Nothing Then
                        CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=ERROR_INSUFFICIENT_MEMORY)
                     Else
                        CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=CInt(Result) >> &H4%)
                     End If

                     Success = True
                  Case &H49%
                     If FreeAllocatedMemory(CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.ES)) << &H4%) Then
                        CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=&H0%)
                        Flags = SET_BIT(Flags, False, CARRY_FLAG_INDEX)
                     Else
                        CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=ERROR_INVALID_MEMORY_BLOCK_ADDRESS)
                        Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
                     End If
                     Success = True
                  Case &H4A%
                     Result = ModifyAllocatedMemory(CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.ES)) << &H4%, CInt(CPU.Registers(CPU8086Class.Registers16BitE.BX)) << &H4%)
                     CPU.Registers(CPU8086Class.Registers16BitE.BX, NewValue:=(LargestFreeMemoryBlock() >> &H4%))
                     Flags = SET_BIT(Flags, (Result IsNot Nothing), CARRY_FLAG_INDEX)

                     If Result IsNot Nothing Then
                        CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=CInt(Result))
                     End If
                     Success = True
                  Case &H4C%
                     TerminateProgram($"Program terminated with return code: {CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)):X2}.{NewLine}")
                     Success = True
                  Case &H4E%
                     FindFile(Flags, IsFirst:=True)
                     Success = True
                  Case &H4F%
                     FindFile(Flags)
                     Success = True
                  Case &H57%
                     Select Case CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL))
                        Case &H0%
                           GetFileDateTime(Flags)
                           Success = True
                     End Select
               End Select
            Case &H23%
               TerminateProgram($"CTRL+Break.{NewLine}")
               Success = True
            Case &H24%
               TerminateProgram($"INT 24h - Critical error.{NewLine}")
               Success = True
            Case &H27%
               TerminateProgram($"Terminate and stay resident.{NewLine}")
               Success = True
         End Select

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function

   'This function returns the size of the largest amount of memory that can be allocated.
   Private Function LargestFreeMemoryBlock() As Integer
      Try
         Dim FreeBlockSize As New Integer
         Dim LargestBlock As Integer = &H0%
         Dim LastFreeBlock As New Integer
         Dim PreviousEndAddress As Integer = LOWEST_ADDRESS

         Allocations.Sort(Function(Allocation1, Allocation2) Allocation1.Item1.CompareTo(Allocation2.Item1))

         If Allocations.Count = 0 Then
            LargestBlock = HIGHEST_ADDRESS - LOWEST_ADDRESS
         Else
            For Each Allocation As Tuple(Of Integer, Integer) In Allocations
               FreeBlockSize = Allocation.Item1 - PreviousEndAddress
               If FreeBlockSize > LargestBlock Then
                  LargestBlock = FreeBlockSize
               End If

               PreviousEndAddress = Allocation.Item2 + &H1%
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
         WriteStringToMemory(ENVIRONMENT_TEXT, ENVIRONMENT_SEGMENT << &H4%)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure loads the specified MS-DOS program into the emulated CPU's memory after processing its header and returns whether or not it succeeded.
   Public Function LoadMSDOSProgram(FileName As String) As Boolean
      Try
         Dim Executable As New List(Of Byte)(File.ReadAllBytes(FileName))
         Dim LoadAddress As Integer = CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.CS)) << &H4%
         Dim Success As Boolean = True

         If Executable.Count >= &H2% AndAlso Executable.GetRange(&H0%, EXE_MZ_SIGNATURE.Length).SequenceEqual(EXE_MZ_SIGNATURE) Then
            Executable = LoadMZEXE(Executable, FileName, LoadAddress)
            Success = Executable.Any
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
            CPU.Registers(CPU8086Class.SegmentRegistersE.SS, NewValue:=CPU.Registers(CPU8086Class.SegmentRegistersE.CS))
            CPU.Registers(CPU8086Class.Registers16BitE.SP, NewValue:=&HFFFC%)
            CPU.PutWord((CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.SS)) << &H4%) + CInt(CPU.Registers(CPU8086Class.Registers16BitE.SP)), CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.CS)))
            CPU.PutWord((CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.SS)) << &H4%) + CInt(CPU.Registers(CPU8086Class.Registers16BitE.SP)) + &H2%, &H0%)

            CreatePSP(LoadAddress)

            Executable.CopyTo(CPU.Memory, LoadAddress + PSP_SIZE)
         End If

         If Allocations.FindIndex(Function(Allocation) Allocation.Item1 = CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H4%) < 0 Then
            Allocations.Add(Tuple.Create(CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H4%, ((Executable.Count >> &H4%) + &H1%) << &H4%))
            ProcessSegments.Push(Allocations.Last.Item1)
         Else
            SyncLock Synchronizer
               CPUEvent.Append($"Memory allocation failure.{NewLine}")
            End SyncLock

            Success = False
         End If

         Return Success
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try

      Return False
   End Function

   'This procedure loads the specified MZ executable and returns the result.
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

                  If RelocationItemFlatAddress < Executable.Count Then
                     RelocationItem = BitConverter.ToUInt16(Executable.ToArray(), RelocationItemFlatAddress)
                     RelocationItem = (BitConverter.ToUInt16(Executable.ToArray(), RelocationItemFlatAddress) + (LoadAddress >> &H4%)) And &HFFFF%

                     Executable(RelocationItemFlatAddress) = ToByte(RelocationItem And &HFF%)
                     Executable(RelocationItemFlatAddress + &H1%) = ToByte(RelocationItem >> &H8%)
                  End If

                  Position += &H4%
               Loop Until Position >= (RelocationTable + RelocationTableSize) - &H1%
            End If

            CreatePSP(LoadAddress - PSP_SIZE)

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

   'This procedure opens a file.
   Private Sub OpenFile(ByRef Flags As Integer)
      Try
         Dim FileAccessO As New FileAccess
         Dim FileName As String = GetStringZ(CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)), CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX)))
         Dim FileStreamO As FileStream = Nothing
         Dim NextHandle As New Integer?

         Select Case CInt(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL)) And FILE_ACCESS_RW_MASK
            Case &H0%
               FileAccessO = FileAccess.Read
            Case &H1%
               FileAccessO = FileAccess.Write
            Case &H2%
               FileAccessO = FileAccess.ReadWrite
         End Select

         Try
            FileStreamO = New FileStream(FileName, FileMode.Open, FileAccessO)
            NextHandle = GetNextFreeFileHandle()
            Flags = SET_BIT(Flags, False, CARRY_FLAG_INDEX)
         Catch MSDOSException As Exception
            CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=GetMSDOSErrorCode(MSDOSException))
            Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
         End Try

         If NextHandle IsNot Nothing Then
            CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=CInt(NextHandle))
            OpenFiles.Add(New Tuple(Of FileStream, Integer)(FileStreamO, CInt(NextHandle)))
         End If
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure sends the printer buffer's content to the default printer.
   Private Sub PrinterDocumentO_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrinterDocumentO.PrintPage
      Try
         e.Graphics.DrawString(PrinterBuffer.ToString(), New Font("Consolas", 10), Brushes.Black, New Point(8, 8))
         PrinterBuffer.Clear()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure reads a file.
   Private Sub ReadFile(ByRef Flags As Integer)
      Try
         Dim Bytes() As Byte = {}
         Dim Count As Integer = CInt(CPU.Registers(CPU8086Class.Registers16BitE.CX))
         Dim OpenFileToBeRead As Tuple(Of FileStream, Integer) = Nothing

         ReDim Bytes(&H0% To Count - &H1%)

         Select Case DirectCast(CPU.Registers(CPU8086Class.Registers16BitE.BX), STDFileHandlesE)
            Case STDFileHandlesE.STDAUX, STDFileHandlesE.STDIN
               For Character As Integer = &H0% To Count - &H1%
                  Bytes(Character) = ToByte(GetKeyWithEcho())
               Next Character
            Case STDFileHandlesE.STDERR, STDFileHandlesE.STDOUT, STDFileHandlesE.STDPRN
               CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=ERROR_ACCESS_DENIED)
               Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
            Case Else
               OpenFileToBeRead = OpenFiles.FirstOrDefault(Function(OpenedFile) OpenedFile.Item2 = CInt(CPU.Registers(CPU8086Class.Registers16BitE.BX)))

               Try
                  Count = OpenFileToBeRead.Item1.Read(Bytes, offset:=&H0%, Count)
                  CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=Count)
                  ReDim Preserve Bytes(&H0% To Count - &H1%)
                  WriteBytesToMemory(Bytes, (CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H4%) Or CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX)))
                  Flags = SET_BIT(Flags, False, CARRY_FLAG_INDEX)
               Catch MSDOSException As Exception
                  CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=GetMSDOSErrorCode(MSDOSException))
                  Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
               End Try
         End Select
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure resets the MS-DOS environment.
   Public Sub ResetMSDOS()
      Try
         Allocations.Clear()
         AvailableDevices = True
         CommandTail = ""
         DTA = New Integer
         OpenFiles.Clear()
         PrinterBuffer.Clear()
         ProcessSegments.Clear()
         SwitchCharacter = "-"c
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure seeks inside a file.
   Private Sub SeekFile(ByRef Flags As Integer)
      Try
         Dim OpenFileToBeSought As Tuple(Of FileStream, Integer) = Nothing

         Select Case DirectCast(CPU.Registers(CPU8086Class.Registers16BitE.BX), STDFileHandlesE)
            Case STDFileHandlesE.STDAUX, STDFileHandlesE.STDERR, STDFileHandlesE.STDIN, STDFileHandlesE.STDOUT, STDFileHandlesE.STDPRN
               CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=ERROR_ACCESS_DENIED)
               Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
            Case Else
               OpenFileToBeSought = OpenFiles.FirstOrDefault(Function(OpenedFile) OpenedFile.Item2 = CInt(CPU.Registers(CPU8086Class.Registers16BitE.BX)))
               Try
                  OpenFileToBeSought.Item1.Seek((CInt(CPU.Registers(CPU8086Class.Registers16BitE.CX)) << &H10%) Or CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX)), DirectCast(CPU.Registers(CPU8086Class.SubRegisters8BitE.AL), SeekOrigin))
                  CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=OpenFileToBeSought.Item1.Position And &HFFFF%)
                  CPU.Registers(CPU8086Class.Registers16BitE.DX, NewValue:=OpenFileToBeSought.Item1.Position >> &H10%)
                  Flags = SET_BIT(Flags, False, CARRY_FLAG_INDEX)
               Catch MSDOSException As Exception
                  CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=GetMSDOSErrorCode(MSDOSException))
                  Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
               End Try
         End Select
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure terminates the currently running MS-DOS program.
   Private Sub TerminateProgram(Message As String, Optional IsResident As Boolean = False)
      Try
         CPU.ClockToken.Cancel()

         If Not IsResident Then FreeAllocatedMemory(ProcessSegments.Pop())

         SyncLock Synchronizer
            CPUEvent.Append(Message)
         End SyncLock
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure writes information for the specified file to the DTA.
   Private Sub WriteDTA(FilePath As String)
      Try
         Dim FileSize As Long = New FileInfo(FilePath).Length
         Dim Offset As Integer = ((DTA And &HFFFF0000%) >> &HC%) + (DTA And &HFFFF%)

         CPU.Memory(Offset + DTAE.Attribute) = ToByte(File.GetAttributes(FilePath))
         CPU.PutWord(Offset + DTAE.FileDate, DATE_TO_MSDOS_DATE(File.GetLastWriteTime(FilePath)))
         CPU.PutWord(Offset + DTAE.FileTime, TIME_TO_MSDOS_TIME(File.GetLastWriteTime(FilePath)))
         CPU.PutWord(Offset + DTAE.FileSize, CInt(FileSize And &HFF%))
         CPU.PutWord(Offset + DTAE.FileSize + &H2%, CInt(FileSize And &HFF00%) >> &H10%)
         WriteStringToMemory($"{Path.GetFileName(FilePath)}{ToChar(&H0%)}", Offset + DTAE.FileName)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub

   'This procedure writes to a file.
   Private Sub WriteFile(ByRef Flags As Integer)
      Try
         Dim Buffer As New StringBuilder
         Dim Bytes() As Byte = {}
         Dim Count As Integer = CInt(CPU.Registers(CPU8086Class.Registers16BitE.CX))
         Dim OpenFileToBeWritten As Tuple(Of FileStream, Integer) = Nothing
         Dim Position As New Integer
         Dim STDHandle As STDFileHandlesE = DirectCast(CPU.Registers(CPU8086Class.Registers16BitE.BX), STDFileHandlesE)

         Select Case STDHandle
            Case STDFileHandlesE.STDAUX, STDFileHandlesE.STDERR, STDFileHandlesE.STDOUT, STDFileHandlesE.STDPRN
               Position = (CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H4%) + CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX))
               Select Case STDHandle
                  Case STDFileHandlesE.STDAUX
                     SyncLock Synchronizer
                        CPUEvent.Append($"STDAUX:{NewLine}")

                        For Character As Integer = &H0% To Count - &H1%
                           Buffer.Append(ESCAPE_BYTE(CPU.Memory(Position And CPU8086Class.ADDRESS_MASK)))
                           Position += &H1%
                        Next Character

                        CPUEvent.Append($"{Buffer}{NewLine}")
                     End SyncLock
                  Case STDFileHandlesE.STDPRN
                     PrinterBuffer.Clear()

                     For Character As Integer = &H0% To Count - &H1%
                        PrinterBuffer.Append(ESCAPE_BYTE(CPU.Memory(Position And CPU8086Class.ADDRESS_MASK)))
                        Position += &H1%
                     Next Character

                     PrinterDocumentO.Print()
                  Case Else
                     For Character As Integer = &H0% To Count - &H1%
                        TeleType(CPU.Memory(Position And CPU8086Class.ADDRESS_MASK))
                        Position += &H1%
                     Next Character
               End Select

               CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=Count)
            Case Else
               OpenFileToBeWritten = OpenFiles.FirstOrDefault(Function(OpenedFile) OpenedFile.Item2 = CInt(CPU.Registers(CPU8086Class.Registers16BitE.BX)))
               Bytes = CPU.Memory.ToList.GetRange((CInt(CPU.Registers(CPU8086Class.SegmentRegistersE.DS)) << &H4%) Or CInt(CPU.Registers(CPU8086Class.Registers16BitE.DX)), Count).ToArray()

               Try
                  OpenFileToBeWritten.Item1.Write(Bytes, offset:=&H0%, Count)
                  CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=Count)
                  Flags = SET_BIT(Flags, False, CARRY_FLAG_INDEX)
               Catch MSDOSException As Exception
                  CPU.Registers(CPU8086Class.Registers16BitE.AX, NewValue:=GetMSDOSErrorCode(MSDOSException))
                  Flags = SET_BIT(Flags, True, CARRY_FLAG_INDEX)
               End Try
         End Select
      Catch ExceptionO As Exception
         DisplayException(ExceptionO.Message)
      End Try
   End Sub
End Module
