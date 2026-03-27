'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Runtime.InteropServices
Imports System.Threading

'This class contains the PC-Speaker.
Public Class PCSpeakerClass
   <StructLayout(LayoutKind.Sequential)>
   Public Structure WAVEFORMATEX
      Public wFormatTag As Short
      Public nChannels As Short
      Public nSamplesPerSec As Integer
      Public nAvgBytesPerSec As Integer
      Public nBlockAlign As Short
      Public wBitsPerSample As Short
      Public cbSize As Short
   End Structure

   <StructLayout(LayoutKind.Sequential)>
   Public Structure WAVEHDR
      Public lpData As IntPtr
      Public dwBufferLength As Integer
      Public dwBytesRecorded As Integer
      Public dwUser As IntPtr
      Public dwFlags As Integer
      Public dwLoops As Integer
      Public lpNext As IntPtr
      Public reserved As IntPtr
   End Structure

   <DllImport("Winmm.dll")>
   Private Shared Function waveOutClose(
    hWaveOut As IntPtr) As Integer
   End Function

   <DllImport("Winmm.dll")>
   Private Shared Function waveOutOpen(
    ByRef hWaveOut As IntPtr,
    uDeviceID As Integer,
    ByRef lpFormat As WAVEFORMATEX,
    dwCallback As IntPtr,
    dwInstance As IntPtr,
    dwFlags As Integer) As Integer
   End Function

   <DllImport("Winmm.dll")>
   Private Shared Function waveOutPrepareHeader(
    hWaveOut As IntPtr,
    ByRef lpWaveOutHdr As WAVEHDR,
    uSize As Integer) As Integer
   End Function

   <DllImport("Winmm.dll")>
   Private Shared Function waveOutWrite(
    hWaveOut As IntPtr,
    ByRef lpWaveOutHdr As WAVEHDR,
    uSize As Integer) As Integer
   End Function

   Private Const PIT_CLOCK As Double = 1193182    'Defines the PIT's frequency.
   Private Const SAMPLE_RATE As Integer = 44100   'Defines the sample rate.

   Private AudioThread As Thread = Nothing   'Contains the thread that drives the pc-speaker.
   Private Frequency As New Double           'Contains the frequency of the tone to be generated.
   Private Phase As New Double               'Contains the phase.
   Private Running As Boolean = False        'Indicates whether a tone is being generated.
   Public Enabled As Boolean = False         'Indicates whether the pc-speaker is enabled.


   'This procedure generates a tone on a loop.
   Private Sub AudioLoop()
      Dim BufferSize As Integer = &H400%
      Dim Bytes() As Byte = {}
      Dim Format As New WAVEFORMATEX()
      Dim Header As WAVEHDR = Nothing
      Dim SampleBuffer(&H0% To BufferSize - &H1%) As Short
      Dim WaveH As IntPtr = Nothing

      Format.nChannels = &H1%
      Format.nSamplesPerSec = SAMPLE_RATE
      Format.wBitsPerSample = &H10%
      Format.wFormatTag = &H1%
      Format.nBlockAlign = CShort(Format.nChannels * Format.wBitsPerSample / &H8%)
      Format.nAvgBytesPerSec = Format.nSamplesPerSec * Format.nBlockAlign

      waveOutOpen(WaveH, -1, Format, IntPtr.Zero, IntPtr.Zero, &H0%)

      While Running
         GenerateSamples(SampleBuffer)

         ReDim Bytes(&H0% To SampleBuffer.Length * &H2% - &H1%)
         Buffer.BlockCopy(SampleBuffer, &H0%, Bytes, &H0%, Bytes.Length)

         Header = New WAVEHDR()

         Header.dwBufferLength = Bytes.Length
         Header.lpData = Marshal.AllocHGlobal(bytes.Length)

         Marshal.Copy(Bytes, &H0%, Header.lpData, Bytes.Length)

         waveOutPrepareHeader(WaveH, header, Marshal.SizeOf(header))
         waveOutWrite(WaveH, Header, Marshal.SizeOf(Header))

         Thread.Sleep(10)
      End While

      waveOutClose(WaveH)
   End Sub

   'This procedure generates the samples for the tone to be generated.
   Private Sub GenerateSamples(SampleBuffer() As Short)
      For Sample As Integer = 0 To SampleBuffer.Length - 1
         If Enabled AndAlso Frequency > 0 AndAlso Frequency < Short.MaxValue Then
            Phase += Frequency / SAMPLE_RATE

            If Phase >= 1 Then
               Phase -= 1
            End If

            If Phase < 0.5 Then
               SampleBuffer(Sample) = 12000
            Else
               SampleBuffer(Sample) = -12000
            End If
         Else
            SampleBuffer(Sample) = 0
         End If
      Next Sample
   End Sub

   'This procedure sets the frequency of the tone to be generated based on the specified divisor.
   Public Sub SetFrequency(Divisor As Integer)
      If Divisor <= 0 Then
         Frequency = 0
      Else
         Frequency = PIT_CLOCK / Divisor
      End If
   End Sub

   'This procedure starts the pc-speaker.
   Public Sub Start()
      Running = True
      AudioThread = New Thread(AddressOf AudioLoop)
      AudioThread.IsBackground = True
      AudioThread.Start()
   End Sub

   'This procedure stops the pc-speaker.
   Public Sub [Stop]()
      Running = False
      If AudioThread IsNot Nothing Then AudioThread.Join()
   End Sub
End Class


