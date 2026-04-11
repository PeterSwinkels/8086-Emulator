'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.Marshal
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
   Public Shared Function waveOutUnprepareHeader(
    ByVal hwo As IntPtr,
    ByRef pwh As WAVEHDR,
    ByVal cbwh As Integer) As Integer
   End Function

   <DllImport("Winmm.dll")>
   Private Shared Function waveOutWrite(
    hWaveOut As IntPtr,
    ByRef lpWaveOutHdr As WAVEHDR,
    uSize As Integer) As Integer
   End Function

   Private Const CALLBACK_NULL As Integer = &H0%
   Private Const MMSYSERR_NOERROR As Integer = &H0%
   Private Const WAVE_FORMAT_PCM As Integer = &H1%
   Private Const WAVE_MAPPER As Integer = -1%

   Private Const BUFFER_SIZE As Integer = &H400%   'Defines the wave buffer size.
   Private Const PIT_CLOCK As Double = 1193182     'Defines the PIT's frequency.
   Private Const SAMPLE_RATE As Integer = 44100    'Defines the sample rate.
   Private Const VOLUME As Integer = &H1000%       'Defines the volume.

   Private AudioThread As Thread = Nothing   'Contains the thread that drives the pc-speaker.
   Private Frequency As New Double           'Contains the frequency of the tone to be generated.
   Private Phase As New Double               'Contains the phase.
   Private Running As Boolean = False        'Indicates whether a tone is being generated.
   Public Enabled As Boolean = False         'Indicates whether the pc-speaker is enabled.

   'This procedure generates a tone on a loop.
   Private Sub AudioLoop()
      Dim Format As New WAVEFORMATEX() With {.nChannels = &H1%, .nSamplesPerSec = SAMPLE_RATE, .wBitsPerSample = &H10%, .wFormatTag = WAVE_FORMAT_PCM, .nBlockAlign = CShort(.nChannels * .wBitsPerSample / &H8%), .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign}
      Dim Header As WAVEHDR = Nothing
      Dim RawWaveData() As Byte = {}
      Dim Samples(&H0% To BUFFER_SIZE - &H1%) As UShort
      Dim WaveH As IntPtr = Nothing

      If waveOutOpen(WaveH, WAVE_MAPPER, Format, IntPtr.Zero, IntPtr.Zero, CALLBACK_NULL) = MMSYSERR_NOERROR Then
         While Running
            GenerateSamples(Samples)
            ReDim RawWaveData(&H0% To Samples.Length * &H2% - &H1%)
            Buffer.BlockCopy(Samples, &H0%, RawWaveData, &H0%, RawWaveData.Length)
            Header = New WAVEHDR With {.dwBufferLength = RawWaveData.Length, .lpData = AllocHGlobal(RawWaveData.Length)}
            Copy(RawWaveData, &H0%, Header.lpData, RawWaveData.Length)

            If waveOutPrepareHeader(WaveH, Header, SizeOf(Header)) = MMSYSERR_NOERROR Then
               waveOutWrite(WaveH, Header, SizeOf(Header))
               Thread.Sleep(10)
               waveOutUnprepareHeader(WaveH, Header, SizeOf(Header))
               FreeHGlobal(Header.lpData)
            End If
         End While

         waveOutClose(WaveH)
      End If
   End Sub

   'This procedure generates the samples for the tone to be generated.
   Private Sub GenerateSamples(Samples() As UShort)
      For Sample As Integer = 0 To Samples.Length - 1
         If Enabled AndAlso Frequency > 0 AndAlso Frequency < Short.MaxValue Then
            Phase += Frequency / SAMPLE_RATE
            If Phase >= 1 Then
               Phase -= 1
            End If
            Samples(Sample) = CUShort(If(Phase < 0.5, VOLUME, (-VOLUME) And &HFFFF%))
         Else
            Samples(Sample) = &H0%
         End If
      Next Sample
   End Sub

   'This procedure sets the frequency of the tone to be generated based on the specified divisor.
   Public Sub SetFrequency(Divisor As Integer)
      Frequency = If(Divisor <= 0, 0, PIT_CLOCK / Divisor)
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
      If AudioThread IsNot Nothing Then
         AudioThread.Join()
      End If
   End Sub
End Class


