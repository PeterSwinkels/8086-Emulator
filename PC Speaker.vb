'This class's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports SDL2.SDL
Imports System
Imports System.Runtime.InteropServices

'This class contains the PC-Speaker.
Public Class PCSpeakerClass
   Private Const PIT_CLOCK As Double = 1193182.0   'Defines the PIT's frequency.
   Private Const VOLUME As Short = 20000S          'Defines the tone's frequency.

   Public Enabled As Boolean = False                             'Indicates whether or not the pc-speaker is enabled.
   Private AudioCallbackDelegate As SDL_AudioCallback = Nothing   'Contains the delegate used to play tones.
   Private DeviceID As UInteger = &H0UI                           'Contains the audio device's id.
   Private Frequency As Double = 0.0                              'Contains the tone's frequency.
   Private Phase As Double = 0.0                                  'Contains the phase while generating a tone.
   Private SampleRate As Integer = 44100                          'Contains the sample rate.

   'This procedure plays tones based on the current frequency.
   Private Sub AudioCallback(userdata As IntPtr, stream As IntPtr, len As Integer)
      Dim SampleCount As Integer = len \ 2
      Dim Buffer(&H0% To SampleCount - &H1%) As Short
      Dim LocalFrequency As New Double
      Dim LocalEnabled As New Boolean
      Dim LocalPhase As New Double
      Dim LocalSampleRate As New Integer
      Dim Position As New Double
      Dim SamplesPerPeriod As New Double

      SyncLock SYNCHRONIZER
         LocalFrequency = Frequency
         LocalEnabled = Enabled
         LocalPhase = Phase
         LocalSampleRate = SampleRate
      End SyncLock

      If LocalEnabled AndAlso LocalFrequency > &H0% AndAlso LocalFrequency <= Short.MaxValue Then
         SamplesPerPeriod = LocalSampleRate / LocalFrequency

         For Sample As Integer = &H0% To SampleCount - &H1%
            Position = LocalPhase Mod SamplesPerPeriod
            Buffer(Sample) = If(Position < SamplesPerPeriod / 2, VOLUME, -VOLUME)
            LocalPhase += 1
         Next Sample

         SyncLock SYNCHRONIZER
            Phase = LocalPhase
         End SyncLock

         Marshal.Copy(Buffer, &H0%, stream, SampleCount)
      Else
         Marshal.Copy(Buffer, &H0%, stream, SampleCount)
      End If
   End Sub

   'This procedure sets the frequency based on the specified divisor.
   Public Sub SetFrequency(Divisor As Integer)
      SyncLock SYNCHRONIZER
         Frequency = If(Divisor <= &H0%, &H0%, PIT_CLOCK / Divisor)
      End SyncLock
   End Sub

   'This procedure turns the pc-speaker on.
   Public Sub Start()
      Try
         Dim Desired As SDL_AudioSpec = Nothing
         Dim Obtained As SDL_AudioSpec = Nothing

         If DeviceID = &H0UI Then
            If SDL_Init(SDL_INIT_AUDIO) >= &H0% Then
               AudioCallbackDelegate = AddressOf AudioCallback

               Desired = New SDL_AudioSpec() With {.callback = AudioCallbackDelegate, .channels = &H1%, .format = AUDIO_S16SYS, .freq = SampleRate, .samples = &H200%}
               Obtained = New SDL_AudioSpec()
               DeviceID = SDL_OpenAudioDevice(Nothing, &H0%, Desired, Obtained, &H0%)
               If Not DeviceID = &H0UI Then
                  SampleRate = Obtained.freq
                  SDL_PauseAudioDevice(DeviceID, &H0%)
               End If
            End If
         End If
      Catch ExceptionO As Exception
         SyncLock SYNCHRONIZER
            CPU_EVENT.Append($"{ExceptionO.Message}{Environment.NewLine}")
         End SyncLock
      End Try
   End Sub

   'This procedure turns the pc-speaker off.
   Public Sub [Stop]()
      Try
         If Not DeviceID = &H0UI Then
            SDL_PauseAudioDevice(DeviceID, &H1%)
            SDL_CloseAudioDevice(DeviceID)
            DeviceID = &H0UI
         End If

         SDL_Quit()
      Catch ExceptionO As Exception
         SyncLock SYNCHRONIZER
            CPU_EVENT.Append($"{ExceptionO.Message}{Environment.NewLine}")
         End SyncLock
      End Try
   End Sub
End Class
