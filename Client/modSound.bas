Attribute VB_Name = "modSound"
'    Seyerdin Online - A MMO RPG based on Odyssey Online Classic - In memory of Clay Rance
'    Copyright (C) 2020  Samuel Cook and Eric Robinson
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.

'///////////////////////////////////////////
'///////   FMOD Based Sound       /////////
'/////////////////////////////////////////
Option Explicit

Public Enum FSOUND_INITMODES
    FSOUND_INIT_USEDEFAULTMIDISYNTH = &H1     'Causes MIDI playback to force software decoding.
    FSOUND_INIT_GLOBALFOCUS = &H2             'For DirectSound output - sound is not muted when window is out of focus.
    FSOUND_INIT_ENABLEOUTPUTFX = &H4          'For DirectSound output - Allows FSOUND_FX api to be used on global software mixer output!
    FSOUND_INIT_ACCURATEVULEVELS = &H8        'This latency adjusts FSOUND_GetCurrentLevels, but incurs a small cpu and memory hit
    FSOUND_INIT_PS2_DISABLECORE0REVERB = &H10 'PS2 only - Disable reverb on CORE 0 to regain SRAM
    FSOUND_INIT_PS2_DISABLECORE1REVERB = &H20 'PS2 only - Disable reverb on CORE 1 to regain SRAM
    FSOUND_INIT_PS2_SWAPDMACORES = &H40       'PS2 only - By default FMOD uses DMA CH0 for mixing, CH1 for uploads, this flag swaps them around
    FSOUND_INIT_DONTLATENCYADJUST = &H80      'Callbacks are not latency adjusted, and are called at mix time.  Also information functions are immediate
    FSOUND_INIT_GC_INITLIBS = &H100             'Gamecube only - Initializes GC audio libraries
    FSOUND_INIT_STREAM_FROM_MAIN_THREAD = &H200 'Turns off fmod streamer thread, and makes streaming update from FSOUND_Update called by the user
End Enum

Public Enum FSOUND_MODES
    FSOUND_LOOP_OFF = 1            ' For non looping samples.
    FSOUND_LOOP_NORMAL = 2         ' For forward looping samples.
    FSOUND_LOOP_BIDI = 4           ' For bidirectional looping samples.  (no effect if in hardware).
    FSOUND_8BITS = 8               ' For 8 bit samples.
    FSOUND_16BITS = 16             ' For 16 bit samples.
    FSOUND_MONO = 32               ' For mono samples.
    FSOUND_STEREO = 64             ' For stereo samples.
    FSOUND_UNSIGNED = 128          ' For source data containing unsigned samples.
    FSOUND_SIGNED = 256            ' For source data containing signed data.
    FSOUND_DELTA = 512             ' For source data stored as delta values.
    FSOUND_IT214 = 1024            ' For source data stored using IT214 compression.
    FSOUND_IT215 = 2048            ' For source data stored using IT215 compression.
    FSOUND_HW3D = 4096             ' Attempts to make samples use 3d hardware acceleration. (if the card supports it)
    fsound_2d = 8192               ' Ignores any 3d processing.  overrides FSOUND_HW3D.  Located in software.
    FSOUND_STREAMABLE = 16384      ' For realtime streamable samples.  If you dont supply this sound may come out corrupted.
    FSOUND_LOADMEMORY = 32768      ' For FSOUND_Sample_Load - name will be interpreted as a pointer to data
    FSOUND_LOADRAW = 65536         ' For FSOUND_Sample_Load/FSOUND_Stream_Open - will ignore file format and treat as raw pcm.
    FSOUND_MPEGACCURATE = 131072   ' For FSOUND_Stream_Open - scans MP2/MP3 (VBR also) for accurate FSOUND_Stream_GetLengthMs/FSOUND_Stream_SetTime.
    FSOUND_FORCEMONO = 262144      ' For forcing stereo streams and samples to be mono - needed with FSOUND_HW3D - incurs speed hit
    FSOUND_HW2D = 524288           ' 2d hardware sounds.  allows hardware specific effects
    FSOUND_ENABLEFX = 1048576      ' Allows DX8 FX to be played back on a sound.  Requires DirectX 8 - Note these sounds cant be played more than once, or have a changing frequency
    FSOUND_MPEGHALFRATE = 2097152  ' For FMODCE only - decodes mpeg streams using a lower quality decode, but faster execution
    FSOUND_XADPCM = 4194304        ' For XBOX only - Describes a user sample that its contents are compressed as XADPCM
    FSOUND_VAG = 8388608           ' For PS2 only - Describes a user sample that its contents are compressed as Sony VAG format.
    FSOUND_NONBLOCKING = 16777216  ' For FSOUND_Stream_OpenFile - Causes stream to open in the background and not block the foreground app - stream plays only when ready.
    FSOUND_GCADPCM = &H2000000       ' For Gamecube only - Contents are compressed as Gamecube DSP-ADPCM format */
    FSOUND_MULTICHANNEL = &H4000000  ' For PS2 only - Contents are interleaved into a multi-channel (more than stereo) format */
    FSOUND_USECORE0 = &H8000000      ' For PS2 only - Sample/Stream is forced to use hardware voices 00-23 */
    FSOUND_USECORE1 = &H10000000     ' For PS2 only - Sample/Stream is forced to use hardware voices 24-47 */
    FSOUND_LOADMEMORYIOP = &H20000000 ' For PS2 only - "name" will be interpreted as a pointer to data for streaming and samples.  The address provided will be an IOP address
    FSOUND_STREAM_NET = &H80000000    ' Specifies an internet stream
    
    fsound_normal = FSOUND_16BITS Or FSOUND_SIGNED Or FSOUND_MONO
End Enum

Public Declare Function FSOUND_Init Lib "fmod.dll" Alias "_FSOUND_Init@12" (ByVal mixrate As Long, ByVal maxchannels As Long, ByVal Flags As FSOUND_INITMODES) As Byte
Public Declare Function FSOUND_Close Lib "fmod.dll" Alias "_FSOUND_Close@0" () As Long
Public Declare Function FSOUND_Sample_Load Lib "fmod.dll" Alias "_FSOUND_Sample_Load@20" (ByVal Index As Long, ByVal Name As String, ByVal mode As FSOUND_MODES, ByVal offset As Long, ByVal length As Long) As Long
Public Declare Function FSOUND_PlaySound Lib "fmod.dll" Alias "_FSOUND_PlaySound@8" (ByVal Channel As Long, ByVal Sptr As Long) As Long
Public Declare Function FSOUND_StopSound Lib "fmod.dll" Alias "_FSOUND_StopSound@4" (ByVal Channel As Long) As Byte
Public Declare Function FSOUND_Sample_Free Lib "fmod.dll" Alias "_FSOUND_Sample_Free@4" (ByVal Sptr As Long) As Long
Public Declare Function FSOUND_Stream_Open Lib "fmod.dll" Alias "_FSOUND_Stream_Open@16" (ByVal Filename As String, ByVal mode As FSOUND_MODES, ByVal offset As Long, ByVal length As Long) As Long
Public Declare Function FSOUND_Stream_Close Lib "fmod.dll" Alias "_FSOUND_Stream_Close@4" (ByVal stream As Long) As Byte
Public Declare Function FSOUND_Stream_Play Lib "fmod.dll" Alias "_FSOUND_Stream_Play@8" (ByVal Channel As Long, ByVal stream As Long) As Long
Public Declare Function FSOUND_Stream_Stop Lib "fmod.dll" Alias "_FSOUND_Stream_Stop@4" (ByVal stream As Long) As Byte
Public Declare Function FSOUND_SetSFXMasterVolume Lib "fmod.dll" Alias "_FSOUND_SetSFXMasterVolume@4" (ByVal Volume As Long) As Long
Public Declare Function FSOUND_SetVolume Lib "fmod.dll" Alias "_FSOUND_SetVolume@8" (ByVal Channel As Long, ByVal Vol As Long) As Byte
Public Declare Function FSOUND_Stream_SetEndCallback Lib "fmod.dll" Alias "_FSOUND_Stream_SetEndCallback@12" (ByVal stream As Long, ByVal callback As Long, ByVal userdata As Long) As Byte

Public Const FSOUND_FREE = -1
Public Const MUSIC_FADE_LENGTH = 200

Public SongStartDelay As Long
Public CurrentSongNum As Long
Private CurrentStream As Long
Private CurrentStreamChannel As Long
Public MusicFading As Boolean
Public MusicFade As Single
Public NextSong As Long

Type SoundType
    Sample As Long
    Loaded As Boolean
    LastUsed As Long
End Type

Public Sounds(1 To 255) As SoundType
    
Sub Sound_Init()
    FSOUND_Init 65535, 32, 0
End Sub

Sub Sound_Unload()
    Dim A As Long
    For A = 1 To 255
    
    Next A
    Sound_StopStream
    FSOUND_Close
End Sub

Sub Sound_Free(ByVal Sound As Long)
    If Sound > 0 And Sound <= 255 Then
        If Sounds(Sound).Loaded Then
            Sounds(Sound).Loaded = False
            Sounds(Sound).LastUsed = 0
            FSOUND_Sample_Free Sounds(Sound).Sample
        End If
    End If
End Sub

Sub Sound_PlaySound(ByVal Sound As Long, Optional volumepercent As Single = 1)
    Dim tChannel As Long
    If Options.Wav Then
       
        If Not Sounds(Sound).Loaded Then
            If Exists("Data/Sound/Sound" & Sound & ".wav") Then
                Sounds(Sound).Loaded = True
                Sounds(Sound).Sample = FSOUND_Sample_Load(FSOUND_FREE, "Data/Sound/Sound" & Sound & ".wav", fsound_normal, 0, 0)
            Else
                Exit Sub
            End If
        End If

        Sounds(Sound).LastUsed = GetTickCount + 360000
        tChannel = FSOUND_PlaySound(FSOUND_FREE, Sounds(Sound).Sample)
        FSOUND_SetVolume tChannel, Options.SoundVolume * volumepercent

'        If Exists("Data/Sound/sound" & Sound & ".wav") Then
'            Sounds.Add "Data/Sound/Sound" & Sound & ".wav"
'            'tStream = FSOUND_Stream_Open("Data/Sound/sound" & Sound & ".wav", FSOUND_NORMAL, 0, 0)
'            'tStreamChannel = FSOUND_Stream_Play(FSOUND_FREE, tStream)
'            'FSOUND_SetVolume tStreamChannel, Options.SoundVolume
'        End If
    End If
End Sub


Sub Sound_PlayStream(Song As Long)
    If Options.MIDI Then
        If Song <> CurrentSongNum Then
            If CurrentStream <> 0 Then
                If MusicFading = False Then
                    MusicFading = True
                    MusicFade = MUSIC_FADE_LENGTH
                End If
                NextSong = Song
            Else
                If Exists("Data/Music/" & Song & ".mp3") Then
                    CurrentSongNum = Song
                    CurrentStream = FSOUND_Stream_Open("Data/Music/" & Song & ".mp3", fsound_normal, 0, 0)
                    CurrentStreamChannel = FSOUND_Stream_Play(FSOUND_FREE, CurrentStream)
                    FSOUND_SetVolume CurrentStreamChannel, Options.MusicVolume
                Else
                    CurrentSongNum = 0
                End If
            End If
        End If
    End If
End Sub



Sub Sound_StopStream()
    If CurrentStream <> 0 Then
        If MusicFading = False Then
            MusicFading = True
            MusicFade = MUSIC_FADE_LENGTH
        Else
            MusicFading = False
            MusicFade = 0
            FSOUND_Stream_Stop CurrentStream
            CurrentStream = 0
            CurrentStreamChannel = 0
            If NextSong > 0 Then
                Sound_PlayStream NextSong
                NextSong = 0
            End If
        End If
    End If
End Sub

Sub Sound_SetStreamFadeVolume()
    Dim Volume As Long
    Volume = Options.MusicVolume * (MusicFade / MUSIC_FADE_LENGTH)
    Sound_SetStreamVolume Volume
End Sub

Sub Sound_SetStreamVolume(Volume As Long)
    If Volume <= 0 Then Volume = 0
    If Volume > 255 Then Volume = 255
    FSOUND_SetVolume CurrentStreamChannel, Volume
End Sub
