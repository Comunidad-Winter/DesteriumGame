VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsAudio.cls - Controls all audio using DirectSound and DirectMusic
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@hotmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

'16/03/2004 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - First Release
'
'03/03/2007 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Add: Sound and Music activation / deactivation from within the class.
'   - Add: Documentation
'   - Change: Volume is no longer the dB amplification / atenuation to be
'           applied but a percentage for easier interaction with user.
'
'30/03/2008 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Add: 3D-like sound effects by altering frequency to emulate Doppler effect,
'           volume to emulate distance and panning for position.
'
'28/07/2008 - Franco Zeoli (Noich) - (zeolifranco@gmail.com)
'   -Add: Optional sound effects
'
'23/03/2010 - Torres Patricio(Pato) - (paatoo.torres@gmail.com)
'   -Add: ORE code to reproduce MP3

''
' Controls all audio using DirectSound and DirectMusic 7.
'
' @file     clsAudio.cls
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.2.1
' @date     20070303

Option Explicit

''
' Maximum number of sound buffers to hold in memory at any time.
Private Const NumSoundBuffers As Byte = 30

''
' The distance at which the sound is inaudible.
Private Const MAX_DISTANCE_TO_SOURCE As Integer = 150

''
' The frequency variation coefficient for Doppler effect simulation.
Private Const DELTA_FQ As Integer = 75

''
' Custom sound buffer structure.
'
' @param    FileName    The name of the file loaded in this buffer
' @param    looping     Set to true if the sound will loop until stopped
' @param    X           The X position of the sound in map coordinates
' @param    Y           The Y position of the sound in map coordinates
' @param    normalFq    The frequency of the sound when not affected by Doppler effect.
' @param    Buffer      The DirectSoundBuffer object to wich the file is loaded
Private Type SoundBuffer
    FileName As String
    looping As Boolean
    X As Byte
    Y As Byte
    normalFq As Long
    Buffer As DirectSoundBuffer
End Type

''
' States how to set a sound's looping state.
'
' @param    Default     If wave is in memory use it's looping setting, otherwise false
' @param    Disabled    Never loop
' @param    Enabled     Loop until stopped
Public Enum LoopStyle
    Default = 0
    Disabled = 1
    Enabled = 2
End Enum

''
' Set to true once the object is created to prevent multiple initializations of the same object.
Dim inited As Boolean

''
'The one and only DirectSound object
Dim DirectSound As DirectSound

''
'Array of all existing sound buffers
Dim DSBuffers(1 To NumSoundBuffers) As SoundBuffer

''
'Directory from which to load sound files by default
Dim SoundPath As String

''
'The volume for sounds (FXs)
Dim SndVolume As Long

''
'DirectMusic's Performance object
Dim Performance As DirectMusicPerformance

''
'Currently loaded segment
Dim Segment As DirectMusicSegment

''
'The one and only DirectMusic Loader
Dim Loader As DirectMusicLoader

''
'State of the currently loaded segment
Dim SegState As DirectMusicSegmentState

''
'States wether music is supported by hardware or not
Dim Music As Boolean

''
'Directory from which to load music files by default
Dim MusicPath As String

''
'Last loaded music file. Used to be able to load it when music gets activated
Dim CurrentMusicFile As String

''
'States wether music is currently activated or not
Dim mActivated As Boolean

''
'States wether sound is currently activated or not
Dim sActivated As Boolean

''
'States the last position where the listener was in the X axis
Dim lastPosX As Integer

''
'States the last position where the listener was in the Y axis
Dim lastPosY As Integer

''
'States wether sound effects are currently activated or not
Dim sEffectsActivated As Boolean

'**** Used By MP3 Playing. *****
Dim IMC   As IMediaControl
Dim IBA   As IBasicAudio
Dim IME   As IMediaEvent
Dim IMPos As IMediaPosition

''
'Constructor. Performs no action at all.

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'
'***************************************************
End Sub

''
'Destructor. Releases all created objects assuring no memory-leaks.

Private Sub Class_Terminate()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 12/31/09
      'Destroy's all DirectMusic and DirectSound objects
      '12/31/09 - Pato: Use StopWave method to prevent errors.
      '***************************************************
10    On Local Error Resume Next
          'Stop every channel being used and destroy the buffer
20        Call StopWave
          
          'Stop music
30        Call Performance.Stop(Segment, SegState, 0, 0)
          
40        Call Performance.CloseDown
          
50        Set DirectSound = Nothing
          
60        Set Segment = Nothing
70        Set Performance = Nothing
80        Set Loader = Nothing
          
90        inited = False
End Sub

''
'Initializes the clsAudio object.
'
' @param    dX          DirectX7 object to be used when creating the DirectSound and Directmusic objects.
' @param    hWnd        Handle to the window to be used by DirectMusic when playing midi files.
' @param    sound_path  Directory in wich to search for sound files.
' @param    music_path  Directory in wich to search for music files.
'
' @return   True if no errors occured, false otherwise.

Public Function Initialize(ByRef dX As DirectX7, ByVal hWnd As Long, ByVal _
    sound_path As String, ByVal music_path As String) As Boolean
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Creates all DS and DM objects
      '***************************************************
10    On Error GoTo ErrHandler
          
          ' Prevent multiple creations
20        If inited Then Exit Function
          
          'Validate DX object
30        If dX Is Nothing Then GoTo ErrHandler
          
          'Initialize DirectSound
40        Set DirectSound = dX.DirectSoundCreate("")
          
50        Call DirectSound.SetCooperativeLevel(hWnd, DSSCL_PRIORITY)
          
          'Set sound path
60        SoundPath = sound_path
          
          'Set default volume
70        SndVolume = 0
          
          'Initialize DirectMusic
80        Music = InitializedirectMusic(dX, hWnd)
          
          'Set music path
90        MusicPath = music_path
          
100       mActivated = True
110       sActivated = True
          
120       CurrentMusicFile = vbNullString
          
          'Set tempo to 0 and volume of music
130       MusicTempo = 0
140       Call Performance.SetMasterVolume(200)
          
150       inited = True
          
160       Initialize = True
170   Exit Function

ErrHandler:
End Function

''
' Creates and configures all DirectMusic objects.
'
' @param    dX          DirectX7 object to be used when creating the DirectSound and Directmusic objects.
' @param    hWnd        Handle to the window to be used by DirectMusic when playing midi files.
'
' @return   True if no errors occured, false otherwise.

Private Function InitializedirectMusic(ByRef dX As DirectX7, ByVal hWnd As _
    Long) As Boolean
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Initialize DM. Returns false if music isn't supported by HW
      '***************************************************
10    On Error GoTo ErrHandler
20        Set Loader = DirectX.DirectMusicLoaderCreate()
          
30        Set Performance = dX.DirectMusicPerformanceCreate()
40        Call Performance.Init(DirectSound, hWnd)
50        Call Performance.SetMasterAutoDownload(True)        'Enable auto download of instruments
          
          'Set to port default, 1 channel group is enough for MIDIs
60        Call Performance.SetPort(-1, 1)
          
70        InitializedirectMusic = True
80    Exit Function

ErrHandler:
End Function

''
' Plays a new WAVE file.
'
' @param    FileName    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    srcX        The X position of the sound source, in map coordinates. -1 if sound should not be altered by 3D alterations.
' @param    srcY        The Y position of the sound source, in map coordinates. -1 if sound should not be altered by 3D alterations.
' @param    LoopSound   The loop state for the sound.
'
' @return   The buffer index in which the sound is playing, or 0 in case of error.

Public Function PlayWave(ByVal FileName As String, Optional ByVal srcX As Byte _
    = 0, Optional ByVal srcY As Byte = 0, Optional ByVal LoopSound As LoopStyle = _
    LoopStyle.Default) As Long
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 28/07/08
      'Last Modified by: Franco Zeoli (Noich)
      'Added support for 3D sound effects.
      '***************************************************
10    On Error GoTo ErrHandler
          Dim bufferIndex As Long
          
20        If Not SoundActivated Then Exit Function
          
          'Get the buffer index were wave was loaded
30        bufferIndex = LoadWave(FileName, LoopSound)
40        If bufferIndex = 0 Then Exit Function   'If an error ocurred abort
          
50        With DSBuffers(bufferIndex)
60            If .looping Then
70                .Buffer.Play DSBPLAY_LOOPING
80            Else
90                .Buffer.Play DSBPLAY_DEFAULT
100           End If
              
              'Apply volume
110           Call .Buffer.SetVolume(SndVolume)
              
              'Store position
120           .X = srcX
130           .Y = srcY
              
              'Get standard frequency
140           .normalFq = .Buffer.GetFrequency()
150       End With
          
160       PlayWave = bufferIndex
          
170       If sEffectsActivated Then
180           If srcX <> 0 And srcY <> 0 Then
190               Call Update3DSound(bufferIndex, 0, 0)
200           End If
210       End If
220   Exit Function

ErrHandler:
End Function

Private Sub ResetSoundValues()
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modification: 03/04/10
      'Reset the frequency, volume and pan.
      '***************************************************
          Dim i As Long
          
10        For i = 1 To NumSoundBuffers
20            If Not DSBuffers(i).Buffer Is Nothing Then
30                With DSBuffers(i)
40                    Call .Buffer.SetFrequency(.normalFq)
50                    Call .Buffer.SetVolume(SndVolume)
60                    Call .Buffer.SetPan(0)
70                End With
80            End If
90        Next i
End Sub

Private Sub Update3DSound(ByVal bufferIndex As Long, ByVal deltaX As Integer, _
    ByVal deltaY As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 05/09/09
      'Last Modified by: Marco Vanotti (MarKoxX)
      'Fixed Dopplers formula
      'Refreshes the sound's arguments.
      'Fixed the sndVolume formula
      '***************************************************
          Dim linearDistanceOld As Single
          Dim linearDistanceNew As Single
          Dim deltaDistance As Single
          Dim distanceXOld As Long
          Dim distanceYOld As Long
          Dim distanceXNew As Long
          Dim distanceYNew As Long
          Dim perceivedFq As Long
          
10        With DSBuffers(bufferIndex)
20            distanceXOld = .X - lastPosX
30            distanceYOld = .Y - lastPosY
              
40            distanceXNew = distanceXOld + deltaX
50            distanceYNew = distanceYOld + deltaY
              
60            linearDistanceOld = Sqr(distanceXOld * distanceXOld + distanceYOld * _
                  distanceYOld)
70            linearDistanceNew = Sqr(distanceXNew * distanceXNew + distanceYNew * _
                  distanceYNew)
              
80            deltaDistance = linearDistanceNew - linearDistanceOld
              
              'Update frequency emulating Doppler effect
90            perceivedFq = .normalFq * (1 - deltaDistance / DELTA_FQ)
100           If perceivedFq < DSBFREQUENCY_MIN Then perceivedFq = DSBFREQUENCY_MIN
              
110           Call .Buffer.SetFrequency(perceivedFq)
              
              'Set volumen amortiguation according to distance
120           If linearDistanceNew > MAX_DISTANCE_TO_SOURCE Then
130               Call .Buffer.SetVolume(SndVolume + (DSBVOLUME_MIN - SndVolume))
140           Else
150               Call .Buffer.SetVolume(SndVolume + (linearDistanceNew / _
                      MAX_DISTANCE_TO_SOURCE) * (DSBVOLUME_MIN - SndVolume))
160           End If
              
              'Prevent division by zero
170           If linearDistanceNew = 0 Then linearDistanceNew = 1
              
              'Set panning according to relative position of the source to the listener
180           Call .Buffer.SetPan((distanceXNew / linearDistanceNew) * DSBPAN_RIGHT)
190       End With
End Sub

''
' Updates 3D sounds based on the movement of the listener.
'
' @param    x   The new coordinate of the listener in the x axis (consider 0 as the furthest point to the left).
' @param    y   The new coordinate of the listener in the y axis (consider 0 as the furthest point to the front).
'
' @return   The buffer index in which the sound is playing, or 0 in case of error.

Public Sub MoveListener(ByVal X As Integer, ByVal Y As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 28/07/08
      'Last Modification by: Franco Zeoli (Noich)
      'Moves the position of the listener updating all 3D sounds.
      '***************************************************
          Dim i As Long
          Dim deltaX As Integer
          Dim deltaY As Integer
          
10        If Not sEffectsActivated Then Exit Sub
          
20        deltaX = X - lastPosX
30        deltaY = Y - lastPosY
          
40        For i = 1 To NumSoundBuffers
50            If Not DSBuffers(i).Buffer Is Nothing Then
60                If DSBuffers(i).Buffer.GetStatus And DSBSTATUS_PLAYING Then
70                    If DSBuffers(i).X <> 0 And DSBuffers(i).Y <> 0 Then
80                        Call Update3DSound(i, deltaX, deltaY)
90                    End If
100               End If
110           End If
120       Next i
          
130       lastPosX = X
140       lastPosY = Y
End Sub

''
' Loads a new WAVE file.
'
' @param    FileName    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    looping     The loop state for the sound.
'
' @return   The buffer index in which the sound is playing, or 0 in case of error.

Private Function LoadWave(ByVal FileName As String, ByVal looping As LoopStyle) _
    As Long
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 04/28/09
      'Last Modification by: Marco Vanotti (MarKoxX)
      'Loads a new wave into memory (if it's not already loaded) and returns
      'the buffer index
      '***************************************************
10    On Error GoTo ErrHandler
          Dim bufferDesc As DSBUFFERDESC
          Dim waveFormat As WAVEFORMATEX
          Dim i As Long
          
20        FileName = UCase$(FileName)
          
30        If Not FileExist(SoundPath & FileName, vbArchive) Then Exit Function
          
          'Check if the buffer is in memory and not playing
40        For i = 1 To NumSoundBuffers
50            If DSBuffers(i).FileName = FileName Then
60                If DSBuffers(i).Buffer.GetStatus = 0 Or _
                      DSBuffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Then
                      'Found it!!! We just play this one :)
70                    LoadWave = i
                      
                      'Set looping if needed
80                    If looping <> LoopStyle.Default Then DSBuffers(i).looping = _
                          (looping = LoopStyle.Enabled)
                      
                      'Reset default sound frecuency to avoid issues.
90                    Call DSBuffers(i).Buffer.SetFrequency(DSBuffers(i).normalFq)
                      
100                   Exit Function
110               End If
120           End If
130       Next i
          
          'Not in memory, search for an empty buffer
140       For i = 1 To NumSoundBuffers
150           If DSBuffers(i).Buffer Is Nothing Then
160               Exit For
170           End If
180       Next i
          
          'No empty buffers?? Search for one not being played and use that slot
190       If i > NumSoundBuffers Then
200           For i = 1 To NumSoundBuffers
210               If DSBuffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Or _
                      DSBuffers(i).Buffer.GetStatus() = 0 Then
220                   Exit For
230               End If
240           Next i
250       End If
          
          'Still no free buffers?? Get the first non-looping sound
260       If i > NumSoundBuffers Then
270           For i = 1 To NumSoundBuffers
280               If Not (DSBuffers(i).Buffer.GetStatus() And DSBSTATUS_LOOPING) Then
290                   Call DSBuffers(i).Buffer.Stop
300                   Exit For
310               End If
320           Next i
330       End If
          
          'If every sound is looping, just ignore new sound unless it's looping too (in which case we set it at the first slot by default)
340       If i > NumSoundBuffers Then
350           If Not looping Then Exit Function
360           i = 1
370       End If
          
          'Create new sound
380       bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or _
              DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
          
390       waveFormat.nFormatTag = WAVE_FORMAT_PCM
400       waveFormat.nChannels = 2
410       waveFormat.lSamplesPerSec = 22050
420       waveFormat.nBitsPerSample = 16
430       waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * _
              waveFormat.nChannels
440       waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * _
              waveFormat.nBlockAlign
          
450       With DSBuffers(i)
460           Set .Buffer = Nothing   'Get rid of any previous data
              
470           .FileName = FileName
480           .looping = (looping = LoopStyle.Enabled)
              
490           Set .Buffer = DirectSound.CreateSoundBufferFromFile(SoundPath & _
                  FileName, bufferDesc, waveFormat)
500       End With
          
510       LoadWave = i
520   Exit Function

ErrHandler:
End Function

''
' Stops a given sound or all of them.
'
' @param    bufferIndex     The index of the sound file to stop playing. If 0, all sounds are stopped.

Public Sub StopWave(Optional ByVal bufferIndex As Long = 0)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Stops the sound in the given buffer index, or every sound if
      'index is 0
      '***************************************************
10        If bufferIndex > 0 And bufferIndex <= NumSoundBuffers Then
20            If DSBuffers(bufferIndex).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
30                Call DSBuffers(bufferIndex).Buffer.Stop
40            End If
50        ElseIf bufferIndex = 0 Then
              Dim i As Long
60            For i = 1 To NumSoundBuffers
70                If Not DSBuffers(i).Buffer Is Nothing Then
80                    If DSBuffers(i).Buffer.GetStatus() <> 0 And _
                          DSBuffers(i).Buffer.GetStatus() <> DSBSTATUS_BUFFERLOST Then
90                        Call DSBuffers(i).Buffer.Stop
100                   End If
110               End If
120           Next i
130       End If
End Sub

''
' Retrieves wether there are sounds currentyl playing or not.
'
' @return   true if there is at least one sound playing, false otherwise.

Public Property Get PlayingSound() As Boolean
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Returns true if there is at least one sound playing
      '***************************************************
          Dim i As Long
          
10        For i = 1 To NumSoundBuffers
20            If Not DSBuffers(i).Buffer Is Nothing Then
30                If DSBuffers(i).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
40                    PlayingSound = True
50                    Exit Property
60                End If
70            End If
80        Next i
End Property

''
' Sets the volume of sound.
'
' @param    Volume  New value to be set, in percentage (0 to 100).

Public Property Let SoundVolume(ByVal Volume As Byte)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Sets the sound's volume as a percentage
      '***************************************************
          Dim i As Long
          
10        If Volume > 100 Then Exit Property
          
          'Take percentage to actual value
20        SndVolume = Volume * (DSBVOLUME_MAX - DSBVOLUME_MIN) / 100 + DSBVOLUME_MIN
          
30        For i = 1 To NumSoundBuffers
40            If Not DSBuffers(i).Buffer Is Nothing Then
50                If DSBuffers(i).Buffer.GetStatus And DSBSTATUS_PLAYING Then
60                    Call DSBuffers(i).Buffer.SetVolume(SndVolume)
70                End If
80            End If
90        Next i
End Property

''
' Gets the volume of sound.
'
' @return   The current sound volume in percentage.

Public Property Get SoundVolume() As Byte
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Sets the music's volume as a percentage
      '***************************************************
10        SoundVolume = CByte((SndVolume - DSBVOLUME_MIN) * 100 / (DSBVOLUME_MAX - _
              DSBVOLUME_MIN))
End Property

''
' Plays a new MIDI file.
'
' @param    file    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    Loops   Number of times to loop sound. For eternal loops set it to -1.
'
' @return   True if file was properly loaded and started playing, False otherwise.

Public Function PlayMIDI(Optional ByVal file As String = "", Optional ByVal _
    Loops As Long = -1) As Boolean
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Loads a midi and start to play it
      '***************************************************
10    On Error GoTo ErrHandler
20        If LenB(file) > 0 Then CurrentMusicFile = file
          
30        If Not MusicActivated Then Exit Function
          
40        If Not Music Then Exit Function
          
50        If PlayingMusic Then StopMidi
          
60        If LenB(file) > 0 Then
70            If Not LoadMidi(file) Then Exit Function
80        Else
              'Make sure we have a loaded segment
90            If Segment Is Nothing Then Exit Function
100       End If
          
110       MusicMP3Stop
          
          'Play it
120       Call Segment.SetRepeats(Loops)
          
130       Set SegState = Performance.PlaySegment(Segment, 0, 0)
          
140       PlayMIDI = True
150   Exit Function

ErrHandler:
End Function

''
' Loads a new MIDI file.
'
' @param    file    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
'
' @return   True if file was properly loaded and started playing, False otherwise.

Private Function LoadMidi(ByVal file As String) As Boolean
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Loads a midi to memory
      '***************************************************
10    On Error GoTo ErrHandler
20        If Not FileExist(MusicPath & file, vbArchive) Then Exit Function
          
30        Call StopMidi
          
          'Destroy old object
40        Set Segment = Nothing
          
50        Set Segment = Loader.LoadSegment(MusicPath & file)
          
60        If Segment Is Nothing Then
70            Exit Function
80        End If
          
90        Call Segment.SetStandardMidiFile
          
100       LoadMidi = True
110   Exit Function

ErrHandler:
End Function

''
' Stops playing the currently loaded MIDI file.

Public Sub StopMidi()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Stops playing a midi
      '***************************************************
10    On Error GoTo ErrHandler
20        If PlayingMusic Then
30            Call Performance.Stop(Segment, SegState, 0, 0)
40        End If
50    Exit Sub

ErrHandler:
End Sub

''
' Sets the music's tempo.
'
' @param    Tempo   New Tempo value. Must be in the range of 0.25 to 2.0

Public Property Let MusicTempo(ByVal Tempo As Single)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Sets the music's tempo
      '***************************************************
10        If Not Music Then Exit Property
          
20        If Tempo < 0.25 Or Tempo > 2# Then Exit Property
          
30        Call Performance.SetMasterTempo(Tempo)
End Property

''
' Sets the music's volume.
'
' @param    Volume  Percentage to wich to set the music volume.

Public Property Let MusicVolume(ByVal Volume As Long)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Sets the music's volume as a percentage
      '***************************************************
10        If Not Music Then Exit Property
          
20        If Volume < 0 Or Volume > 100 Then Exit Property
          
          ' Volume ranges from -10000 to 10000
30        Call Performance.SetMasterVolume(Volume * 200 - 10000)
End Property

''
' Gets the music's tempo.
'
' @return   Current music'tempo.

Public Property Get MusicTempo() As Single
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Returns the music's tempo
      '***************************************************
10        If Not Music Then Exit Property
          
20        MusicTempo = Performance.GetMasterTempo()
End Property

''
' Gets the music's volume.
'
' @return   Current music's volume as a percentage.

Public Property Get MusicVolume() As Long
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Returns the music's volume as a percentage
      '***************************************************
10        If Not Music Then Exit Property
          
20        MusicVolume = (Performance.GetMasterVolume() + 10000) / 200
End Property

''
' Checks wether there is music currently playing.
'
' @return   True if there is a MIDI playing, false otherwise.

Public Property Get PlayingMusic() As Boolean
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Returns True if a midi is being played
      '***************************************************
10        If Not Music Or Segment Is Nothing Then Exit Property
          
20        PlayingMusic = Performance.IsPlaying(Segment, SegState)
End Property

''
' Retrieves the music's length.
'
' @return   Current music's length.

Public Property Get MusicLenght() As Long
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 16/03/04
      'Returns the lenght of the current midi
      '***************************************************
10        If Not Music Then Exit Property
          
20        MusicLenght = Segment.GetLength()
End Property

''
' Destroy the MP3 objects.
'
' @return   True If the objects were destroy.

Public Function MusicMP3Empty() As Boolean
      '**************************************
      '* Created by: Fredrik Alexandersson  *
      '* Modify Date: 20/5/2003             *
      '**************************************
10        On Error GoTo ErrorHandl
          
20        If ObjPtr(IMC) > 0 Then
30            Call MusicMP3Stop
40        End If
          
50        Set IBA = Nothing
60        Set IME = Nothing
70        Set IMPos = Nothing
80        Set IMC = Nothing
          
90        MusicMP3Empty = True
100       Exit Function

ErrorHandl:
110       MusicMP3Empty = False
120       End
End Function

''
' Loads a new MP3 file.
'
' @param    FilePath    The path to the file to be loaded.
' @return   True    If file was properly loaded and started playing, False otherwise.

Public Function MusicMP3Load(ByVal FilePath As String, Optional ByVal Volume As _
    Long = 0, Optional ByVal Balance As Long = 0) As Boolean '**** Loads a MP3 *****
      '**************************************
      '* Created by: Fredrik Alexandersson  *
      '* Modify Date: 20/5/2003             *
      '**************************************
10        On Error GoTo ErrorHandl
          
20        If Not FileExist(FilePath, vbArchive) Then Exit Function
          
30        If Not MusicMP3Empty() = True Then GoTo ErrorHandl
          
40        Set IMC = New FilgraphManager
50        Call IMC.RenderFile(FilePath)
          
60        Set IBA = IMC
          
70        If SoundVolume > 0 Then SoundVolume = 0
80        If Volume < -4000 Then Volume = -4000
          
90        IBA.Volume = Volume
          
100       If Balance > 5000 Then Balance = 5000
110       If Balance < -5000 Then Balance = -5000
          
120       IBA.Balance = Balance
          
130       Set IME = IMC
          
140       Set IMPos = IMC
150       If ObjPtr(IMPos) Then IMPos.Rate = 1#
160       IMPos.CurrentPosition = 0
          
170       MusicMP3Load = True
180       Exit Function

ErrorHandl:
190       MusicMP3Load = False
End Function

''
' Plays a new MP3 file.
'
' @param    FilePath    The path to the file to be loaded. If the path is null then play the loaded MP3.
'
' @return   True    If file was properly loaded and started playing, False otherwise.

Public Function MusicMP3Play(Optional ByVal FilePath As String = "") As Boolean
      '**************************************
      '* Created by: Fredrik Alexandersson  *
      '* Modify Date: 20/5/2003             *
      '**************************************
10        On Error GoTo ErrorHandl
          
20        If Not MusicActivated Then Exit Function
          
30        If Not Music Then Exit Function
          
40        If LenB(FilePath) > 0 Then
50            If Not MusicMP3Load(FilePath) Then Exit Function
60        Else
70            If IMC Is Nothing Then Exit Function
80        End If
          
90        StopMidi
100       Call IMC.Run
          
110       MusicMP3Play = True
120       Exit Function

ErrorHandl:
130       MusicMP3Play = False
End Function

''
' Check if the song need looping.

Public Sub MusicMP3GetLoop()
10        On Error GoTo ErrorHandl
          
20        If IMPos.StopTime = IMPos.CurrentPosition Then
30            MusicMP3Stop
40            MusicMP3Play
50        End If
ErrorHandl:
End Sub

''
' Stop a the current MP3 file.
'
' @return   True    If file was properly stop, False otherwise.

Public Function MusicMP3Stop() As Boolean
      '**************************************
      '* Created by: Fredrik Alexandersson  *
      '* Modify Date: 20/5/2003             *
      '**************************************
10        On Error GoTo ErrorHandl
          
20        IMC.Stop
30        IMPos.CurrentPosition = 0
          
40        MusicMP3Stop = True
50        Exit Function

ErrorHandl:
60        MusicMP3Stop = False
End Function

''
' Pause a the current MP3 file.
'
' @return   True    If file was properly pause, False otherwise.

Public Function MusicMP3Pause() As Boolean
      '***************************************************
      ' Created by: Juan Martín Sotuyo Dodero
      ' Modify Date: 25/6/2004
      '***************************************************
10        On Error GoTo ErrorHandl
          
20        Call IMC.Pause
          
30        MusicMP3Pause = True
40        Exit Function

ErrorHandl:
50        MusicMP3Pause = False
End Function

''
' Set the volume to the current MP3.
'
' @param    SoundVolume The new volume to the current MP3.
' @return   True    If the volume was properly change, False otherwise.

Public Function MusicMP3VolumeSet(ByVal SoundVolume As Long) As Boolean
      '**************************************
      '* Created by: Fredrik Alexandersson  *
      '* Modify Date: 20/5/2003             *
      '**************************************
10        On Error GoTo ErrorHandl
          
20        If SoundVolume > 0 Then SoundVolume = 0
30        If SoundVolume < -4000 Then SoundVolume = -4000
          
40        IBA.Volume = SoundVolume

50        MusicMP3VolumeSet = True
60        Exit Function

ErrorHandl:
70        MusicMP3VolumeSet = False
End Function

''
' Retrieves wether music is activated or not.
'
' @return   True if music is activated, false otherwise.

Public Property Get MusicActivated() As Boolean
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 03/03/07
      'Retrieves wether the music is activated or not
      '***************************************************
10        MusicActivated = mActivated
End Property

''
' Sets wether music is activated or not.
'
' @param    nValue  True if music should be activated, false otherwise.

Public Property Let MusicActivated(ByVal nValue As Boolean)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 03/03/07
      'Sets wether the music is activated or not
      '***************************************************
10        If mActivated = nValue Then Exit Property
          
20        mActivated = nValue
          
30        If nValue Then
40            Call PlayMIDI(CurrentMusicFile)
50        Else
60            Call StopMidi
70        End If
End Property

''
' Retrieves wether sound is activated or not.
'
' @return   True if sound is activated, false otherwise.

Public Property Get SoundActivated() As Boolean
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 03/03/07
      'Retrieves wether the sound is activated or not
      '***************************************************
10        SoundActivated = sActivated
End Property

''
' Sets wether sound is activated or not.
'
' @param    nValue  True if sound should be activated, false otherwise.

Public Property Let SoundActivated(ByVal nValue As Boolean)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 03/03/07
      'Sets wether the sound is activated or not
      '***************************************************
10        If sActivated = nValue Then Exit Property
          
20        sActivated = nValue
          
30        If Not nValue Then
40            Call StopWave
50        End If
End Property

''
' Sets wether sound effects are activated or not.
'
' @param nValue True if sound effects should be activated, false otherwise.

Public Property Let SoundEffectsActivated(ByVal nValue As Boolean)
      '***************************************************
      'Author: Franco Zeoli (Noich)
      'Last Modification: 03/04/10
      'Sets wether the sound effects are activated or not
      '03/04/10 - Pato: If sound effects are disable, then reset the frequency, volume and pan.
      '***************************************************
10        sEffectsActivated = nValue
          
20        If Not nValue Then
30            Call ResetSoundValues
40        End If
End Property

''
' Retrieves wether sound effects are activated or not.
'
' @return   True if sound effects are activated, false otherwise.

Public Property Get SoundEffectsActivated() As Boolean
      '***************************************************
      'Author: Franco Zeoli (Noich)
      'Last Modification: 28/07/08
      'Retrieves wether the sound effects are activated or not
      '***************************************************
10        SoundEffectsActivated = sEffectsActivated
End Property
