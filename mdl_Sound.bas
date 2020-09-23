Attribute VB_Name = "mdl_Sound"
'OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
' You are free to use this file as you want
' It gives you high quality stereo sound support, easy to use
' and directsound powered.
' Each buffer supports frequency, pan and volume manipulation.
' CAME FROM:
' Lucky
' theluckyleper@home.com
' http://members.home.net/theluckyleper
'OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO


'DirectX Variables
Dim ds As DirectSound

'User defined type to determine a buffer's capabilities
Private Type BufferCaps
    Volume As Boolean               'Can this buffer's volume be changed?
    Frequency As Boolean            'Can the frequency be altered?
    Pan As Boolean                  'Can we pan the sound from left to right?
    Loop As Boolean                 'Is this sound looping?
    Delete As Boolean               'Should this sound be deleted after playing?
End Type

'User defined type to contain sound data
Private Type SoundArray
    DSBuffer As DirectSoundBuffer   'The buffer that contains the sound
    DSState As String               'Describes the current state of the buffer (ie. "Playing", "Stopped")
    DSNotification As Long          'Contains the event reference returned by the DirectX7 object
    DSCaps As BufferCaps            'Describes the buffer's capabilities
    DSSourceName As String          'The name of the source file
    DSFile As Boolean               'Is the source in a seperate file?
    DSResource As Boolean           'Or is it in a resource?
    DSEmpty As Boolean              'Is this SoundArray index empty?
End Type
Dim Sound() As SoundArray          'Contains all the data needed for sound manipulation

'Constant that contains the path inside the app.path in which the sounds are stored
Const DataLocation = "\"

'Wave Format Setting Contants
Const NumChannels = 2              'How many channels will we be playing on?
Const SamplesPerSecond = 22050     'How many cycles per second (hertz)?
Const BitsPerSample = 16           'What bit-depth will we use?

Public Sub Initialize(ByRef Handle As Long)

    'If we can't initialize properly, trap the error
    On Local Error GoTo ErrOut

    'Make the DirectSound object
    Set ds = frm_Game.dx.DirectSoundCreate("")
    
    'Set the DirectSound object's cooperative level (Priority gives us sole control)
    ds.SetCooperativeLevel Handle, DSSCL_PRIORITY
    
    'Initialize our Sound array to zero
    ReDim Sound(0)
    Sound(0).DSEmpty = True
    Sound(0).DSState = "empty"
    
    'Exit sub before the error code
    Exit Sub
    
ErrOut:
    'Display an error message and exit if initialization failed
    MsgBox "Unable to initialize DirectSound."
    End

End Sub

Public Function LoadSound(SourceName As String, IsFile As Boolean, IsResource As Boolean, IsDelete As Boolean, IsFrequency As Boolean, IsPan As Boolean, IsVolume As Boolean, IsLoop As Boolean, FormObject As Form) As Integer

Dim i As Integer
Dim Index As Integer
Dim DSBufferDescription As DSBUFFERDESC
Dim DSFormat As WAVEFORMATEX
Dim DSPosition(0) As DSBPOSITIONNOTIFY

    'Search the sound array for any empty spaces
    Index = -1
    For i = 0 To UBound(Sound)
        If Sound(i).DSEmpty = True Then 'If there is an empty space, us it
            Index = i
            Exit For
        End If
    Next
    If Index = -1 Then                  'If there's no empty space, make a new spot
        ReDim Preserve Sound(UBound(Sound) + 1)
        Index = UBound(Sound)
    End If
    LoadSound = Index                   'Set the return value of the function
    
    'Load the Sound array with the data given
    With Sound(Index)
        .DSEmpty = False                'This Sound(index) is now occupied with data
        .DSFile = IsFile                'Is this sound to be loaded from a file?
        .DSResource = IsResource        'Or is it to be loaded from a resource?
        .DSSourceName = SourceName      'What is the name of the source?
        .DSState = "Stopped"            'Set the current state to "Stopped"
        .DSCaps.Delete = IsDelete       'Is this sound to be deleted after it is played?
        .DSCaps.Frequency = IsFrequency 'Is this sound to have frequency altering capabilities?
        .DSCaps.Loop = IsLoop           'Is this sound to be looped?
        .DSCaps.Pan = IsPan             'Is this sound to have Left and Right panning capabilities?
        .DSCaps.Volume = IsVolume       'Is this sound capable of altered volume settings?
    End With
    
    'Set the buffer description according to the data provided
    With DSBufferDescription
        If Sound(Index).DSCaps.Delete = True Then .lFlags = .lFlags Or DSBCAPS_CTRLPOSITIONNOTIFY
        If Sound(Index).DSCaps.Frequency = True Then .lFlags = .lFlags Or DSBCAPS_CTRLFREQUENCY
        If Sound(Index).DSCaps.Pan = True Then .lFlags = .lFlags Or DSBCAPS_CTRLPAN
        If Sound(Index).DSCaps.Volume = True Then .lFlags = .lFlags Or DSBCAPS_CTRLVOLUME
    End With

    'Set the Wave Format
    With DSFormat
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = NumChannels
        .lSamplesPerSec = SamplesPerSecond
        .nBitsPerSample = BitsPerSample
        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
    End With
    
    'Load the sound into the buffer
    If Sound(Index).DSFile = True Then          'If it's in a file...
        Set Sound(Index).DSBuffer = ds.CreateSoundBufferFromFile(App.Path & DataLocation & Sound(Index).DSSourceName, DSBufferDescription, DSFormat)
    ElseIf Sound(Index).DSResource = True Then  'If it's in a resource...
        Set Sound(Index).DSBuffer = ds.CreateSoundBufferFromResource("", Sound(Index).DSSourceName, DSBufferDescription, DSFormat)
    End If
    
    'If the sound is to be deleted after it plays, we must create an event for it
    If Sound(Index).DSCaps.Delete = True Then
        Sound(Index).DSNotification = frm_Game.dx.CreateEvent(FormObject)        'Make the event (has to be created in a Form Object) and get its handle
        DSPosition(0).hEventNotify = Sound(Index).DSNotification        'Place this event handle in an DSBPOSITIONNOTIFY variable
        DSPosition(0).lOffset = DSBPN_OFFSETSTOP                        'Define the position within the wave file at which you would like the event to be triggered
        Sound(Index).DSBuffer.SetNotificationPositions 1, DSPosition()  'Set the "notification position" by passing the DSBPOSITIONNOTIFY variable
    End If
    
End Function

Public Sub RemoveSound(Index As Integer)

    'Destroy the event associated with the ending of this sound, if there was one
    If Sound(Index).DSCaps.Delete = True And Sound(Index).DSNotification <> 0 Then frm_Game.dx.DestroyEvent Sound(Index).DSNotification
    
    'Reset all the variables in the sound array
    With Sound(Index)
        Set .DSBuffer = Nothing
        .DSCaps.Delete = False
        .DSCaps.Frequency = False
        .DSCaps.Loop = False
        .DSCaps.Pan = False
        .DSCaps.Volume = False
        .DSEmpty = True
        .DSFile = False
        .DSNotification = 0
        .DSResource = False
        .DSSourceName = ""
        .DSState = "empty"
    End With
        
End Sub

Public Sub PlaySound(Index As Integer)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'If the sound is not "paused" then reset it's position to the beginning
    If Sound(Index).DSState <> "paused" Then Sound(Index).DSBuffer.SetCurrentPosition 0
    
    'Play looped or singly, as appropriate
    If Sound(Index).DSCaps.Loop = False Then Sound(Index).DSBuffer.Play DSBPLAY_DEFAULT
    If Sound(Index).DSCaps.Loop = True Then Sound(Index).DSBuffer.Play DSBPLAY_LOOPING
    
    'Set the state to "playing"
    Sound(Index).DSState = "playing"

End Sub

Public Sub StopSound(Index As Integer)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'Stop the buffer and reset to the beginning
    Sound(Index).DSBuffer.Stop
    Sound(Index).DSBuffer.SetCurrentPosition 0
    Sound(Index).DSState = "stopped"

End Sub

Public Sub PauseSound(Index As Integer)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'Stop the buffer
    Sound(Index).DSBuffer.Stop
    Sound(Index).DSState = "paused"

End Sub

Public Sub SetFrequency(Index As Integer, Freq As Long)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'Check to make sure that the buffer has the capability of altering its frequency
    If Sound(Index).DSCaps.Frequency = False Then Exit Sub

    'Alter the frequency according to the Freq provided
    Sound(Index).DSBuffer.SetFrequency Freq

End Sub

Public Sub SetVolume(Index As Integer, Vol As Long)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'Check to make sure that the buffer has the capability of altering its volume
    If Sound(Index).DSCaps.Volume = False Then Exit Sub

    'Alter the volume according to the Vol provided
    Sound(Index).DSBuffer.SetVolume Vol

End Sub

Public Sub SetPan(Index As Integer, Pan As Long)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'Check to make sure that the buffer has the capability of altering its pan
    If Sound(Index).DSCaps.Pan = False Then Exit Sub

    'Alter the pan according to the Pan provided
    Sound(Index).DSBuffer.SetPan Pan

End Sub

Public Function GetFrequency(Index As Integer) As Long

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Function
    
    'Check to make sure that the buffer has the capability of altering its frequency
    If Sound(Index).DSCaps.Frequency = False Then Exit Function
    
    'Return the frequency value
    GetFrequency = Sound(Index).DSBuffer.GetFrequency()

End Function

Public Function GetVolume(Index As Integer) As Long

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Function
    
    'Check to make sure that the buffer has the capability of altering its volume
    If Sound(Index).DSCaps.Volume = False Then Exit Function
    
    'Return the volume value
    GetVolume = Sound(Index).DSBuffer.GetVolume()

End Function

Public Function GetPan(Index As Integer) As Long

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Function
    
    'Check to make sure that the buffer has the capability of altering its pan
    If Sound(Index).DSCaps.Pan = False Then Exit Function
    
    'Return the pan value
    GetPan = Sound(Index).DSBuffer.GetPan()

End Function

Public Function GetState(Index As Integer) As String

    'Returns the current state of the given sound
    GetState = Sound(Index).DSState

End Function

Public Function DXCallback(ByVal eventid As Long) As Integer

Dim i As Integer
    
    'Find the sound that caused this event to be triggered
    For i = 0 To UBound(Sound)
        If Sound(i).DSNotification = eventid Then
            Exit For
        End If
    Next
    
    'Return the ID
    DXCallback = i

End Function

Public Sub Terminate()

Dim i As Integer

    'Delete all of the sounds created
    For i = 0 To UBound(Sound)
        RemoveSound i
    Next

End Sub

