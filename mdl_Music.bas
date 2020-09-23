Attribute VB_Name = "mdl_Music"
'OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
' You are free to use this file as you want
'OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO

Option Explicit

Dim mPerf As DirectMusicPerformance
Dim mSeg As DirectMusicSegment
Dim mLoader As DirectMusicLoader

Public Sub Initialize()
    Dim x As Integer
    Dim udtPortCaps As DMUS_PORTCAPS
    'Initialize
    Set mLoader = frm_Game.dx.DirectMusicLoaderCreate()
    Set mPerf = frm_Game.dx.DirectMusicPerformanceCreate()
    mPerf.Init Nothing, 0
    'Get the port
    For x = 1 To mPerf.GetPortCount
        Call mPerf.GetPortCaps(x, udtPortCaps)
        If udtPortCaps.lFlags And DMUS_PC_SHAREABLE Then
            mPerf.SetPort x, 1
            Exit For
        End If
    Next x
End Sub

Public Sub Play(strSong As String)
    'Load and play the file
    Set mSeg = mLoader.LoadSegment(App.Path & "\" & strSong)
    mSeg.SetLoopPoints 0, mSeg.GetLength - 3000
    mSeg.SetRepeats 99
    mPerf.PlaySegment mSeg, 0, 0
End Sub

Public Sub Terminate()
    Set mSeg = Nothing
    mPerf.CloseDown
    Set mPerf = Nothing
    Set mLoader = Nothing
End Sub

Public Sub StopMusic()
    mPerf.Stop mSeg, Nothing, 0, 0
End Sub
