VERSION 5.00
Begin VB.Form frm_Game 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Mario Game"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frm_Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This game was written ages ago :), I dont have time for it now, you can
' modify it, make levels, etc. as you like.
' Levels are made in MS Paint, but converted to ascii data with my program
'OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
' Soldier Of Misfortune
' Game author: Sala Bojan, (C) Hallsoft 2002-2003
' You are free to use this code as you like, it is made for
' you to learn something, it will teach you how to make good
' 2d smooth side-scrooling platform games.
' If you need help, contact me.
' BUT, before you start making some "game", first you must
' visit this site: http://members.home.net/theluckyleper !!
' there you will find lots of tutorials, files, dx games, etc.
' IMPORTANT: ####################################################
' You might need to change the game speed, open the "config.som"
' file and there you will see this: "15".
' For slow computers: Change it between "1" and "15" to improve
' the game speed, choose the speed you want.
' For fast computers: Change it between "15" and "40" to slow it down.
'OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO

' If you want even better 2d games (this code is pretty old, and badly written :), VOTE!
' That is all I can say, happy coding...

Option Explicit

' Basic directx stuff
Public dx As DirectX7 ' Direct[all together], as name says...
Public dD As DirectDraw7 ' To draw
Public dI As DirectInput ' To use keyboard, mouse, etc. , MUCH better then standard windows input

Private dGC As DirectDrawGammaControl
Private drOriginal As DDGAMMARAMP
Private drRamp As DDGAMMARAMP

Private dCaps As DDSCAPS2 ' Screen capabillities, dx must know what do we need, nice graphic or bad...
Private dFront As DDSURFACEDESC2 ' Description of the front surface, primary drawing surface
Private dScreen As DDSURFACEDESC2 ' This is the same, but for a screen (what we see on the monitor)

Private dsFront As DirectDrawSurface7 ' Front surface, this is something like a large bitmap in the memory
Private dsBack As DirectDrawSurface7 ' Back surface, here we put all our gfx, text, etc.

Private diKeyboard As DirectInputDevice ' Just keyboard for now
Private diKeyState As DIKEYBOARDSTATE ' What button :)

Private dFont As New StdFont ' Just font for printing
Private wFont As New StdFont

'User defined type to determine a buffer's capabilities
Private Type BufferCaps
    Volume As Boolean               'Can this buffer's volume be changed?
    Frequency As Boolean            'Can the frequency be altered?
    Pan As Boolean                  'Can we pan the sound from left to right?
    Loop As Boolean                 'Is this sound looping?
    Delete As Boolean               'Should this sound be deleted after playing?
End Type

' Monitor
Private Type tScreen
    x As Long ' Right
    Y As Long ' Bottom
    Bits As Long
    Refreshr As Long
End Type
Private dRes As tScreen

' Sprite
Private Type tSprite
    Surface As Long
    IsCreated As Boolean ' :)
    Pos As RECT ' Just position (and size, right and bottom...)
    Shape As RECT ' We may need to use animations, or tu cut the picture, so this is a shape of it
    Visible As Boolean
    Width As Long ' Just for some infos, surface, etc..
    Height As Long ' Same
    IsBackBitmap As Boolean ' If it is just a picture
    IsHero As Boolean
    OffScreen As Boolean ' If it is leaving the screen
    IsJumping As Boolean
    JumpAgain As Boolean
    IsFalling As Boolean
    IsGroundSprite As Boolean
    CanCollide As Boolean
    Enemy As Long ' If it is enemy
    vDirection As Long
    vDirect As Boolean
    MovingFrameDelay As Long
    eAnimationDelay As Long
    IsDead As Boolean
    eSpeed As Long
    eFrames As Long
    eCurFrame As Long
    eTolerance As Long
    eCanRotate As Boolean
    Item As Long
    IsConstant As Boolean
    sType As Long
    IsShowed As Boolean
    IsSpiky As Boolean
    IsLevelDone As Boolean
End Type
Private Sprite(8000) As tSprite
Private Sprites As Long

Private Type tSurf
    FileName As String
    Width As Long
    Height As Long
    Surf As DirectDrawSurface7 ' Surface for directdraw
    Description As DDSURFACEDESC2 ' Surface description
    ColorKey As DDCOLORKEY ' Surface color key, for transparency
End Type
Private Surf(99) As tSurf
Private Surfaces As Long

' Some game things
Private gRunning As Boolean ' Is the game running, if we turn this off, it will end
Private gFPSC As Long
Private gFps As Long
Private gSleepTick As Long
Private gMsTimer As Long

Private gSpeed As Long
Private gJumpSpeed As Long
Private gJumpMaxFactor As Long
Private gGravityFactor As Long
Private gGameDelay As Single

Private gScore As Long
Private gGreens As Long
Private gGreenCount As Long
Private gLevel As Long
Private gLives As Long
Private gWarning As Long

Private gMenu As Boolean
Private gAbout As Boolean
Private gArtist As Boolean

Private Const aHero = 2
Private Const aStart = 3

' This is for getting collision between two rects, not pixel collision (not yet ;)
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function GetTickCount& Lib "kernel32" ()

Public i As Long ' :)
Public C As Long


Private Sub Form_Load()
    MsgBox "Please compile the game (remove this msgbox :), and read README.TXT file!"
    
    dRes.x = 320 ' Width
    dRes.Y = 200 ' Height
    dRes.Bits = 16 ' Many, many Colors, use 8-bit if you have weak pc
    dRes.Refreshr = 0 ' Default, about 60-80hz (WARNING: You must set this if the screen is "messy")
    dFont.Name = "Arial"
    dFont.Size = 8
    dFont.Bold = True
    wFont.Name = "Arial"
    wFont.Size = 12
    wFont.Bold = True
    Me.Show
    dRaise
    gSpeed = 2
    gJumpSpeed = 6
    gJumpMaxFactor = 15
    gGravityFactor = 3
    Open App.Path & "\config.som" For Input As #1
        Input #1, gGameDelay
    Close #1
    gLives = 3
    gLevel = 0
    gLoadSounds
    gMenu = True
    dMainLoop
End Sub

Public Sub dRaise()
    
    gRunning = True ' Game is running

    Set dx = New DirectX7 ' Create directx
    Set dD = dx.DirectDrawCreate("") ' Create directdraw7 engine
    
    dD.SetCooperativeLevel Me.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN ' Set standard fullscreen mode
    ' It will run in 320x240x8 mode
    dD.SetDisplayMode dRes.x, dRes.Y, dRes.Bits, dRes.Refreshr, DDSDM_DEFAULT   ' Set display resolution, color mode, and refresh rate (0 for default)
    
    ' Now  I will describe the front buffer, this is the PRIMARY buffer !
    With dFront
        .lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT ' Set the flags we will use
        .lBackBufferCount = 1 ' One back, and one front buffer for double buffering of course
        .ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE ' It will be primary, it will flip, and its complex :)
    End With
    Set dsFront = dD.CreateSurface(dFront) ' Create surface with its description

    ' Here you must create back buffer, it holds the screen
    dCaps.lCaps = DDSCAPS_BACKBUFFER ' What it is
    Set dsBack = dsFront.GetAttachedSurface(dCaps)
    dsBack.GetSurfaceDesc dScreen ' Get screen description
    
    ' Init the KeyBoard
    Set dI = dx.DirectInputCreate
    Set diKeyboard = dI.CreateDevice("GUID_SysKeyboard")
    diKeyboard.SetCommonDataFormat DIFORMAT_KEYBOARD
    diKeyboard.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE

    ' Init the gama control
    Set dGC = dsFront.GetDirectDrawGammaControl
    dGC.GetGammaRamp DDSGR_DEFAULT, drOriginal
    
    ' Init sound module
    mdl_Sound.Initialize Me.hWnd
    
    'Init music (midi)
    mdl_Music.Initialize
End Sub


Public Sub dTerminate()
    dD.RestoreDisplayMode
    dD.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
    ' Erase all the things
    Set dD = Nothing
    Set dx = Nothing
    Set dI = Nothing
    
    For i = 1 To Surfaces
        Set Surf(i).Surf = Nothing ' Remove the surfaces
    Next i
    
    mdl_Sound.Terminate
    
    mdl_Music.Terminate
     
    Unload Me ' Unload the form
End Sub

Public Sub dMainLoop()
    ' This is the main sub, it will loop the code until you press ESC, but it is
    ' activating events (doevents), so the cpu is not focused only on this code.
    ' to improve the performance, just turn off DoEvents line, but it may crash or hang.
g_start:
    Dim rScreen As RECT, rRes As RECT, rLeft As RECT, rRight As RECT, rBottom As RECT, rTop As RECT, reHor As RECT, reVer As RECT, reHit As RECT, rHit As RECT
    Dim T1&, T2&, T3&, MV&, MVDIR&, MJUMP&, MACC!, MJACC&, MJACE&, MSPACE&, cR%, cG%, cB%, gJustStarted As Boolean, SG%, mPos&, gUp&, gDown&, gEsc&
    If gMenu Then
        gLoadLevel App.Path & "\menu.dat", True
        mdl_Music.StopMusic
    Else
        Select Case gLevel
        Case 0
            gLoadLevel App.Path & "\level0.dat", False
        Case 1
            gLoadLevel App.Path & "\level1.dat", False
        Case 2
             gLoadLevel App.Path & "\level2.dat", False
        End Select
        mdl_Music.Play "level" & gLevel & ".mid"
        gJustStarted = True
    End If
    SG = -99
    gMsTimer = GetTickCount
    Do While gRunning ' Start the game
        If gMsTimer + gGameDelay <= GetTickCount Then
            gMsTimer = GetTickCount ' I will two counters, one that will slowdown and one that will just show the fps
            If T3 >= 1000 Then ' We have one second
                gFps = gFPSC ' Set the fps same as fps counter
                gFPSC = 0: T3 = 0 ' Reset
            End If
            T1 = GetTickCount ' Get the first tick, this is for fps, it may look strange but it is  safer (I have 400mhz computer :)
            
            '// Now start the game ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            dsBack.BltColorFill rScreen, 0   ' Fill the screen with blackness, clear it
            diKeyboard.Acquire ' Enable the keyboard
            diKeyboard.GetDeviceStateKeyboard diKeyState ' Trap the keys
            If diKeyState.Key(DIK_ESCAPE) Then
                If gEsc = 0 Then
                    If gMenu Then
                        If gAbout Then
                            gAbout = False
                        Else
                            gRunning = False
                        End If
                    Else
                        gMenu = True: gLevel = 0: GoTo g_start ' If escape pressed
                    End If
                End If
                gEsc = 1
            Else
                gEsc = 0
            End If
            If gMenu Then
                If diKeyState.Key(DIK_SPACE) Then
                    Select Case mPos
                    Case 0
                        gMenu = False
                        gLevel = 0
                        GoTo g_start
                    Case 1
                        If Not gAbout Then gAbout = True: gArtist = True ' :)
                    Case 2
                        If Not gAbout Then gAbout = True: gArtist = False
                    Case 3
                        gRunning = False
                    End Select
                End If
                If Not gAbout Then
                    If diKeyState.Key(DIK_UP) Then
                        If gUp = 0 Then If mPos > 0 Then mPos = mPos - 1: mdl_Sound.PlaySound 3
                        gUp = 1
                    Else
                        gUp = 0
                    End If
                    If diKeyState.Key(DIK_DOWN) Then
                        If gDown = 0 Then If mPos < 3 Then mPos = mPos + 1: mdl_Sound.PlaySound 3
                        gDown = 1
                    Else
                        gDown = 0
                    End If
                End If
                With dsBack
                    dFont.Size = 13
                    dFont.Name = "Arial Black"
                    .SetForeColor vbWhite
                    .SetFont dFont
                    .DrawText 23, 16, "SOLDIER OF MISFORTUNE", False
                    
                    .SetForeColor vbBlue
                    dFont.Name = "Arial"
                    If gAbout Then
                        Dim gA&: gA = 0
                        dFont.Size = 8
                        .SetFont dFont
                        .SetForeColor vbGreen
                        If gArtist Then
                            .DrawText 20, 50 + gA, "This game could be better and much larger if I had ", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "some more graphics. So, if you can draw graphics", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "just like in this game, and you want to create a", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "game, then contact me, and we can create cool 2d", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "games as long as you can draw :).", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "Almost all the bitmaps in this game came from:", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "http://www.arifeldman.com", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "", False: gA = gA + 14
                        Else
                            .DrawText 20, 50 + gA, "Game is made by Sala Bojan with DirectX 7, in      ", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "Visual Basic 6.0. Graphics are done by Ari Feldman.", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "If you need help, or you want to report a bug, then", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "you can send me an email to alas@eunet.yu, or visit", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "www.hallsoft.tk or www.univerzalsoft.com.          ", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "If you liked the game then you could give me some  ", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "support, you dont have to pay, just visit the site:", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "www.planetsourcecode.com/vb, enter the title of    ", False: gA = gA + 14
                            .DrawText 20, 50 + gA, "this game in the search box, and VOTE for it !     ", False: gA = gA + 14
                        End If
                    Else
                        dFont.Size = 8
                        .SetFont dFont
                        .SetForeColor vbWhite
                        .DrawText dRes.x - 128, dRes.Y - 20, "© Hallsoft/Sala Bojan 2003", False
                        dFont.Size = 12
                        .SetFont dFont
                        .SetForeColor vbBlue
                        .DrawText 100, 64, "Start game !", False
                        .DrawText 100, 88, "For artists", False
                        .DrawText 100, 112, "About the game", False
                        .DrawText 100, 136, "Exit", False
                        .SetForeColor vbRed
                        .setDrawWidth 2
                        .DrawRoundedBox 80, (64 + (24 * mPos)) - 1, 240, (88 + (24 * mPos)) - 3, 32, 32
                    End If
                End With
            Else
                '// Now make the gravity force
                sPosition aHero, Sprite(aHero).Pos.Left, Sprite(aHero).Pos.Top + gGravityFactor
                Sprite(aHero).IsFalling = True
                
                '// Hero Collision part
                rLeft.Left = Sprite(aHero).Pos.Left: rLeft.Right = Sprite(aHero).Pos.Right - 20
                rLeft.Top = Sprite(aHero).Pos.Top + 6: rLeft.Bottom = Sprite(aHero).Pos.Bottom - 6
                rRight.Left = Sprite(aHero).Pos.Left + 20: rRight.Right = Sprite(aHero).Pos.Right
                rRight.Top = Sprite(aHero).Pos.Top + 6: rRight.Bottom = Sprite(aHero).Pos.Bottom - 6
                rBottom.Left = Sprite(aHero).Pos.Left + 6: rBottom.Right = Sprite(aHero).Pos.Right - 6
                rBottom.Top = Sprite(aHero).Pos.Top + 20: rBottom.Bottom = Sprite(aHero).Pos.Bottom
                rTop.Left = Sprite(aHero).Pos.Left + 6: rTop.Right = Sprite(aHero).Pos.Right - 6
                rTop.Top = Sprite(aHero).Pos.Top: rTop.Bottom = Sprite(aHero).Pos.Bottom - 20
                rHit.Left = Sprite(aHero).Pos.Left: rHit.Right = Sprite(aHero).Pos.Right
                rHit.Top = Sprite(aHero).Pos.Top + 7: rHit.Bottom = Sprite(aHero).Pos.Bottom - 7
                
                For i = aStart To Sprites
                    If Sprite(i).Visible = True Then
                        If Sprite(i).Pos.Left < dRes.x And Sprite(i).Pos.Right > 0 Then
                            ' This is "sliding" collision, it will simply stick on the ground
                            If Sprite(i).IsGroundSprite Then
                                If Sprite(i).IsLevelDone Then
                                    If IntersectRect(rRes, Sprite(i).Pos, Sprite(aHero).Pos) Then
                                        '// Jump to next level, if all diamonds are present
                                        If gGreens >= gGreenCount Then
                                            For C = 0 To 99
                                                gSleep 8
                                                gSetGamma -C, -C, -C
                                            Next C
                                            MJACC = 2: MJACE = 0
                                            gLevel = gLevel + 1
                                            GoTo g_start
                                        Else
                                            gWarning = 1
                                        End If
                                    End If
                                End If
                                If Sprite(i).CanCollide Then
                                    If IntersectRect(rRes, rTop, Sprite(i).Pos) Then Sprite(aHero).IsJumping = False ': mdl_Sound.PlaySound 4
                                    If IntersectRect(rRes, rLeft, Sprite(i).Pos) Then sPosition aHero, Sprite(i).Pos.Right, Sprite(aHero).Pos.Top
                                    If IntersectRect(rRes, rRight, Sprite(i).Pos) Then sPosition aHero, Sprite(i).Pos.Left - 23, Sprite(aHero).Pos.Top
                                    If IntersectRect(rRes, rBottom, Sprite(i).Pos) Then
                                        sPosition aHero, Sprite(aHero).Pos.Left, Sprite(i).Pos.Top - Sprite(aHero).Shape.Bottom
                                        Sprite(aHero).IsJumping = False: MJUMP = 0
                                        Sprite(aHero).IsFalling = False
                                        MJACC = 2: MJACE = 0
                                        If Sprite(i).IsSpiky Then If Not Sprite(aHero).IsDead Then sKillHero
                                    End If
                                End If
                                If Sprite(aHero).Pos.Top > dRes.Y Then
                                    If Not Sprite(aHero).IsDead Then
                                        mdl_Sound.PlaySound 0
                                        mdl_Sound.PlaySound 6
                                        sKillHero
                                    End If
                                End If
                                If Sprite(aHero).IsDead Then
                                    If i = Sprites Then
                                        If Sprite(i).Pos.Bottom < 0 Then
                                            For C = 0 To 99
                                                gSleep 8
                                                gSetGamma -C, -C, -C
                                            Next C
                                            Sprite(aHero).IsDead = False
                                            Sprite(aHero).Visible = True
                                            sPosition aHero, 32, 32
                                            If gLives < 0 Then gLives = 3: gMenu = True: gLevel = 0: GoTo g_start
                                            GoTo g_start
                                        End If
                                        sPosition i, Sprite(i).Pos.Left, Sprite(i).Pos.Top - 3
                                        Sprite(i).eCurFrame = Sprite(i).eCurFrame + 1
                                        If Sprite(i).eCurFrame = 5 Then
                                            If Sprite(i).Shape.Left = 0 Then
                                                Sprite(i).Shape.Left = 30: Sprite(i).Shape.Right = 60
                                            Else
                                                Sprite(i).Shape.Left = 0: Sprite(i).Shape.Right = 30
                                            End If
                                            Sprite(i).eCurFrame = 0
                                        End If
                                    End If
                                End If
                                ' Use items
                                If Sprite(i).Item > 0 Then
                                    If Not Sprite(i).IsDead Then
                                        If Not Sprite(aHero).IsDead Then
                                            If IntersectRect(rRes, Sprite(aHero).Pos, Sprite(i).Pos) Then
                                                Sprite(i).IsDead = True
                                                mdl_Sound.PlaySound 3
                                                Select Case Sprite(i).Item
                                                    Case 1: gScore = gScore + 10: gGreens = gGreens + 1
                                                    Case 2: gScore = gScore + 35
                                                    Case 3: gScore = gScore + 12
                                                    Case 4: gScore = gScore + 16
                                                    Case 5: gScore = gScore + 18
                                                    Case 6: gScore = gScore + 22
                                                End Select
                                            End If
                                        End If
                                    End If
                                    ' Dead items
                                    If Sprite(i).IsDead Then
                                        Sprite(i).Shape.Left = 112: Sprite(i).Shape.Right = 128
                                        Sprite(i).eCurFrame = Sprite(i).eCurFrame + 1
                                        If Sprite(i).eCurFrame = 10 Then Sprite(i).Visible = False: Sprite(i).eCurFrame = 0: Sprite(i).IsDead = False
                                    End If
                                End If
                                '// Check for enemy collision, and move them
                                ' Dead enemies
                                If Sprite(i).Enemy > 0 Then
                                    If Sprite(i).IsDead Then
                                        Sprite(i).MovingFrameDelay = Sprite(i).MovingFrameDelay + 1
                                        If Sprite(i).MovingFrameDelay = 10 Then
                                            Sprite(i).MovingFrameDelay = 0
                                            Sprite(i).Visible = False
                                            sShowSprite Sprite(i).Pos.Left, Sprite(i).Pos.Top + 4, CStr(17 + Sprite(i).Enemy)
                                        End If
                                    End If
                                    ' We must use two IF statements, it is much faster
                                    If Not Sprite(i).IsDead Then
                                        If Sprite(i).Pos.Left < dRes.x Then
                                            If Sprite(i).Pos.Right > 0 Then
                                            
                                                sPosition i, Sprite(i).Pos.Left, Sprite(i).Pos.Top + 2
                                                reHor.Left = Sprite(i).Pos.Left + 3: reHor.Right = Sprite(i).Pos.Right - 3
                                                reHor.Top = Sprite(i).Pos.Top + 7: reHor.Bottom = Sprite(i).Pos.Bottom - 7
                                                reVer.Left = Sprite(i).Pos.Left + 6: reVer.Right = Sprite(i).Pos.Right - 6
                                                reVer.Top = Sprite(i).Pos.Top: reVer.Bottom = Sprite(i).Pos.Bottom
                                                reHit.Left = Sprite(i).Pos.Left: reHit.Right = Sprite(i).Pos.Right
                                                reHit.Top = Sprite(i).Pos.Top + Sprite(i).eTolerance: reHit.Bottom = Sprite(i).Pos.Bottom
                                                
                                                ' Well... out hero is dead now
                                                If IntersectRect(rRes, rHit, reHor) Then
                                                    If Not Sprite(aHero).IsDead Then
                                                        sKillHero
                                                        mdl_Sound.PlaySound 6
                                                    End If
                                                End If
                                                ' Kill enemy
                                                If Not Sprite(aHero).IsDead Then
                                                    If IntersectRect(rRes, rBottom, reHit) Then
                                                        Sprite(i).IsDead = True
                                                        Select Case Sprite(i).Enemy
                                                        Case 2
                                                            mdl_Sound.PlaySound 1
                                                        Case 3
                                                            mdl_Sound.PlaySound 4
                                                        Case 4
                                                            mdl_Sound.PlaySound 5
                                                        Case 1 To 99
                                                            mdl_Sound.PlaySound 2
                                                        End Select
                                                        Sprite(i).MovingFrameDelay = 0
                                                        Sprite(i).Shape.Left = (Sprite(i).eFrames - 1) * (Sprite(i).Width / Sprite(i).eFrames)
                                                        Sprite(i).Shape.Right = Sprite(i).Shape.Left + Sprite(i).Width / Sprite(i).eFrames
                                                        Sprite(aHero).IsJumping = False: MJUMP = 0
                                                        Sprite(aHero).IsFalling = False
                                                        MJACC = 2: MJACE = 0
                                                        Sprite(aHero).JumpAgain = True
                                                    End If
                                                End If
                                                
                                                For C = aStart To Sprites
                                                    If Sprite(C).Pos.Left - 80 < Sprite(i).Pos.Left Then
                                                        If Sprite(C).Pos.Right + 80 > Sprite(i).Pos.Right Then
                                                            If Sprite(C).Visible Then
                                                                Sprite(i).vDirect = True
                                                                If Sprite(C).IsGroundSprite Then
                                                                    If Sprite(C).CanCollide Or Sprite(C).Enemy > 0 Then
                                                                        If C <> i Then
                                                                            If IntersectRect(rRes, reHor, Sprite(C).Pos) Then
                                                                                If Sprite(i).vDirect Then Sprite(i).vDirection = -Sprite(i).vDirection
                                                                                Sprite(i).vDirect = False
                                                                            End If
                                                                            If IntersectRect(rRes, reVer, Sprite(C).Pos) Then
                                                                                sPosition i, Sprite(i).Pos.Left, Sprite(C).Pos.Top - Sprite(i).Shape.Bottom
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Next C
                                                sPosition i, Sprite(i).Pos.Left + Sprite(i).vDirection, Sprite(i).Pos.Top
                                                Sprite(i).MovingFrameDelay = Sprite(i).MovingFrameDelay + 1
                                                ' Animate the sprite
                                                If Sprite(i).MovingFrameDelay = Sprite(i).eAnimationDelay Then
                                                    If Sprite(i).eCanRotate Then
                                                        If Sprite(i).vDirection > 0 Then
                                                            If Sprite(i).eCurFrame = 2 Then
                                                                Sprite(i).eCurFrame = 3
                                                            Else
                                                                Sprite(i).eCurFrame = 2
                                                            End If
                                                        Else
                                                            If Sprite(i).eCurFrame = 0 Then
                                                                Sprite(i).eCurFrame = 1
                                                            Else
                                                                Sprite(i).eCurFrame = 0
                                                            End If
                                                        End If
                                                    Else
                                                        Sprite(i).eCurFrame = Sprite(i).eCurFrame + 1: If Sprite(i).eCurFrame = Sprite(i).eFrames - 1 Then Sprite(i).eCurFrame = 0
                                                    End If
                                                    Sprite(i).Shape.Left = Sprite(i).eCurFrame * (Sprite(i).Width / Sprite(i).eFrames)
                                                    Sprite(i).Shape.Right = Sprite(i).Shape.Left + Sprite(i).Width / Sprite(i).eFrames
                                                    Sprite(i).MovingFrameDelay = 0
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
        
                '// Just move the hero
                If MVDIR = 0 Then
                    Sprite(aHero).Shape.Left = 0
                    Sprite(aHero).Shape.Right = 23
                Else
                    Sprite(aHero).Shape.Left = 115
                    Sprite(aHero).Shape.Right = 138
                End If
                
                '// Jump
                If diKeyState.Key(DIK_SPACE) Then
                    If MJACE = 0 Then If MJACC < gJumpMaxFactor Then MJACC = MJACC + 1
                    If MSPACE = 0 Then
                        If Not Sprite(aHero).IsJumping Then
                            If Not Sprite(aHero).IsFalling Then
                                Sprite(aHero).IsJumping = True
                            End If
                        End If
                    End If
                    MSPACE = 1
                Else
                    MSPACE = 0
                End If
                
                If Sprite(aHero).IsJumping Then
                    If MJUMP < MJACC Then
                        MJUMP = MJUMP + 1
                        sPosition aHero, Sprite(aHero).Pos.Left, Sprite(aHero).Pos.Top - gJumpSpeed: MACC = 6
                    Else
                        If MACC > 0 Then
                            MACC = MACC - 0.45
                            sPosition aHero, Sprite(aHero).Pos.Left, Sprite(aHero).Pos.Top - MACC
                        End If
                        MJACE = 1
                    End If
                    If MVDIR = 0 Then
                        Sprite(aHero).Shape.Left = 46: Sprite(aHero).Shape.Right = 69
                    Else
                        Sprite(aHero).Shape.Left = 69: Sprite(aHero).Shape.Right = 92
                    End If
                End If
                ' Move keys
                If Not Sprite(aHero).IsDead Then
                    If diKeyState.Key(DIK_RIGHT) Then
                        MVDIR = 0
                        If Not Sprite(aHero).IsJumping Then
                            Sprite(aHero).Shape.Left = 23
                            Sprite(aHero).Shape.Right = 46
                            MV = MV + 1
                            If MV >= 6 Then
                                Sprite(aHero).Shape.Left = 0
                                Sprite(aHero).Shape.Right = 23
                                If MV = 12 Then MV = 0
                            End If
                        End If
                        If Sprite(aHero).Pos.Left > 200 Then
                            For i = aStart To Sprites
                                If Sprite(i).IsGroundSprite Then
                                    sPosition i, Sprite(i).Pos.Left - gSpeed, Sprite(i).Pos.Top
                                End If
                            Next i
                        Else
                            sPosition aHero, Sprite(aHero).Pos.Left + gSpeed, Sprite(aHero).Pos.Top
                        End If
                    Else
                        If diKeyState.Key(DIK_LEFT) Then
                            MVDIR = 1
                            If Not Sprite(aHero).IsJumping Then
                                Sprite(aHero).Shape.Left = 115
                                Sprite(aHero).Shape.Right = 138
                                MV = MV + 1
                                If MV >= 6 Then
                                    Sprite(aHero).Shape.Left = 92
                                    Sprite(aHero).Shape.Right = 115
                                    If MV = 12 Then MV = 0
                                End If
                            End If
                            If Sprite(aHero).Pos.Left < 120 Then
                                For i = aStart To Sprites
                                    If Sprite(i).IsGroundSprite Then
                                        sPosition i, Sprite(i).Pos.Left + gSpeed, Sprite(i).Pos.Top
                                    End If
                                Next i
                            Else
                                sPosition aHero, Sprite(aHero).Pos.Left - gSpeed, Sprite(aHero).Pos.Top
                            End If
                            'mdl_Sound.PlaySound 5
                        End If
                    End If
                End If
                If Sprite(aHero).IsFalling Then
                    If MVDIR = 0 Then
                        Sprite(aHero).Shape.Left = 46: Sprite(aHero).Shape.Right = 69
                    Else
                        Sprite(aHero).Shape.Left = 69: Sprite(aHero).Shape.Right = 92
                    End If
                End If
                
                If Sprite(aHero).JumpAgain Then Sprite(aHero).IsJumping = True: Sprite(aHero).JumpAgain = False
            End If
            ' Show all
            If gJustStarted Then
                ' This was too slow to use, the player would drive crazy
                ' but if you want to implement some "intro" show then you do it here
    '            gSleep 1
    '            SG = SG + 1
    '            gSetGamma CInt(SG), CInt(SG), CInt(SG)
    '            If SG = 0 Then gJustStarted = False
                gJustStarted = False
            End If
            
            sRender
    
            '// End
            gFPSC = gFPSC + 1 ' Count the fps
            T2 = GetTickCount ' Get the tickcount
            T3 = T3 + (T2 - T1) ' Calc the interval in ms, for one loop
        End If
        DoEvents
    Loop
    diKeyboard.Unacquire
    dTerminate
End Sub

Public Sub sAdd(sFile As String, Optional W&, Optional H&)
    Dim NeedSurf As Boolean
    
    NeedSurf = True
    For i = 0 To Surfaces
        If Surf(i).FileName = sFile Then
            NeedSurf = False
            Exit For
        End If
    Next i
    
    If NeedSurf Then
        Surfaces = Surfaces + 1: i = Surfaces
    
        ' Get the size
        If W = 0 And H = 0 Then
            Dim sPic As IPictureDisp
            Set sPic = LoadPicture(App.Path & "\" & sFile)
            W = sPic.Width / 26.4594594594595
            H = sPic.Height / 26.4594594594595
        
            Surf(Surfaces).Width = W
            Surf(Surfaces).Height = H
        End If
        
        ' Tell the dx what we will use here
        Surf(Surfaces).Description.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT  ' Capabillities, w and h
        Surf(Surfaces).Description.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN 'Or DDSCAPS_VIDEOMEMORY   ' It is in memory, not on the screen
        
        ' Loop all the bitmaps in the root dir
        Surf(Surfaces).Description.lWidth = W
        Surf(Surfaces).Description.lHeight = H ' Set the size
        
        Surf(Surfaces).FileName = sFile
        
        Set Surf(Surfaces).Surf = dD.CreateSurfaceFromFile(App.Path & "\" & Surf(Surfaces).FileName, Surf(Surfaces).Description)     ' Create the surface, DirectDraw will load it from bitmap
    
        Surf(Surfaces).Surf.SetColorKey DDCKEY_SRCBLT, Surf(Surfaces).ColorKey
    End If
    
    Sprites = Sprites + 1 ' Increase the counting
    
    Sprite(Sprites).Surface = i
    
    Sprite(Sprites).IsCreated = True
    
    Sprite(Sprites).Visible = True
    Sprite(Sprites).Width = Surf(i).Description.lWidth
    Sprite(Sprites).Height = Surf(i).Description.lHeight
    Sprite(Sprites).Shape.Left = 0
    Sprite(Sprites).Shape.Right = 0
    Sprite(Sprites).Shape.Right = Surf(i).Description.lWidth
    Sprite(Sprites).Shape.Bottom = Surf(i).Description.lHeight
    
    ' The rect of the sprite will be set later
End Sub

Public Sub sPosition(sIndex As Long, x&, Y&)
    ' Position the sprite, and set its rect
    Sprite(sIndex).Pos.Left = x
    Sprite(sIndex).Pos.Top = Y
    Sprite(sIndex).Pos.Right = x + Sprite(sIndex).Shape.Right - Sprite(sIndex).Shape.Left
    Sprite(sIndex).Pos.Bottom = Y + Sprite(sIndex).Shape.Bottom - Sprite(sIndex).Shape.Top
End Sub

Public Sub gLoadCustomSprites()
    sAdd "back1.bmp"
    sPosition Sprites, 0, 0
    Sprite(Sprites).IsBackBitmap = True
    Sprite(Sprites).IsConstant = True
    sAdd "hero.bmp"
    sPosition Sprites, 32, 32
    Sprite(Sprites).Shape.Right = 23
    Sprite(Sprites).Shape.Left = 0
    Sprite(Sprites).IsHero = True
    Sprite(Sprites).OffScreen = True
    Sprite(Sprites).IsConstant = True
End Sub

Public Sub sRender()
    ' Here we draw the basic sprites on the screen
    Dim Temp As RECT, x&, Y&
    For i = 1 To Sprites
        If Sprite(i).Visible Then
            ' Blt the sprites (surfaces)
            Temp.Left = Sprite(i).Shape.Left
            Temp.Top = Sprite(i).Shape.Top
            Temp.Right = Sprite(i).Shape.Right
            Temp.Bottom = Sprite(i).Shape.Bottom
            x = Sprite(i).Pos.Left
            Y = Sprite(i).Pos.Top
            
            ' Here we will cut the rect if it has left the screen
            ' if we let X or Y to be pass zero, then fastblt will
            ' give invalid rect error, here we solve that problem.
            If Sprite(i).OffScreen Then
                If Sprite(i).Pos.Right > dRes.x Then
                    Temp.Right = Temp.Left + (dRes.x - Sprite(i).Pos.Left)
                Else
                    Temp.Right = Sprite(i).Shape.Right
                End If
                If Sprite(i).Pos.Left < 0 Then
                    x = 0 ' Block
                    Temp.Left = Abs(Sprite(i).Pos.Left) + Sprite(i).Shape.Left ' set the rest part
                Else
                    Temp.Left = Sprite(i).Shape.Left
                End If
                
                If Sprite(i).Pos.Bottom > dRes.Y Then
                    Temp.Bottom = Temp.Top + (dRes.Y - Sprite(i).Pos.Top)
                Else
                    Temp.Bottom = Sprite(i).Shape.Bottom
                End If
                If Sprite(i).Pos.Top < 0 Then
                    Y = 0 ' Block
                    Temp.Top = Abs(Sprite(i).Pos.Top) + Sprite(i).Shape.Top ' set the rest part
                Else
                    Temp.Top = Sprite(i).Shape.Top
                End If
            End If
            ' The hart of the whole code, one line of code will draw all this on the screen
            If Sprite(i).Pos.Left < dRes.x And Sprite(i).Pos.Right > 0 Then
                dsBack.BltFast x, Y, Surf(Sprite(i).Surface).Surf, Temp, DDBLTFAST_SRCCOLORKEY
            End If
        End If
    Next i
    
    dsBack.SetForeColor vbWhite
    dsBack.SetFont dFont
    If Not gMenu Then
        dsBack.DrawText 1, 5, "Score: " & gScore, False
        dsBack.DrawText 70, 5, "Green Diamonds: " & gGreens, False
        dsBack.DrawText 200, 5, "Lives: " & gLives, False
    End If
    'dsBack.DrawText 1, 128, gFps & " FPS", False
    If gWarning > 0 Then
        dsBack.SetForeColor vbRed
        dsBack.SetFont wFont
        dsBack.DrawText 15, 80, "You must find " & gGreenCount - gGreens & " green diamonds !", False
        gWarning = gWarning + 1: If gWarning = 60 Then gWarning = 0
        dsBack.SetFont dFont
    End If
    
    dsFront.Flip Nothing, DDFLIP_WAIT ' We need to flip it
End Sub
    
Public Sub gLoadLevel(File As String, JustImage As Boolean)
    Dim T$, x&, Y&, R&, S As RECT
    sClear
    If Not JustImage Then gLoadCustomSprites
    gGreens = 0
    gGreenCount = 0
    dsBack.BltColorFill S, 0
    gSetGamma 0, 0, 0
    dsBack.SetForeColor vbWhite
    dFont.Size = 18
    dsBack.SetFont dFont
    dsBack.DrawText 90, 80, "LOADING...", False
    dsFront.Flip Nothing, DDFLIP_WAIT
    DoEvents
    Open File For Input As #1
        Do Until EOF(1)
            Input #1, x, Y, T
            sShowSprite x * 16, Y * 16, T
        Loop
    Close #1
    dFont.Size = 8
End Sub


Public Sub sShowSprite(x&, Y&, T$)
    Dim R%

    Select Case T
        '// TRANSPARENT OBJECTS
        Case 5: sAdd "palm.bmp"
        Case 6: sAdd "palm_s.bmp"
        Case 7: sAdd "ltree_s.bmp"
        Case 8: sAdd "ltree.bmp"
        Case 9: sAdd "cloud.bmp"
        Case 23: sAdd "level_s.bmp": Sprite(Sprites).IsLevelDone = True
        Case 24: sAdd "level.bmp"
        Case 16
            sAdd "items.bmp"
            Sprite(Sprites).Shape.Left = 80: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Sprite(Sprites).Item = 1
        Case 17
            sAdd "items.bmp"
            Sprite(Sprites).Shape.Left = 0: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Sprite(Sprites).Item = 2
        Case 18
            sAdd "items.bmp"
            Sprite(Sprites).Shape.Left = 32: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Sprite(Sprites).Item = 3
        Case 19
            sAdd "items.bmp"
            Sprite(Sprites).Shape.Left = 64: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Sprite(Sprites).Item = 4
        Case 20
            sAdd "items.bmp"
            Sprite(Sprites).Shape.Left = 48: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Sprite(Sprites).Item = 5
        Case 21
            sAdd "items.bmp"
            Sprite(Sprites).Shape.Left = 16: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Sprite(Sprites).Item = 6
        
        '// ENEMIES
        Case 11
            If gLevel = 1 Then
                sAdd "alien.bmp"
                Sprite(Sprites).Enemy = 4
                Sprite(Sprites).eAnimationDelay = 3
                Sprite(Sprites).eSpeed = 2
                Sprite(Sprites).Shape.Left = 0: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 20
                Sprite(Sprites).eFrames = 5
                Sprite(Sprites).eTolerance = 3
                Sprite(Sprites).eCanRotate = True
            Else
                sAdd "enemy.bmp"
                Sprite(Sprites).Enemy = 1
                Sprite(Sprites).eAnimationDelay = 20
                Sprite(Sprites).eSpeed = 1
                Sprite(Sprites).Shape.Left = 0: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 20
                Sprite(Sprites).eFrames = 3
                Sprite(Sprites).eTolerance = 8
            End If
        Case 12
            sAdd "punksnotdead.bmp"
            Sprite(Sprites).Enemy = 2
            Sprite(Sprites).eAnimationDelay = 5
            Sprite(Sprites).eSpeed = 1
            Sprite(Sprites).Shape.Left = 0: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 20
            Sprite(Sprites).eFrames = 5
            Sprite(Sprites).eTolerance = 3
            Sprite(Sprites).eCanRotate = True
        Case 13
            sAdd "redhat.bmp"
            Sprite(Sprites).Enemy = 3
            Sprite(Sprites).eAnimationDelay = 5
            Sprite(Sprites).eSpeed = 1
            Sprite(Sprites).Shape.Left = 0: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 20
            Sprite(Sprites).eFrames = 5
            Sprite(Sprites).eTolerance = 3
            Sprite(Sprites).eCanRotate = True
    Case Else
        '// HARD OBJECTS
        Select Case gLevel
        Case 0
            sAdd "earth.bmp"
        Case 1
            sAdd "concrete.bmp"
        Case 2
            sAdd "earth.bmp"
        End Select
        Sprite(Sprites).CanCollide = True
        Select Case T
            Case "0": Sprite(Sprites).Shape.Left = 64: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Case "1": Sprite(Sprites).Shape.Left = 0: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Case "2l": Sprite(Sprites).Shape.Left = 48: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Case "2r": Sprite(Sprites).Shape.Left = 32: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Case "3l": Sprite(Sprites).Shape.Left = 80: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Case "3r": Sprite(Sprites).Shape.Left = 16: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Case "10": Sprite(Sprites).Shape.Left = 208: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Case "14": Sprite(Sprites).Shape.Left = 224: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Case "15": Sprite(Sprites).Shape.Left = 240: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
            Case "22": Sprite(Sprites).Shape.Left = 256: Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16: Sprite(Sprites).IsSpiky = True
            ' These are some random objects
            Case "4"
                Randomize
                R = CLng((Rnd * 6) + 0)
                Sprite(Sprites).Shape.Left = 96 + (16 * R): Sprite(Sprites).Shape.Right = Sprite(Sprites).Shape.Left + 16
                Sprite(Sprites).CanCollide = False
        End Select
    End Select
    Sprite(Sprites).IsGroundSprite = True
    Sprite(Sprites).OffScreen = True
    Sprite(Sprites).vDirection = -Sprite(Sprites).eSpeed
    If Sprite(Sprites).Item = 1 Then gGreenCount = gGreenCount + 1
    If Sprite(Sprites).Enemy > 0 Then
        sPosition Sprites, x, Y - 4
    Else
        sPosition Sprites, x, Y
    End If
End Sub

Public Sub sKillHero()
    ' To animate hero when dying
    gLives = gLives - 1
    Sprite(aHero).Visible = False
    Sprite(aHero).IsDead = True
    sAdd "deadhero.bmp"
    Sprite(Sprites).Shape.Left = 0
    Sprite(Sprites).Shape.Right = 30
    Sprite(Sprites).IsShowed = True: Sprite(Sprites).OffScreen = True: Sprite(Sprites).IsGroundSprite = True
    sPosition Sprites, Sprite(aHero).Pos.Left, Sprite(aHero).Pos.Top
End Sub


Public Function ConvToSignedValue(lngValue As Long) As Integer
    ' Cheezy method for converting to signed integer
    If lngValue <= 32767 Then
        ConvToSignedValue = CInt(lngValue)
        Exit Function
    End If
    ConvToSignedValue = CInt(lngValue - 65535)
End Function

Public Function ConvToUnSignedValue(intValue As Integer) As Long
    ' Cheezy method for converting to unsigned integer
    If intValue >= 0 Then
        ConvToUnSignedValue = intValue
        Exit Function
    End If
    ConvToUnSignedValue = intValue + 65535
End Function

Public Sub gSetGamma(intRed As Integer, intGreen As Integer, intBlue As Integer)
    'Alter the gamma ramp to the percent given by comparing to original state
    'A value of zero ("0") for intRed, intGreen, or intBlue will result in the
    'gamma level being set back to the original levels. Anything ABOVE zero will
    'fade towards FULL colour, anything below zero will fade towards NO colour
    For i = 0 To 255
        If intRed < 0 Then drRamp.red(i) = ConvToSignedValue(ConvToUnSignedValue(drOriginal.red(i)) * (100 - Abs(intRed)) / 100)
        If intRed = 0 Then drRamp.red(i) = drOriginal.red(i)
        If intRed > 0 Then drRamp.red(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(drOriginal.red(i))) * (100 - intRed) / 100))
        If intGreen < 0 Then drRamp.green(i) = ConvToSignedValue(ConvToUnSignedValue(drOriginal.green(i)) * (100 - Abs(intGreen)) / 100)
        If intGreen = 0 Then drRamp.green(i) = drOriginal.green(i)
        If intGreen > 0 Then drRamp.green(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(drOriginal.green(i))) * (100 - intGreen) / 100))
        If intBlue < 0 Then drRamp.blue(i) = ConvToSignedValue(ConvToUnSignedValue(drOriginal.blue(i)) * (100 - Abs(intBlue)) / 100)
        If intBlue = 0 Then drRamp.blue(i) = drOriginal.blue(i)
        If intBlue > 0 Then drRamp.blue(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(drOriginal.blue(i))) * (100 - intBlue) / 100))
    Next
    dGC.SetGammaRamp DDSGR_DEFAULT, drRamp
End Sub

Public Sub gSleep(MS As Long)
    gSleepTick = GetTickCount()
    Do While gSleepTick + MS > GetTickCount()
    Loop
End Sub

Public Sub gLoadSounds()
    mdl_Sound.LoadSound "fall.wav", True, False, False, False, False, False, False, Me
        mdl_Sound.SetVolume 0, -3000
    mdl_Sound.LoadSound "punk.wav", True, False, False, False, False, False, False, Me
    mdl_Sound.LoadSound "hit.wav", True, False, False, False, False, False, False, Me
    mdl_Sound.LoadSound "item.wav", True, False, False, False, False, False, False, Me
    mdl_Sound.LoadSound "redhat.wav", True, False, False, False, False, False, False, Me
    mdl_Sound.LoadSound "alien.wav", True, False, False, False, False, False, False, Me
    mdl_Sound.LoadSound "dead.wav", True, False, False, False, False, True, False, Me
        mdl_Sound.SetVolume 6, -2000
End Sub


Public Sub sClear()
    ' Well... here I am clearing all the sprites
    For i = 0 To Sprites
        Sprite(i).CanCollide = False
        Sprite(i).eAnimationDelay = 0
        Sprite(i).eCanRotate = False
        Sprite(i).eCurFrame = 0
        Sprite(i).eFrames = 0
        Sprite(i).Enemy = 0
        Sprite(i).eSpeed = 0
        Sprite(i).eTolerance = 0
        Sprite(i).IsBackBitmap = False
        Sprite(i).IsConstant = False
        Sprite(i).IsDead = False
        Sprite(i).IsFalling = False
        Sprite(i).IsGroundSprite = False
        Sprite(i).IsHero = False
        Sprite(i).IsJumping = False
        Sprite(i).IsLevelDone = False
        Sprite(i).IsShowed = False
        Sprite(i).IsSpiky = False
        Sprite(i).Item = 0
        Sprite(i).JumpAgain = False
        Sprite(i).MovingFrameDelay = 0
        Sprite(i).OffScreen = False
        Sprite(i).sType = 0
        Sprite(i).IsCreated = False
        Sprite(i).vDirect = False
        Sprite(i).vDirection = 0
        Sprite(i).Visible = False
    Next i
    
    For i = 0 To Surfaces
        Set Surf(i).Surf = Nothing
        Surf(i).FileName = ""
    Next i
    
    Sprites = 0
End Sub
