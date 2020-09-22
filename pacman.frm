VERSION 5.00
Begin VB.Form frmPacMan 
   Appearance      =   0  'Flat
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desktop PacMan"
   ClientHeight    =   8430
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10065
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   Icon            =   "pacman.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picGame 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      ScaleHeight     =   423
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   559
      TabIndex        =   2
      Top             =   0
      Width           =   8415
   End
   Begin VB.PictureBox Bmps 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   5880
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   152
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3255
      Left            =   1200
      ScaleHeight     =   215
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
         Begin VB.Menu mnuEasy 
            Caption         =   "Easy"
         End
         Begin VB.Menu mnuNormal 
            Caption         =   "Normal"
         End
         Begin VB.Menu mnuHard 
            Caption         =   "Hard"
         End
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
         Checked         =   -1  'True
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu Settings 
      Caption         =   "Settings"
      Begin VB.Menu ScreenSz 
         Caption         =   "Screen Size"
         Begin VB.Menu mnuTimesHalf 
            Caption         =   "1:2"
         End
         Begin VB.Menu mnuTimesOne 
            Caption         =   "1:1"
         End
         Begin VB.Menu mnuTimesTwo 
            Caption         =   "2:1"
         End
         Begin VB.Menu mnuThreeTimesOne 
            Caption         =   "3:1"
         End
      End
      Begin VB.Menu SNDs 
         Caption         =   "Sounds"
         Begin VB.Menu mnuSoundOn 
            Caption         =   "On"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSoundOff 
            Caption         =   "Off"
         End
      End
   End
End
Attribute VB_Name = "frmPacMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private g_InGame As Boolean
Private Ghosts() As New CGhosts
Private pPacPosition As New CPacMan
Private pPacTiles(NumberOfTilesPerRow * NumberOfRows)  As New CTilesAndFoods
Private FoodMap(SizeX, SizeY) As Boolean
Private bPaused As Boolean
Private bSND As Boolean
Private lSoundState As Long
Private NUM_SPECIAL_TILES As Integer
Private Type SpecialTile
    IDX As Integer
    IDY As Integer
End Type
Private tSuperPacmanTimer As Integer
Private pSpecialTile() As SpecialTile
Private animCounter As Integer

'######################################################################
'#                                                                    #
'# Visual Basic Events.                                               #
'#                                                                    #
'######################################################################

Private Sub Form_Activate()
    'Here's where it all starts
    frmPacMan.Bmps.Picture = LoadPicture(App.Path + "\pm.bmp")
    g_InGame = False
    g_screenx = 210
    g_screeny = 190
    
    DoResize
    bPaused = False
    mnuPause.Checked = False
    mnuTimesOne.Checked = True
    'Sound On
    bSND = True
    g_bGameRunning = True
    MainGameLoop
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call sndPlaySound("", SND_LOOP Or SND_ASYNC)
    g_bGameRunning = False
    End
End Sub

Private Sub mnuEasy_Click()
    StartGame
End Sub

Private Sub mnuPause_Click()
    If bPaused = False Then
        bPaused = True
        mnuPause.Checked = True
        Timer1.Interval = 0
        tmrAnimationPaccy.Interval = 0
    Else
        bPaused = False
        mnuPause.Checked = False
        Timer1.Interval = 1
        tmrAnimationPaccy.Interval = 200
    End If
End Sub

Private Sub mnuQuit_Click()
    End
End Sub

Private Sub mnuSoundOff_Click()
    mnuSoundOn.Checked = False
    mnuSoundOff.Checked = True
    bSND = False
End Sub

Private Sub mnuSoundOn_Click()
    mnuSoundOn.Checked = True
    mnuSoundOff.Checked = False
    bSND = True
End Sub

Private Sub mnuThreeTimesOne_Click()
    g_screenx = 210 * 3
    g_screeny = 190 * 3
    mnuTimesHalf.Checked = False
    mnuTimesOne.Checked = False
    mnuTimesTwo.Checked = False
    mnuThreeTimesOne.Checked = True
    DoResize
End Sub

Private Sub mnuTimesHalf_Click()
    g_screenx = 210 / 2
    g_screeny = 190 / 2
    mnuTimesHalf.Checked = True
    mnuTimesOne.Checked = False
    mnuTimesTwo.Checked = False
    mnuThreeTimesOne.Checked = False
    DoResize
End Sub

Private Sub mnuTimesOne_Click()
    g_screenx = 210
    g_screeny = 190
    mnuTimesHalf.Checked = False
    mnuTimesOne.Checked = True
    mnuTimesTwo.Checked = False
    mnuThreeTimesOne.Checked = False
    DoResize
End Sub

Private Sub mnuTimesTwo_Click()
    g_screenx = 210 * 2
    g_screeny = 190 * 2
    mnuTimesHalf.Checked = False
    mnuTimesOne.Checked = False
    mnuTimesTwo.Checked = True
    mnuThreeTimesOne.Checked = False
    DoResize
End Sub

Private Sub picGame_KeyDown(KeyCode As Integer, Shift As Integer)
  If g_InGame = True Then
        Select Case KeyCode
        Case vbKeyLeft:
            pPacPosition.pDirection = 0
        Case vbKeyRight:
            pPacPosition.pDirection = 1
        Case vbKeyUp:
            pPacPosition.pDirection = 2
        Case vbKeyDown:
            pPacPosition.pDirection = 3
        End Select
    Else
        If KeyCode = vbKeyReturn Then StartGame
    End If
End Sub

'######################################################################
'#                                                                    #
'# The main game loop...                                              #
'# This calls and updates all the things that make this a game!!!!!!  #
'#                                                                    #
'######################################################################

Private Sub MainGameLoop()

    Dim frequency As Currency
    Dim currentTime As Currency
    Dim startTime As Currency
    Dim time_elapsed As Double
    Dim result As Double
    Dim next_time As Double
    Dim last_time As Double
    Dim perf_flag As Boolean
    Dim time_scale As Double
    Dim time_count As Double
    
    time_scale = 0.001
    perf_flag = True
   
    ' get the frequency counter
    ' return zero if hardware doesn't support high-res performance counters
    ' In this case we use the timeGettimer
    
    If QueryPerformanceFrequencyAny(frequency) = 0 Then
         perf_flag = False
    Else
        time_count = Val(frequency) / FPS 'fps
        time_scale = (1 / Val(frequency))
    End If
    
        
    '--------main game loop-----
    Do While g_bGameRunning = True
        

        If (perf_flag = True) Then
            QueryPerformanceCounterAny currentTime
        Else
            currentTime = timeGetTime()
        End If
        
            If (currentTime > next_time) Then
            'Set our counter
                time_elapsed = (currentTime - last_time) * time_scale
                last_time = currentTime
                DoEvents
                'Must do events so the keyboard pressess will be detected
                If g_InGame = True Then
                    'Clear the backbuffer
                    picBG.Cls
                    'Draw the tiles to the backbuffer
                    prepare_tiles
                    'Draw the food to the backbuffer
                    UpdateFood
                    'Check to see if we have completed the level
                    If TotalNumOfFoods <= 0 Then
                        MsgBox "Level Complete"
                        LevelUP
                    End If
                    'Draw the PacMan to the backbuffer
                    UpdateAnimations
                    pPacPosition.Draw_PacMan
        
                    Dim GhostCount As Integer
                    'Draw the ghosts to the backbuffer
                    For GhostCount = 0 To NUM_GHOSTS
                        Ghosts(GhostCount).Draw_Ghost
                    Next GhostCount
                
                    'Flip the backbuffer to the viewable screen
                    Call StretchBlt(picGame.hdc, 0, 0, g_screenx, g_screeny, picBG.hdc, 0, 0, 210, 190, SRCCOPY)
               
                Else
                    StartScreen
                End If
                    
                    'Increase the timer 'All done this time
                    next_time = currentTime + time_count
            End If
      
    Loop
    
End Sub



'######################################################################
'#                                                                    #
'# StartGame...........                                               #
'# Initailize everything for the game                                 #
'#                                                                    #
'######################################################################
Private Sub StartGame()
    load_map
    prepare_tiles
    SetFood
    pPacPosition.pLives = 3
    tSuperPacmanTimer = 0
    Dim gdc As Integer
    
    For gdc = 0 To NUM_GHOSTS
        Ghosts(gdc).ghScreenX = Ghosts(gdc).ghSrcX * tTileSize
        Ghosts(gdc).ghScreenY = Ghosts(gdc).ghSrcY * tTileSize
        Ghosts(gdc).ghDirection = 1
        Ghosts(gdc).ghFrame = 1
        Ghosts(gdc).ghAlive = True
    Next gdc
    
    pPacPosition.pScreenX = pPacPosition.pSrcX * tTileSize
    pPacPosition.pScreenY = pPacPosition.pSrcY * tTileSize
    pPacPosition.pDirection = 0
    pPacPosition.pbSuperPacMan = False
    
    g_InGame = True
    
    'Main Game loop
    MainGameLoop
    
    
End Sub

'######################################################################
'#                                                                    #
'# StartScreen.........                                               #
'# We are not paying the game so display the start screen             #
'#                                                                    #
'######################################################################
Private Sub StartScreen()
    Call StretchBlt(picGame.hdc, 0, 0, g_screenx, g_screeny, Bmps.hdc, 0, 10, 110, 90, SRCCOPY)
    'Play sound
    If bSND = True Then
        If lSoundState = 0 Then
             Call sndPlaySound(App.Path & "\sounds\GAMEBEGINNING.wav", 1)
             lSoundState = 1
        End If
    End If
End Sub

'######################################################################
'#                                                                    #
'# load_map............                                               #
'# Loads the ASCII map from a file into the array                     #
'#                                                                    #
'######################################################################
Private Sub load_map()
    
    
    If g_LevelID = 0 Then g_LevelID = 1
    gMapPath = App.Path & "\leveldata\"
    gMapName = "level" & CStr(g_LevelID) & ".pac"
    Open gMapPath & gMapName For Input As #1
        
    Dim temp1 As String
    Dim results As String
    Dim spos As Integer
    Dim mk As String
    Dim lnt As Integer
    Dim XMap As Integer
    Dim YMap As Integer
    Dim gc As Integer
    
    XMap = 0
    YMap = 0

    'Check to see if the file is correct
    'First Line in file
    Input #1, temp1
    If temp1 <> "[Desktop PacMan Initailization File V1.0]" Then
        MsgBox "There was a problem with the initialization file!", 0, "Error."
        'Close the file
        Close #1
        'End the program
        End
    End If
    'Second Line in file
    Input #1, temp1
    If temp1 <> "[LevelID]" Then
        MsgBox "There was a problem with the initialization file!", 0, "Error."
        'Close the file
        Close #1
        'End the program
        End
    Else 'third Line in file
        Input #1, temp1
        g_LevelID = CInt(temp1)
    End If
    'Fourth Line in file
    Input #1, temp1
    If temp1 <> "[NumberOfGhosts]" Then
        MsgBox "There was a problem with the initialization file!", 0, "Error."
        'Close the file
        Close #1
        'End the program
        End
    Else 'Fifth Line in file
        Input #1, temp1
        NUM_GHOSTS = CInt(temp1) - 1
        ReDim Ghosts(NUM_GHOSTS) 'Create Ghost array based on size of NUM_GHOSTS
    End If
    'Sixth Line in file
    Input #1, temp1
    If temp1 <> "[GhostColors]" Then
        MsgBox "There was a problem with the initialization file!", 0, "Error."
        'Close the file
        Close #1
        'End the program
        End
    Else 'Seventh Line in file
        
        For gc = 0 To NUM_GHOSTS
            Input #1, temp1
            Ghosts(gc).ghColorPointer = CInt(temp1)
        Next gc
    End If
    'Based on number of ghosts this could be any line in the file!!!!!
    Input #1, temp1
    If temp1 <> "[GhostXYPositions]" Then
        MsgBox "There was a problem with the initialization file!", 0, "Error."
        'Close the file
        Close #1
        'End the program
        End
    Else
        For gc = 0 To NUM_GHOSTS
            Input #1, temp1
            Ghosts(gc).ghSrcX = CInt(Left$(temp1, 2))
            Ghosts(gc).ghSrcY = CInt(Right$(temp1, 2))
        Next gc
    End If
    'Based on number of ghosts this could be any line in the file!!!!!
    Input #1, temp1
    If temp1 <> "[PacManXYPosition]" Then
        MsgBox "There was a problem with the initialization file!", 0, "Error."
        'Close the file
        Close #1
        'End the program
        End
    Else
        Input #1, temp1
        pPacPosition.pSrcX = CInt(Left$(temp1, 2))
        pPacPosition.pSrcY = CInt(Right$(temp1, 2))
    End If
    'Based on number of ghosts this could be any line in the file!!!!!
    Input #1, temp1
    If temp1 <> "[NumberOfSpecialFood]" Then
        MsgBox "There was a problem with the initialization file!", 0, "Error."
        'Close the file
        Close #1
        'End the program
        End
    Else
        Input #1, temp1
        NUM_SPECIAL_TILES = CInt(temp1)
        ReDim pSpecialTile(NUM_SPECIAL_TILES)
    End If
    'Based on number of ghosts this could be any line in the file!!!!!
    Input #1, temp1
    If temp1 <> "[SpecialFoodXYPositions]" Then
        MsgBox "There was a problem with the initialization file!", 0, "Error."
        'Close the file
        Close #1
        'End the program
        End
    Else
        For gc = 1 To NUM_SPECIAL_TILES
            Input #1, temp1
            pSpecialTile(gc).IDX = CInt(Left$(temp1, 2))
            pSpecialTile(gc).IDY = CInt(Right$(temp1, 2))
        Next gc
    End If
    'Based on number of ghosts this could be any line in the file!!!!!
    Input #1, temp1
    If temp1 <> "[Map]" Then
        MsgBox "There was a problem with the initialization file!", 0, "Error."
        'Close the file
        Close #1
        'End the program
        End
    Else
        Do Until EOF(1)
            'Get each line of the file and parse into the PacMap array.
            Input #1, temp1
            spos = 1
            lnt = 1
            XMap = 0
           
            For C = 1 To Len(temp1)
                lnt = (C - spos)
                results = Mid$(temp1, spos, lnt + 1)
                PacMap(XMap, YMap) = Asc(results) - 48
                ' Take off 48 cus we start at AScii 49 which is the
                ' number ONE
                
                spos = C + 1
                XMap = XMap + 1
            Next C
            YMap = YMap + 1
        Loop
    End If
    
    'Close the file
    Close #1

    Dim xx As Integer
    Dim yy As Integer
    
    For yy = 0 To 0 ' Not using this yy yet :) for future use :)
        For xx = 0 To NumberOfTilesPerRow
            pPacTiles(yy + xx).tSrcX = xx * 10
            pPacTiles(yy + xx).tSrcY = yy * 10 ' Not using this yy yet --- ZERO
        Next xx
    Next yy
    
End Sub

'######################################################################
'#                                                                    #
'# prepare_tiles.......                                               #
'# Converts the map array into screen coordinates and draws the tiles #
'# to the screen....                                                  #
'#                                                                    #
'######################################################################
Private Sub prepare_tiles()
    
    Dim XGraph As Integer
    Dim YGraph As Integer
    XGraph = 0
    YGraph = 0
    
    For y = 0 To SizeY
        For x = 0 To SizeX
            Call BitBlt(picBG.hdc, OffX + XGraph, OffY + YGraph, tTileSize, tTileSize, Bmps.hdc, pPacTiles(PacMap(x, y)).tSrcX, pPacTiles(PacMap(x, y)).tSrcY, vbSrcPaint)
            XGraph = XGraph + tTileSize
        Next x
        YGraph = YGraph + tTileSize
        XGraph = 0
    Next y

End Sub

'######################################################################
'#                                                                    #
'#  Private Sub UpdateAnimations()...                                               #
'# This calls and updates the movement and animation of the ghosts    #
'# and Pacman....                                                     #
'#                                                                    #
'######################################################################

 Private Sub UpdateAnimations()
    'If we are playing
    If g_InGame = True Then
    
        
    
    'PacMan Movements and animations
    Select Case pPacPosition.pDirection
        'Direction = Left
        Case 0:
        If PacMap(pPacPosition.pSrcX - 1, pPacPosition.pSrcY) = 0 Then
            pPacPosition.pSpeedX = -tTileSize
            pPacPosition.pSpeedY = 0
            pPacPosition.pSrcX = pPacPosition.pSrcX - 1
            pPacPosition.pDirection = 0
        Else
            pPacPosition.pSpeedX = 0
            pPacPosition.pSpeedY = 0
        End If
        
        
        
        pPacPosition.pBltX = 230
        If pPacPosition.pFrame = 1 Then
            pPacPosition.pBltX = pPacPosition.pBltX + tTileSize
            pPacPosition.pFrame = 0
        Else
            pPacPosition.pFrame = 1
        End If
        
        
        'Direction = Right
        Case 1:
        If PacMap(pPacPosition.pSrcX + 1, pPacPosition.pSrcY) = 0 Then
            pPacPosition.pSpeedX = tTileSize
            pPacPosition.pSpeedY = 0
            pPacPosition.pSrcX = pPacPosition.pSrcX + 1
            pPacPosition.pDirection = 1
        Else
            pPacPosition.pSpeedX = 0
            pPacPosition.pSpeedY = 0
        End If
        
        pPacPosition.pBltX = 250
        If pPacPosition.pFrame = 1 Then
            pPacPosition.pBltX = pPacPosition.pBltX + tTileSize
            pPacPosition.pFrame = 0
        Else
            pPacPosition.pFrame = 1
        End If
        
        'Direction = Up
        Case 2:
        If PacMap(pPacPosition.pSrcX, pPacPosition.pSrcY - 1) = 0 Then
            pPacPosition.pSpeedX = 0
            pPacPosition.pSpeedY = -tTileSize
            pPacPosition.pSrcY = pPacPosition.pSrcY - 1
            pPacPosition.pDirection = 2
        Else
            pPacPosition.pSpeedX = 0
            pPacPosition.pSpeedY = 0
        End If
         
         
        pPacPosition.pBltX = 270
        If pPacPosition.pFrame = 1 Then
            pPacPosition.pBltX = pPacPosition.pBltX + tTileSize
            pPacPosition.pFrame = 0
        Else
            pPacPosition.pFrame = 1
        End If
        
        'Direction = Down
        Case 3:
        If PacMap(pPacPosition.pSrcX, pPacPosition.pSrcY + 1) = 0 Then
            
            pPacPosition.pSpeedX = 0
            pPacPosition.pSpeedY = tTileSize
            pPacPosition.pSrcY = pPacPosition.pSrcY + 1
            pPacPosition.pDirection = 3

        Else
            pPacPosition.pSpeedX = 0
            pPacPosition.pSpeedY = 0
        End If
        
        
            pPacPosition.pBltX = 290
            If pPacPosition.pFrame = 1 Then
                pPacPosition.pBltX = pPacPosition.pBltX + tTileSize
                pPacPosition.pFrame = 0
            Else
                pPacPosition.pFrame = 1
            End If
        
    End Select
        
    'Play a sound we Pacman eats the food, increase the score and set the position
    'In the food map to false, Indicating that the food was eaten!
    If FoodMap(pPacPosition.pSrcX, pPacPosition.pSrcY) = True Then
        TotalNumOfFoods = TotalNumOfFoods - 1
        If bSND = True Then sndPlaySound App.Path & "\sounds\pacchomp.wav", 1
        FoodMap(pPacPosition.pSrcX, pPacPosition.pSrcY) = False
        pPacPosition.pScore = pPacPosition.pScore + 1
    End If
    
    'Update PacMan Postition
    pPacPosition.pScreenX = pPacPosition.pScreenX + pPacPosition.pSpeedX
    pPacPosition.pScreenY = pPacPosition.pScreenY + pPacPosition.pSpeedY
    For gg = 0 To NUM_GHOSTS
        Ghosts(gg).Move_Ghost
    Next gg
    
    'Colision detection
    CheckHit
    
    If tSuperPacmanTimer > 0 Then tSuperPacmanTimer = tSuperPacmanTimer + 1
        If tSuperPacmanTimer > 50 Then
            NoMoreSuperPacMan 'No More Super Pacman
        End If
    End If
        
End Sub

'######################################################################
'#                                                                    #
'# DoResize............                                               #
'# Calculate the positions of the screen and window when it's resized #
'#                                                                    #
'######################################################################
Private Sub DoResize()

   Dim x As Long
   Dim y As Long
   x = Screen.TwipsPerPixelX
   y = Screen.TwipsPerPixelY
   
   Dim ClientAreaH As Integer
   Dim ClientAreaW As Integer
   ClientAreaH = y * 48
   ClientAreaW = x * 5
    
   picGame.Height = g_screeny
   picGame.Width = g_screenx

   frmPacMan.Height = (g_screeny * y) + ClientAreaH
   frmPacMan.Width = (g_screenx * x) + ClientAreaW
   
   If g_InGame = False Then StartScreen
   
   DoEvents
   
End Sub

'######################################################################
'#                                                                    #
'# SetFood.............                                               #
'# When the game starts this function sets the food flags to true if  #
'# the map position is zero                                           #
'#                                                                    #
'######################################################################
Private Sub SetFood()
    
    Dim y As Integer
    Dim x As Integer
    Dim kk As Integer
    
    For y = 0 To SizeY
        For x = 0 To SizeX
            If PacMap(x, y) = 0 Then
                FoodMap(x, y) = True
                TotalNumOfFoods = TotalNumOfFoods + 1
            Else
                For kk = 0 To NUM_SPECIAL_TILES
                    'This sets the special (Power-up) foods on the screen
                    'The pSpecialTile structure holds the positions and is init'd on start-up
                    If pSpecialTile(kk).IDX = x Then
                        If pSpecialTile(kk).IDY = y Then
                            FoodMap(x, y) = True
                        End If
                    End If
                Next kk
            End If
        Next x
    Next y
    
End Sub

'######################################################################
'#                                                                    #
'# CheckHit............                                               #
'# Checks to see if a ghost has gotten Pacman or if pacman is         #
'# Powered-Up it checks to see if he's eaten a ghost                  #
'#                                                                    #
'######################################################################
Private Sub CheckHit()
    
    Dim kk As Integer
    If pPacPosition.pbSuperPacMan = False Then
    
        For kk = 0 To NUM_GHOSTS
            If Ghosts(kk).ghAlive = True Then
                If Ghosts(kk).ghSrcX = pPacPosition.pSrcX Then
                    If Ghosts(kk).ghSrcY = pPacPosition.pSrcY Then
                        End_Game
                    End If
                End If
            End If
        Next kk
    ElseIf pPacPosition.pbSuperPacMan = True Then 'SuperPacMan
        
        For kk = 0 To NUM_GHOSTS
            If Ghosts(kk).ghAlive = True Then
                If Ghosts(kk).ghSrcX = pPacPosition.pSrcX Then
                    If Ghosts(kk).ghSrcY = pPacPosition.pSrcY Then
                        Kill_Ghost kk
                    End If
                End If
            End If
        Next kk
    End If
    'Check hit on special foods
    For kk = 0 To NUM_SPECIAL_TILES
        If pPacPosition.pSrcX = pSpecialTile(kk).IDX Then
            If pPacPosition.pSrcY = pSpecialTile(kk).IDY Then
                
                'Set to off map so the powerup food is never shown after eating
                pSpecialTile(kk).IDX = -1
                pSpecialTile(kk).IDY = -1
                'Call the superpacman sub
                SuperPacMan
            End If
        End If
        
    Next kk
    
    
    For kk = 0 To NUM_GHOSTS
     'Release the ghost from the generation area
        If Ghosts(kk).gReleaseFlag = True Then
            Ghosts(kk).gReleaseCounter = Ghosts(kk).gReleaseCounter + 1
            If Ghosts(kk).gReleaseCounter > 100 Then
                If Ghosts(kk).ghAlive = False Then
                    If Ghosts(kk).ghSrcY = 7 And Ghosts(kk).ghSrcX = 13 Then
                        Ghosts(kk).ghSrcY = Ghosts(kk).ghSrcY - 2
                        Ghosts(kk).ghScreenX = Ghosts(kk).ghSrcX * tTileSize
                        Ghosts(kk).ghScreenY = Ghosts(kk).ghSrcY * tTileSize
                        Ghosts(kk).ghAlive = True
                        Ghosts(kk).gReleaseCounter = 0
                        Ghosts(kk).gReleaseFlag = False
                    End If
                End If
            End If
        End If
    Next kk
End Sub

'######################################################################
'#                                                                    #
'# Kill_Ghost............                                             #
'# PacMan ate a ghost now the ghost is removed from the map and       #
'# placed in the Re-Generation area ....                              #
'# We also give PacMan a nice fat bonus for eating the ghost!         #
'#                                                                    #
'######################################################################
Private Sub Kill_Ghost(Id As Integer)
    
    'Sound !!!!!!!!!!!!
    If bSND = True Then sndPlaySound App.Path & "\sounds\GHOSTEATEN.wav", SND_ASYNC
    
    Ghosts(Id).ghAlive = False
    Ghosts(Id).ghFreePTR = Id
    Ghosts(Id).ReGenerate_Ghosts
    'A nice fat bonus
    pPacPosition.pScore = pPacPosition.pScore + 20
    
End Sub


'######################################################################
'#                                                                    #
'# End_Game............                                               #
'# PacMan is dead....                                                 #
'# Show the end animation and display Game Over!!!!                   #
'#                                                                    #
'######################################################################
Private Sub End_Game()
    'Set the in game flag to false
    g_InGame = False

    
    If bSND = True Then sndPlaySound App.Path & "\sounds\killed.wav", 1
    
    'End anim
    Call StretchBlt(picGame.hdc, 0, 0, picGame.Width, picGame.Height, Bmps.hdc, 120, 20, 60, 60, SRCCOPY)
    Sleep 800
    Call StretchBlt(picGame.hdc, 0, 0, picGame.Width, picGame.Height, Bmps.hdc, 180, 20, 60, 60, SRCCOPY)
    Sleep 800
    Call StretchBlt(picGame.hdc, 0, 0, picGame.Width, picGame.Height, Bmps.hdc, 240, 20, 60, 60, SRCCOPY)
    'Wait a while
    Sleep 1000
          
    
End Sub

'######################################################################
'#                                                                    #
'# UpdateFood..........                                               #
'# Updates the food array when a piece of food is eaten then the flag #
'# is set to false...                                                 #
'#                                                                    #
'######################################################################
Private Sub UpdateFood()
    
    Dim y As Integer
    Dim x As Integer
    Dim kk As Integer
    
    For y = 0 To SizeY
        For x = 0 To SizeX
            If FoodMap(x, y) = True Then
                If PacMap(x, y) = 0 Then
                    BitBlt picBG.hdc, x * tTileSize, y * tTileSize, tTileSize, tTileSize, Bmps.hdc, 210, 0, vbSrcCopy
                     
                    'Check 'n blt special foods
                    For kk = 0 To NUM_SPECIAL_TILES
                        If pSpecialTile(kk).IDX = x Then
                            If pSpecialTile(kk).IDY = y Then
                                BitBlt picBG.hdc, x * tTileSize, y * tTileSize, tTileSize, tTileSize, Bmps.hdc, 370, 0, vbSrcCopy
                            End If
                        End If
                    Next kk
                    
                     
                End If
  
            End If
        Next x
    Next y
  
End Sub

'######################################################################
'#                                                                    #
'# SuperPacMan........                                                #
'# This does all of the stuff when Pacman eats a power-up food...     #
'#                                                                    #
'######################################################################
Private Sub SuperPacMan()
        pPacPosition.pbSuperPacMan = True
        g_SuperPacMan = True
        If bSND = True Then sndPlaySound App.Path & "\sounds\fruiteat.wav", SND_ASYNC
        
        'Set Gfx's to scared Ghosts
        Dim vv As Integer
        For vv = 0 To NUM_GHOSTS
            Ghosts(vv).ghColorPointer = 2
        Next vv
        
        tSuperPacmanTimer = 1
End Sub

'######################################################################
'#                                                                    #
'# NoMoreSuperPacMan()........                                        #
'# After the time is update this sets all of the game stuff back to   #
'# normal...If PacMan collides with a ghost after this he will be     #
'# killed!                                                            #
'#                                                                    #
'######################################################################
Private Sub NoMoreSuperPacMan()
        pPacPosition.pbSuperPacMan = False
        g_SuperPacMan = False
        'Set Gfx's back to Ghosts
        Ghosts(2).ghColorPointer = 4
        Ghosts(1).ghColorPointer = 4
        Ghosts(0).ghColorPointer = 0
        tSuperPacmanTimer = 0
End Sub

'######################################################################
'#                                                                    #
'# LevelUP....                                                        #
'# PacMan ate all the food and is still alive so do whatever is needed#
'# here to advance to the next level or give yourself a big pat on the#
'# back!                                                              #
'#                                                                    #
'######################################################################
Private Sub LevelUP()
    If bSND = True Then sndPlaySound App.Path & "\sounds\interm.wav", SND_ASYNC
    Sleep 3000
    g_InGame = False
End Sub
