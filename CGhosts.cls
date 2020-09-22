VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGhosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ghSrcX As Integer
Public ghSrcY As Integer
Public ghScreenX As Integer
Public ghScreenY As Integer
Public ghBltX As Integer
Public ghDirection As Integer
Public ghFrame As Integer
Public ghSpeedX As Integer
Public ghSpeedY As Integer
'Checks if the ghost is still alive
Public ghAlive As Boolean
'If Pacman eats a ghost we set a pointer to the ID of the eaten ghost
Public ghFreePTR As Integer
'A simple int that pointers to a ghosts color
Public ghColorPointer As Integer
Private Const SlowSpeed As Integer = 2
'Realse Timer counter
Public gReleaseCounter As Integer
'Relase boolean
Public gReleaseFlag As Boolean


Public Sub ReGenerate_Ghosts()
    
    ghColorPointer = 4 'Need to add randomness here
    ghSrcX = 7
    ghSrcY = 7
    ghScreenX = ghSrcX * tTileSize
    ghScreenY = ghSrcY * tTileSize
    ghDirection = 1
    ghFrame = 1
    gReleaseCounter = 0
    gReleaseFlag = True
    
End Sub

Public Sub Draw_Ghost()
    Call BitBlt(frmPacMan.picBG.hdc, ghScreenX, ghScreenY, tTileSize, tTileSize, frmPacMan.Bmps.hdc, ghBltX + (tTileSize * ghColorPointer), 0, vbSrcCopy)
   
End Sub

Public Sub Move_Ghost()
    
    Select Case ghDirection
    '******************* LEFT ********************
    Case 0:
          
        If PacMap(ghSrcX - 1, ghSrcY) = 0 Or PacMap(ghSrcX - 1, ghSrcY) = 38 Then
            If g_SuperPacMan = True Then

                ghSpeedX = -SlowSpeed
                ghSpeedY = 0
                
                ghDirection = 0
                
                If (ghScreenX - SlowSpeed) Mod tTileSize = 0 Then ghSrcX = ghSrcX - 1
            Else
                'Realign if not aligned correctly !!!!!!!
                ghScreenY = ghSrcY * tTileSize
                ghScreenX = ghSrcX * tTileSize
                
                ghSpeedX = -tTileSize
                ghSpeedY = 0
                
                ghSrcX = ghSrcX - 1
                ghDirection = 0
               
            End If
        Else
            ghSpeedX = 0
            ghSpeedY = 0
            ghDirection = Ghost_Random_Direction()
        End If
            
        ghBltX = 310
            
        If ghFrame = 1 Then
             ghBltX = ghBltX + tTileSize
            ghFrame = 0
        Else
            ghFrame = 1
        End If
    '******************* RIGHT ********************
    Case 1:
          
        If PacMap(ghSrcX + 1, ghSrcY) = 0 Or PacMap(ghSrcX + 1, ghSrcY) = 38 Then
            If g_SuperPacMan = True Then
                
                
                ghSpeedX = SlowSpeed
                ghSpeedY = 0
                ghDirection = 1
                If (ghScreenX + SlowSpeed) Mod tTileSize = 0 Then ghSrcX = ghSrcX + 1
            Else
            
                'Realign if not aligned correctly !!!!!!!
                ghScreenY = ghSrcY * tTileSize
                ghScreenX = ghSrcX * tTileSize
                
                ghSpeedX = tTileSize
                ghSpeedY = 0
                ghSrcX = ghSrcX + 1
                ghDirection = 1
            End If
        Else
            ghSpeedX = 0
            ghSpeedY = 0
            'Change Direction
            ghDirection = Ghost_Random_Direction()
        End If
            
        ghBltX = 310
            
        If ghFrame = 1 Then
             ghBltX = ghBltX + tTileSize
            ghFrame = 0
        Else
            ghFrame = 1
        End If

    '******************* UP ********************
    Case 2:
          
        If PacMap(ghSrcX, ghSrcY - 1) = 0 Or PacMap(ghSrcX, ghSrcY - 1) = 38 Then
           If g_SuperPacMan = True Then
               
                ghSpeedY = -SlowSpeed
                ghSpeedX = 0
                ghDirection = 2
                
                If (ghScreenY - SlowSpeed) Mod tTileSize = 0 Then ghSrcY = ghSrcY - 1
                
            Else
                           'Realign if not aligned correctly !!!!!!!
                ghScreenY = ghSrcY * tTileSize
                ghScreenX = ghSrcX * tTileSize
                
                ghSpeedY = -tTileSize
                ghSpeedX = 0
                ghSrcY = ghSrcY - 1
                ghDirection = 2
            End If
        Else
            ghSpeedX = 0
            ghSpeedY = 0
            'Change Direction
            ghDirection = Ghost_Random_Direction()
        End If
            
        ghBltX = 310
            
        If ghFrame = 1 Then
             ghBltX = ghBltX + tTileSize
            ghFrame = 0
        Else
            ghFrame = 1
        End If
    '******************* DOWN ********************
    Case 3:
          
        If PacMap(ghSrcX, ghSrcY + 1) = 0 Or PacMap(ghSrcX, ghSrcY + 1) = 38 Then
            If g_SuperPacMan = True Then
                
                ghSpeedY = SlowSpeed
                ghSpeedX = 0
                ghDirection = 3
                If (ghScreenY + SlowSpeed) Mod tTileSize = 0 Then ghSrcY = ghSrcY + 1
            Else
                
                'Re-align if not aligned correctly !!!!!!!
                ghScreenY = ghSrcY * tTileSize
                ghScreenX = ghSrcX * tTileSize
                ghSpeedY = tTileSize
                ghSpeedX = 0
                
                ghSrcY = ghSrcY + 1
                ghDirection = 3
                
            End If
            
        Else
            ghSpeedX = 0
            ghSpeedY = 0
            'Change Direction
            ghDirection = Ghost_Random_Direction()
        End If
            
        ghBltX = 310
            
        If ghFrame = 1 Then
            ghBltX = ghBltX + tTileSize
            ghFrame = 0
        Else
            ghFrame = 1
        End If


    End Select
    
    ghScreenY = ghScreenY + ghSpeedY
    ghScreenX = ghScreenX + ghSpeedX
    

    
End Sub

Private Function Ghost_Random_Direction() As Integer
    
    Dim i As Integer
    Dim NumberOfPaths As Integer
    NumberOfPaths = 4
    
    Randomize Timer

    i = Int((NumberOfPaths * Rnd))
    Ghost_Random_Direction = i
    
End Function


