VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPacMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public pSrcX As Integer
Public pSrcY As Integer
Public pScreenX As Integer
Public pScreenY As Integer
Public pBltX As Integer
Public pDirection As Integer
Public pFrame As Integer
Public pSpeedX As Integer
Public pSpeedY As Integer
Public pScore As Integer
Public pLives As Integer
Public pbSuperPacMan As Boolean


Public Sub Draw_PacMan()
    BitBlt frmPacMan.picBG.hdc, pScreenX, pScreenY, tTileSize, tTileSize, frmPacMan.Bmps.hdc, pBltX, 0, vbSrcCopy
    Set_PacMan_Score
    Set_PacMan_Lives
End Sub

Private Sub Set_PacMan_Score()
    Call TextOut(frmPacMan.picBG.hdc, 0, 0, "Score", 5)
    
    Dim nlen As Integer
    If pScore < 10 Then
        nlen = 1
    ElseIf pScore >= 10 And pScore < 100 Then
        nlen = 2
    ElseIf pScore >= 100 Then
        nlen = 3
    End If
    
    Call TextOut(frmPacMan.picBG.hdc, 35, 0, pScore, nlen)
    
End Sub

Private Sub Set_PacMan_Lives()
    Call TextOut(frmPacMan.picBG.hdc, 70, 0, "Lives", 5)
    Call TextOut(frmPacMan.picBG.hdc, 100, 0, pLives, 1)
End Sub
