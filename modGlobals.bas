Attribute VB_Name = "modGlobals"
Public Const tTileSize = 10
Public Const SizeX As Integer = 20
Public Const SizeY As Integer = 18
Public Const NumberOfTilesPerRow As Integer = 40
Public Const NumberOfRows As Integer = 1

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020             ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086            ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6              ' (DWORD) dest = source AND dest

Public g_screenx As Long
Public g_screeny As Long
Public PacMap(SizeX, SizeY) As Integer

Public NUM_GHOSTS As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_ASYNC = &H1        '  play asynchronously

Public g_SuperPacMan As Boolean
Public TotalNumOfFoods As Integer

Public g_LevelID As Integer
Public gMapPath As String
Public gMapName As String
Public Declare Function QueryPerformanceFrequencyAny Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Any) As Long
Public Declare Function QueryPerformanceCounterAny Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Any) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Const FPS As Integer = 5
Public g_bGameRunning As Boolean
