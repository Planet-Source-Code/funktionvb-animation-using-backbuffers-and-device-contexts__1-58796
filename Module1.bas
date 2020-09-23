Attribute VB_Name = "Module1"
'The following API calls are for:

'blitting
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal X As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

'code timer
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, _
      ByVal lpsz As String, _
      ByVal un1 As Long, _
      ByVal n1 As Long, _
      ByVal n2 As Long, _
      ByVal un2 As Long) As Long
      
' LoadImage Constants
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTSIZE As Long = &H40

'for creating buffers / loading sprites
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'for loading sprites
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'for cleanup
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


' Playing a .wav file
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long

Public Const sndAsync = &H1
Public Const sndLoop = &H8
Public Const sndNoStop = &H10

Public GameSoundPath As String

Public Function GetAppPath()
      GetAppPath = App.Path
      
      If Right$(GetAppPath, 1) <> "\" Then
            GetAppPath = GetAppPath & "\"
      End If
End Function


