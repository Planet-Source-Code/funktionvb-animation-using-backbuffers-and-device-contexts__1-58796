VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MarioMation"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CheckBox Check1 
         Caption         =   "Loop Animation"
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
      End
      Begin VB.PictureBox picDest 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3600
         Left            =   120
         ScaleHeight     =   238
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   318
         TabIndex        =   1
         Top             =   240
         Width           =   4800
      End
      Begin VB.TextBox txtSpeed 
         Height          =   285
         Left            =   5040
         TabIndex        =   7
         Text            =   "15"
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete DC"
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton cmdAnimate 
         Caption         =   "Animate Graphics"
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Graphics"
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Email: funktionill@lycos.com"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   4920
         TabIndex        =   10
         Top             =   3960
         Width           =   2430
      End
      Begin VB.Label Label3 
         Caption         =   "'P' = Pause Animation (picDest must have focus)"
         Height          =   615
         Left            =   5040
         TabIndex        =   9
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Set Animation Speed:"
         Height          =   255
         Left            =   5040
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Created by Shaun Holbach 2005"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3960
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code by: funktionvb
' Date: Monday 17th
' Hours worked: 5

'##########################################
' OK - So...after some gruling hours of work and learning
' I was finally able to grasp Device Contexts and BackBuffers

' I made this example for open source and learning
' if you feel i helped you out, please feel free to drop
' me an email: funktionill@lycos.com

' QUOTE:
' The concept of a program can be done in less than ten minutes.
' The difficult part is sitting at the computer for hours on end making it happen.

' HAPPY PROGRAMMING!

'##########################################


Option Explicit

' Declare Public Variables to store the DC's memory address
Dim DCBackground As Long
' Mario
Dim DCMask As Long
Dim DCSprite As Long
' Goomba
Dim DCGoomba As Long
Dim DCGoombaMask As Long

'our Buffer's DC
Public myBackBuffer As Long
Public myBufferBMP As Long

Const SpriteWidth As Long = 97
Const SpriteHeight As Long = 97
Const BGWid As Long = 320
Const BGHgt As Long = 240

' center of picDest
Dim DestMidX As Integer
Dim DestMidY As Integer

Dim Pause As Boolean                                 ' check if game is paused
Dim Running As Boolean                              ' check if game is started
Dim AnimSpeed As Long                              ' our animation speed
Dim isLoop As Boolean                                   ' loop animation?
      
' GetTickCount Timing Vars
Private gtcDesiredTime As Long
Private gtcStart As Long


Private Sub Check1_Click()

      Select Case Check1.Value
            Case 0      ' unchecked
                  isLoop = False
            Case 1      ' checked
                  isLoop = True
      End Select
      
      Debug.Print Check1.Value
      
End Sub

Private Sub Form_Load()
      With Form1
            .Visible = True
            .ScaleMode = 3      ' vbPixels
      End With
      
      ' disable buttons until first order completed
      txtSpeed.Enabled = False
      cmdAnimate.Enabled = False
      cmdDelete.Enabled = False
      
      ' Setup the animation speed
      txtSpeed = 65
      
      Running = False
      
      ' determine middle of my picture
      DestMidX = (picDest.ScaleWidth / 2) - (SpriteWidth / 2)
      DestMidY = picDest.ScaleHeight / 2 - (SpriteHeight / 2)

End Sub


Private Sub cmdLoad_Click()

      'create a compatable DC for the back buffer..
      myBackBuffer = CreateCompatibleDC(GetDC(0))
      
      'create a compatible bitmap surface for the DC
      'that is the size of our picDest.. (192 X 192)
      'NOTE - the bitmap will act as the actual graphics surface inside the DC
      'because without a bitmap in the DC, the DC cannot hold graphical data..
      myBufferBMP = CreateCompatibleBitmap(GetDC(0), BGWid, BGHgt)
      
      'final step of making the back buffer...
      'load our created blank bitmap surface into our buffer
      '(this will be used as our canvas to draw-on off screen)
      SelectObject myBackBuffer, myBufferBMP
      
      'load our sprites (using the function we made)
      ' MARIO
      DCMask = LoadGraphicDC(App.Path & "\images\mario_mask.bmp")
      DCSprite = LoadGraphicDC(App.Path & "\images\mario.bmp")
      ' GOOMBA
      DCGoomba = LoadGraphicDC(App.Path & "\images\goomba.bmp")
      DCGoombaMask = LoadGraphicDC(App.Path & "\images\goomba_mask.bmp")
      ' BACKDROP
      DCBackground = LoadGraphicDC(App.Path & "\images\BG.bmp")
      
      ' Draw the background from memory to backbuffer
      BitBlt myBackBuffer, 0, 0, BGWid, BGHgt, DCBackground, 0, 0, vbSrcCopy
      BitBlt picDest.hdc, 0, 0, BGWid, BGHgt, myBackBuffer, 0, 0, vbSrcCopy

      
      ' ok now all the graphics are loaded so
      ' enable the next order (setting the timer)
      txtSpeed.Enabled = True
      txtSpeed.SetFocus
      
      ' disable this button...keeps safe from mem leaks
      cmdLoad.Enabled = False
    
End Sub


Private Sub txtSpeed_GotFocus()
      txtSpeed.SelStart = 0
      txtSpeed.SelLength = Len(txtSpeed.Text)
      
      txtSpeed_Change
End Sub


Private Sub txtSpeed_Change()
    Dim TempStr As String
    
    TempStr = txtSpeed.Text
    
    ' test to see if there is a value
    If TempStr <> "" And TempStr <> "0" Then
        gtcDesiredTime = Val(TempStr)
        cmdAnimate.Enabled = True
    Else
        cmdAnimate.Enabled = False
        Exit Sub
    End If

End Sub


Private Sub cmdAnimate_Click()
      Randomize
      
      Dim intSound As Integer         ' for random sound
      
      ' enable deleteDC
      cmdDelete.Enabled = True
      ' disable animate
      cmdAnimate.Enabled = False
      
      ' allow for animation
      Running = True
      
      'set picdest focus
      picDest.SetFocus
      
      ' Play a random .wav file
      intSound = Int(Rnd * 3) + 1
      GameSoundPath = GetAppPath & "\sounds\"
      sndPlaySound GameSoundPath & intSound & ".wav", sndAsync
      
      ' run animation loop
      AnimateMario
End Sub


Private Sub AnimateMario()

      Dim Fps As Double                ' current fps
      Dim FrameCount As Long      ' count frames between fps update
      
      Dim iFrame As Integer           ' determines which frame on x-axis
      Dim iRow As Integer             ' determines if on y-axis
      
      iFrame = 1      ' (max=5)
      iRow = 0        ' (0=top row; 1=bottom row)
      
      While Running   ' is the game running?
      
            If gtcDesiredTime < GetTickCount - gtcStart Then
                  gtcStart = GetTickCount
                  
                  While Pause                             'is the game paused?
                        DoEvents                            'infinit loop
                  Wend
                  
                  ' Draw the background from memory to backbuffer
                  BitBlt myBackBuffer, 0, 0, BGWid, BGHgt, DCBackground, 0, 0, vbSrcCopy
                  
                  ' Draw the sprite from memory to the backbuffer
                  BitBlt myBackBuffer, DestMidX, DestMidY, 96, 96, DCMask, Int(iFrame * 96), Int(iRow * 96), vbSrcAnd
                  BitBlt myBackBuffer, DestMidX, DestMidY, 96, 96, DCSprite, Int(iFrame * 96), Int(iRow * 96), vbSrcPaint
                  
                  ' Blit Backbuffer to PrimaryBuffer (picDest.hc)
                  BitBlt picDest.hdc, 0, 0, BGWid, BGHgt, myBackBuffer, 0, 0, vbSrcCopy
            
                  ' increment animation counter (max=5)
                  iFrame = iFrame + 1
                  
                  If iFrame >= 5 Then
                        iFrame = 1
                        ' increment row
                        iRow = iRow + 1
                        
                        If iRow > 1 Then
                              ' check to loop animation
                              If isLoop = True Then
                                    Running = True
                              Else
                                    Running = False
                                    cmdAnimate.Enabled = True
                              End If
                              
                              iRow = 0
                        End If
                  End If
                  
                  
                  FrameCount = FrameCount + 1             'fps counter
                  
                  If Fps + 1000 <= GetTickCount Then
                        Me.Caption = "Fps: " & FrameCount   'Printing the actual fram count
                        Fps = GetTickCount
                        FrameCount = 0
                  End If
                  
            End If
            
            DoEvents                                     'handle win msgs
      
      Wend
      
      End Sub


Private Sub cmdDelete_Click()
      ' this clears up the memory we used to hold
      ' the graphics and the buffers we made
      DeleteObject DCMask
      DeleteObject DCSprite
      DeleteObject DCBackground
      
      'Delete the bitmap surface that was in the backbuffer
      DeleteObject myBufferBMP
      'Delete the backbuffer HDC
      DeleteDC myBackBuffer
      
      cmdAnimate.Enabled = False
      cmdDelete.Enabled = False
      cmdLoad.Enabled = True
      
      Running = False
      Pause = False
      
End Sub


Private Sub picDest_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyP Then                         'pause code
      
            If Pause = False Then
                  Pause = True
                  Me.Caption = "Paused"
            ElseIf Pause = True Then
                  Pause = False
            End If
          
      End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      ' may not be needed, but just in case
      Pause = False
      Running = False
      
      ' this clears up the memory we used to hold
      ' the graphics and the buffers we made
      DeleteObject DCMask
      DeleteObject DCSprite
      DeleteObject DCBackground
      
      'Delete the bitmap surface that was in the backbuffer
      DeleteObject myBufferBMP
      'Delete the backbuffer HDC
      DeleteDC myBackBuffer
End Sub


Public Function LoadGraphicDC(sFileName As String) As Long
      ' Device Context creation function

        ' Error handling
        On Error Resume Next
        
        ' temp variable to hold our DC address
        Dim LoadGraphicDCTEMP As Long
        
        ' create the DC address compatible with
        ' the DC of the screen
        LoadGraphicDCTEMP = CreateCompatibleDC(GetDC(0))
        
        ' load the graphic file into the DC
        SelectObject LoadGraphicDCTEMP, LoadPicture(sFileName)
        
        ' return the address of the file
        LoadGraphicDC = LoadGraphicDCTEMP
End Function

