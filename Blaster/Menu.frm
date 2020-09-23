VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BLASTER"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   FillColor       =   &H0000FFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   600
      Picture         =   "Menu.frx":0000
      ScaleHeight     =   83
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   0
      Top             =   13320
      Width           =   5310
   End
   Begin VB.Timer tmrExplode 
      Interval        =   200
      Left            =   3720
      Top             =   3480
   End
   Begin VB.Timer tmrRender 
      Interval        =   1
      Left            =   1320
      Top             =   3480
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private HighScore(4) As String
Private MenuChoice As Long
Private DisplayCredits As Boolean

'Menu selection
Private Sub Form_Click()
    Select Case MenuChoice
        Case 1 'New Game
            ShowCursor False
            frmMain.Show
            Unload Me
        
        Case 2 ' Credits
            DisplayCredits = Not DisplayCredits
        
        Case 3 ' Exit
            UnloadAll
    End Select
End Sub

Private Sub Form_Load()
Dim i As Long
    InitParticleEngine
    DisplayCredits = False
    On Error GoTo errOut
    Open "Blaster.hi" For Input As #1
    For i = 0 To 4
        Input #1, HighScore(i)
    Next i
errOut:
    Close #1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MenuChoice = 0
    If X < 125 Or X > 240 Then Exit Sub
    If Y >= 298 And Y < 322 Then MenuChoice = 1
    If Y >= 322 And Y < 346 Then MenuChoice = 2
    If Y >= 346 And Y < 370 Then MenuChoice = 3
End Sub

Private Sub Form_Terminate()
    UnloadAll
End Sub

'Render Menu Screen
Private Sub tmrRender_Timer()
Dim i As Long
    Menu.Cls
    RenderExplosion Menu
    BitBlt Menu.hdc, 30, 20, 350, 83, Picture1.hdc, 0, 0, vbSrcPaint
    
    'Hi Score or Credits?
    If DisplayCredits Then
        PutText "Programers", 135, 145, 12, 45311
        PutText "HLMODER", 135, 180, 16, 45311
        PutText "Paul Bahlawan", 135, 210, 16, 45311
    Else
        PutText "Top Defenders", 128, 135, 14, 45311
        For i = 0 To 4
            PutText HighScore(i), 150, 165 + i * 22, 12, 45311
        Next i
    End If
    
    'Menu choices
    PutText "New Game", 138, 298, 14, IIf(MenuChoice = 1, 16711776, vbBlue)
    PutText "Credits", 155, 322, 14, IIf(MenuChoice = 2, 16711776, vbBlue)
    PutText "Exit", 170, 346, 14, IIf(MenuChoice = 3, 16711776, vbBlue)
    
End Sub

Private Sub tmrExplode_Timer()
    NewExplosion Rnd * 350, Rnd * 350, Rnd * 255, Rnd * 255, Rnd * 255
End Sub

Private Sub PutText(txt As String, X As Long, Y As Long, Size As Long, Colr As Long)
    Me.CurrentX = X
    Me.CurrentY = Y
    Me.FontSize = Size
    Me.ForeColor = Colr
    Print txt
End Sub

