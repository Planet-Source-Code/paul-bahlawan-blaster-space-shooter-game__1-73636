VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BLASTER"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":08DA
   ScaleHeight     =   648.979
   ScaleMode       =   0  'User
   ScaleWidth      =   688.75
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tbShield 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00CC7744&
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Text            =   "Shield"
      Top             =   7680
      Width           =   495
   End
   Begin VB.PictureBox Boss1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1860
      Left            =   960
      Picture         =   "Main.frx":17E0
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   14
      Top             =   10878
      Width           =   3660
   End
   Begin VB.TextBox tbScore 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4560
      TabIndex        =   13
      Text            =   "Score"
      Top             =   7607
      Width           =   975
   End
   Begin VB.TextBox tbEarth 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6360
      TabIndex        =   12
      Text            =   "Earth"
      Top             =   7680
      Width           =   495
   End
   Begin VB.PictureBox PShield 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   3720
      Picture         =   "Main.frx":169A2
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   11
      Top             =   10878
      Width           =   3060
   End
   Begin VB.PictureBox Enemy4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   525
      Left            =   120
      Picture         =   "Main.frx":1A47C
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   10
      Top             =   10878
      Width           =   1200
   End
   Begin VB.TextBox tbHealth 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Text            =   "Health"
      Top             =   7680
      Width           =   495
   End
   Begin VB.PictureBox PlayerShield 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   120
      Picture         =   "Main.frx":1C05A
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   8
      Top             =   10878
      Width           =   2085
   End
   Begin VB.PictureBox Bullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   2040
      Picture         =   "Main.frx":20854
      ScaleHeight     =   255
      ScaleWidth      =   765
      TabIndex        =   7
      Top             =   10878
      Width           =   825
   End
   Begin VB.PictureBox Enemy3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Left            =   120
      Picture         =   "Main.frx":212F2
      ScaleHeight     =   525
      ScaleWidth      =   4200
      TabIndex        =   6
      Top             =   10878
      Width           =   4260
   End
   Begin VB.PictureBox Minigun 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   3720
      Picture         =   "Main.frx":2860C
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   5
      Top             =   10878
      Width           =   2460
   End
   Begin VB.PictureBox Enemy2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   120
      Picture         =   "Main.frx":2D14E
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   4
      Top             =   10878
      Width           =   1260
   End
   Begin VB.PictureBox PDoubleFire 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   3720
      Picture         =   "Main.frx":2F710
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   10878
      Width           =   2310
   End
   Begin VB.PictureBox Player 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   120
      Picture         =   "Main.frx":32376
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   10878
      Width           =   2760
   End
   Begin VB.PictureBox Enemy1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   120
      Picture         =   "Main.frx":382A4
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   1
      Top             =   10878
      Width           =   1260
   End
   Begin VB.PictureBox Screen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      FillStyle       =   6  'Cross
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   505
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   0
      Top             =   37
      Width           =   8295
      Begin VB.PictureBox Boss2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1920
         Left            =   240
         Picture         =   "Main.frx":3A866
         ScaleHeight     =   124
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   304
         TabIndex        =   16
         Top             =   13320
         Width           =   4620
      End
      Begin VB.Timer tmrShoot 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7320
         Top             =   5640
      End
      Begin VB.Timer EnemySpawner 
         Interval        =   3500
         Left            =   7320
         Top             =   4920
      End
      Begin VB.Timer tmrRenderAll 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   6720
         Top             =   4920
      End
   End
   Begin VB.Shape Statusbar 
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   2
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   7742
      Width           =   1200
   End
   Begin VB.Shape Statusbar 
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CC7744&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   1
      Left            =   2640
      Shape           =   4  'Rounded Rectangle
      Top             =   7742
      Width           =   1200
   End
   Begin VB.Shape Statusbar 
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   0
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   7742
      Width           =   1200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShieldSpawnTime As Long


Private Sub EnemySpawner_Timer()
    
    NextWave
    
    'Spawn a shield every X seconds
    ShieldSpawnTime = ShieldSpawnTime + 1
    
    If ShieldSpawnTime = 7 Then
        NewShieldPower Rnd * 525, Rnd * 300
        ShieldSpawnTime = 0
    End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        ShowCursor True
        Unload Me
        Menu.Show
    End If
    If KeyCode = 80 Then 'Press P to pause/unpause
        tmrRenderAll.Enabled = Not tmrRenderAll.Enabled
        EnemySpawner.Enabled = Not EnemySpawner.Enabled
        tmrShoot.Enabled = False
    End If
    
End Sub


Private Sub Form_Load()
    Randomize
    GameInit
    newMessage "GET READY", 180, 240, 18, vbRed, 150
    ShieldSpawnTime = 0
    NextWave
    tmrRenderAll.Enabled = True

End Sub


Private Sub Form_Terminate()
    ShowCursor True
    UnloadAll
End Sub


Private Sub Screen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If EnemySpawner.Enabled Then
        tmrShoot.Enabled = True
        Player1.Shooting = True
    End If
End Sub


Private Sub Screen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y > 500 Then Y = 500
    CurX = X - 22
    CurY = Y - 22
    Player1.X = CurX
    Player1.Y = CurY
    NewFlame Player1.X, Player1.Y
End Sub


Private Sub Screen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrShoot.Enabled = False
    Player1.Shooting = False
End Sub


Private Sub tmrShoot_Timer()
    Shoot
End Sub


Private Sub tmrRenderAll_Timer()
    If Player1.Alive Then
        'Render All!
        Screen.Cls
        RenderStars
        RenderFlame
        RenderMessage
        RenderPlayer
        RenderPowerUps
        MoveEnemies
        RenderEnemies
        MoveBullets
        CheckCollisions
        CheckEnemiesKilled
        RenderExplosion frmMain.Screen
        RenderStatus
    Else
        'Player death sequence
        frmMain.Screen.BackColor = 0
        RenderEnemies
        RenderExplosion frmMain.Screen
        Player1.HitTime = Player1.HitTime - 1
        If Player1.HitTime <= 0 Then
            ShowCursor True
            GameOver.Show vbModal
            Unload frmMain
            Menu.Show
        End If
    End If

End Sub
