Attribute VB_Name = "Game"
Option Explicit
Public EnemiesKilled As Long 'How many Emeies has the player killed
Public EarthCore As Long  'EarthCore that player must Defend!!!
Public Score As Long
Private Type Message
    msg As String
    X As Long
    Y As Long
    Size As Long
    Colr As Long
    Lifetime As Long
End Type
Private Const MsgSlots = 3
Private ScreenMessage(MsgSlots) As Message
Private MsgNum As Long

Public CurX As Long
Public CurY As Long


Public Sub GameInit()
    EarthCore = 40 '40 health points to EarthCore
    EnemiesKilled = 0
    Score = 0
    Level = 0
    LevelMap = ""
    InitParticleEngine
    InitEnemies
    InitBullets
    InitPowerUps
    InitPlayer
    InitMessages
End Sub


Public Sub CheckCollisions()
Dim i As Long
Dim k As Long
    
'Enemy Collision With Player's Bullets
    For i = 0 To nEnemies - 1 'Loop all enemies array
        If Enemies(i).Alive = True Then 'Are they Alive?
            For k = 0 To MaxBullets - 1 'Loop all player Bullets
                If Bullets(k).Alive = True Then 'Are they Alive?
                    If Collision(1, 1, Enemies(i).Width, Enemies(i).Height, Bullets(k).X, Bullets(k).Y, Enemies(i).X, Enemies(i).Y) Then
                        If Enemies(i).Type < 10 Then '(not a bullet)
                            Enemies(i).Health = Enemies(i).Health - Bullets(k).Damage 'Enemy Loses Health
                            Bullets(k).Alive = False
                        End If
                    End If
                End If
            Next k
            
        'Enemy and Player Collide
            If Collision(Player1.Width, Player1.Height, Enemies(i).Width, Enemies(i).Height, Player1.X, Player1.Y, Enemies(i).X, Enemies(i).Y) Then
                Enemies(i).Health = Enemies(i).Health - 50
                If HasShield = True Then 'If the player has the shield powerup
                    Player1.Shield = Player1.Shield - Enemies(i).Damage
                Else
                    Player1.Health = Player1.Health - Enemies(i).Damage
                End If
                Player1.HitTime = 5
            End If
                
        End If
    Next i
    
'Gets DoubleFire Power Up
    If DoubleFire.Lifetime Then
        If Collision(45, 45, DoubleFire.Width, DoubleFire.Height, Player1.X, Player1.Y, DoubleFire.X, DoubleFire.Y) Then
            DoubleFire.Lifetime = 0
            HasDoubleFire = True
            HasNoPowerUp = False
        End If
    End If
    
'Gets minigun Power Up
    If Minigun.Lifetime Then
        If Collision(45, 45, Minigun.Width, Minigun.Height, Player1.X, Player1.Y, Minigun.X, Minigun.Y) Then
            Minigun.Lifetime = 0
            HasMinigun = True
            HasNoPowerUp = False
        End If
    End If
    
'Gets Shield Power Up
    If Shield.Lifetime Then
        If Collision(25, 25, Shield.Width, Shield.Height, Player1.X, Player1.Y, Shield.X, Shield.Y) Then
            Shield.Lifetime = 0
            HasShield = True
            Player1.Shield = Player1.Shield + 50
            If Player1.Shield > 100 Then Player1.Shield = 100
        End If
    End If
    
    'Check Players Health
    If Player1.Health <= 0 Then
        Player1.Health = 0
        EndGame
    End If
    
    'Check Players Shield
    If Player1.Shield <= 0 Then
        Player1.Shield = 0
        HasShield = False
    End If
 
    'If Earth is Destroyed the game is over!
    If EarthCore <= 0 Then
        EarthCore = 0
        EndGame
    End If
            
End Sub

Public Sub CheckEnemiesKilled()

    If HasDoubleFire = False Then
        If EnemiesKilled = 20 Then 'If player has killed 16 enemies put double fire powerup
            NewDoubleFirePower Rnd * 525, Rnd * 300
            EnemiesKilled = 21
        End If
    End If
    
    
    If HasMinigun = False Then
        If EnemiesKilled = 100 Then 'If player has killed 71 enemies put Minigun!
            NewMinigunPower Rnd * 500, Rnd * 300
            EnemiesKilled = 101
        End If
    End If

End Sub

Public Sub EndGame()
    Player1.Alive = False
    Player1.HitTime = 65
    frmMain.EnemySpawner.Enabled = False
    frmMain.tmrShoot.Enabled = False
    frmMain.Screen.BackColor = RGB(128, 0, 0)
    NewBigBlast Player1.X + 22, Player1.Y + 22
    NewExplosion Player1.X + 22, Player1.Y + 21, 255, 64, 64
    NewExplosion Player1.X + 22, Player1.Y + 22, 64, 255, 64
    NewExplosion Player1.X + 22, Player1.Y + 23, 64, 64, 255
End Sub

Public Sub UnloadAll()
    Unload frmMain
    Unload Intro
    Unload Menu
End Sub

Public Sub RenderStatus()
    'Health
    frmMain.Statusbar(0).Width = Player1.Health
    If frmMain.Statusbar(0).FillColor = vbYellow And Player1.Health < 25 Then
        newMessage "Warning", 10, 475, 14, vbRed, 123
    End If
    If frmMain.Statusbar(0).FillColor = vbGreen And Player1.Health < 50 Then
        newMessage "Caution", 10, 475, 14, vbYellow, 123
    End If
    If Player1.Health < 25 Then
        frmMain.Statusbar(0).FillColor = vbRed
    ElseIf Player1.Health < 50 Then
        frmMain.Statusbar(0).FillColor = vbYellow
    Else
        frmMain.Statusbar(0).FillColor = vbGreen
    End If
    
    'Shield
    frmMain.Statusbar(1).Width = Player1.Shield
    
    'Score
    frmMain.tbScore = Score
    
    'Earthcore
    frmMain.Statusbar(2).Width = EarthCore * 2.5
    If frmMain.Statusbar(2).FillColor = vbYellow And EarthCore < 10 Then
        newMessage "Danger", 440, 475, 14, vbRed, 123
    End If
    If frmMain.Statusbar(2).FillColor = vbGreen And EarthCore < 20 Then
        newMessage "Alert", 450, 475, 14, vbYellow, 123
    End If
    If EarthCore < 10 Then
        frmMain.Statusbar(2).FillColor = vbRed
    ElseIf EarthCore < 20 Then
        frmMain.Statusbar(2).FillColor = vbYellow
    Else
        frmMain.Statusbar(2).FillColor = vbGreen
    End If

End Sub

Private Sub InitMessages()
Dim i As Long
    For i = 0 To MsgSlots - 1
        ScreenMessage(i).Lifetime = 0
    Next i
End Sub

Public Sub newMessage(msg As String, X As Long, Y As Long, Size As Long, Colr As Long, Lifetime As Long)
    ScreenMessage(MsgNum).msg = msg
    ScreenMessage(MsgNum).Lifetime = Lifetime
    ScreenMessage(MsgNum).X = X
    ScreenMessage(MsgNum).Y = Y
    ScreenMessage(MsgNum).Size = Size
    ScreenMessage(MsgNum).Colr = Colr
    MsgNum = (MsgNum + 1) Mod MsgSlots
End Sub

Public Sub RenderMessage()
Dim i As Long
    For i = 0 To MsgSlots - 1
        If ScreenMessage(i).Lifetime Then
            frmMain.Screen.CurrentX = ScreenMessage(i).X
            frmMain.Screen.CurrentY = ScreenMessage(i).Y
            frmMain.Screen.FontSize = ScreenMessage(i).Size
            frmMain.Screen.ForeColor = ScreenMessage(i).Colr
            frmMain.Screen.Print ScreenMessage(i).msg
            ScreenMessage(i).Lifetime = ScreenMessage(i).Lifetime - 1
        End If
    Next i
End Sub

