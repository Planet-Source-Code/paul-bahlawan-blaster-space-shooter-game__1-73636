Attribute VB_Name = "Player"
Option Explicit

Private Type Player
    X As Single
    Y As Single
    Width As Long
    Height As Long
    Health As Long
    Shield As Long
    HitTime As Long
    Alive As Boolean
    Shooting As Boolean
End Type

Public Player1 As Player


Public Sub RenderPlayer()
    'Render Player...
    If Player1.HitTime And Not HasShield Then
        'Damaged player
        PutSprite frmMain.Screen, frmMain.Player, 3, Player1.X, Player1.Y, Player1.Width, Player1.Height
        Player1.HitTime = Player1.HitTime - 1
    Else
        If Player1.Shooting Then
        'Player shooting
            PutSprite frmMain.Screen, frmMain.Player, 2, Player1.X, Player1.Y, Player1.Width, Player1.Height
        Else
        'Player not shooting
            PutSprite frmMain.Screen, frmMain.Player, 1, Player1.X, Player1.Y, Player1.Width, Player1.Height
        End If
    End If
    
    'Render MiniGun
    If HasMinigun = True Then
        'Left MG
        If Player1.Shooting Then
            Minigun.X = Player1.X - 20
            Minigun.Y = Player1.Y
            PutAnimatedPU frmMain.Screen, frmMain.Minigun, Minigun
        Else
            PutSprite frmMain.Screen, frmMain.Minigun, Minigun.AniCurFrame, Player1.X - 20, Player1.Y, Minigun.Width, Minigun.Height
        End If
        
        'Right MG
        PutSprite frmMain.Screen, frmMain.Minigun, Minigun.AniCurFrame, Player1.X + 25, Player1.Y, Minigun.Width, Minigun.Height
     End If
     
    'Render Shield on player
    If HasShield Then
        If Player1.HitTime Then
            'Damaged shield
            PutSprite frmMain.Screen, frmMain.PlayerShield, 2, Player1.X, Player1.Y, Player1.Width, Player1.Height
            Player1.HitTime = Player1.HitTime - 1
        Else
            'Regular shield
            PutSprite frmMain.Screen, frmMain.PlayerShield, 1, Player1.X, Player1.Y, Player1.Width, Player1.Height
        End If
    End If

End Sub

Public Sub InitPlayer()
    Player1.X = 300
    Player1.Y = 400
    Player1.Width = 45
    Player1.Height = 45
    Player1.Health = 100
    Player1.Shield = 0
    Player1.Alive = True
    Player1.HitTime = 0
End Sub
