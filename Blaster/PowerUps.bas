Attribute VB_Name = "PowerUps"
Option Explicit

Private Type POWERUP
    X As Single
    Y As Single
    Width As Long
    Height As Long
    Lifetime As Long
    AniNumFrames As Long
    AniCurFrame As Long
    AniSpeed As Long
    AniDelay As Long
End Type



Public DoubleFire As POWERUP '2 Bullets
Public Minigun As POWERUP 'Minigun
Public Shield As POWERUP


Public HasDoubleFire As Boolean 'Has player 2 bullets?
Public HasMinigun As Boolean
Public HasShield As Boolean
Public HasNoPowerUp As Boolean


Public Sub InitPowerUps()
    DoubleFire.Lifetime = 0
    DoubleFire.AniCurFrame = 1
    DoubleFire.AniNumFrames = 5
    DoubleFire.AniSpeed = 2
    DoubleFire.Width = 25
    DoubleFire.Height = 25
    
    Minigun.Lifetime = 0
    Minigun.AniCurFrame = 1
    Minigun.AniNumFrames = 3
    Minigun.AniSpeed = 3
    Minigun.Width = 40
    Minigun.Height = 40
    
    Shield.Lifetime = 0
    Shield.AniCurFrame = 1
    Shield.AniNumFrames = 7
    Shield.AniSpeed = 2
    Shield.Width = 25
    Shield.Height = 25
    
    HasNoPowerUp = True
    HasMinigun = False
    HasDoubleFire = False
    HasShield = False
    
End Sub


Public Sub RenderPowerUps()

    'Double Fire
    If DoubleFire.Lifetime Then
        PutAnimatedPU frmMain.Screen, frmMain.PDoubleFire, DoubleFire ', DoubleFire.X, DoubleFire.Y , 25, 25
        DoubleFire.Lifetime = DoubleFire.Lifetime - 1
    End If

    'Minigun
    If Minigun.Lifetime Then
       'PutSprite frmMain.Screen, frmMain.Minigun, 1, Minigun.X, Minigun.Y, Minigun.Width, Minigun.Height
        PutAnimatedPU frmMain.Screen, frmMain.Minigun, Minigun
        Minigun.Lifetime = Minigun.Lifetime - 1
    End If
    
    'Shield
    If Shield.Lifetime Then
        PutAnimatedPU frmMain.Screen, frmMain.PShield, Shield
        Shield.Lifetime = Shield.Lifetime - 1
    End If

End Sub

Public Sub NewDoubleFirePower(ByVal X As Single, ByVal Y As Single)
    DoubleFire.X = X
    DoubleFire.Y = Y
    DoubleFire.Lifetime = 500
End Sub

Public Sub NewMinigunPower(ByVal X As Single, ByVal Y As Single)
    Minigun.X = X
    Minigun.Y = Y
    Minigun.Lifetime = 500
End Sub

Public Sub NewShieldPower(ByVal X As Single, ByVal Y As Single)
    Shield.X = X
    Shield.Y = Y
    Shield.Lifetime = 200
End Sub

'Animate Power-up Sprites
Public Sub PutAnimatedPU(Destination As Object, Source As Object, ByRef Id As POWERUP)
        
        Id.AniDelay = Id.AniDelay + 1
        
        If Id.AniDelay > Id.AniSpeed Then
            Id.AniDelay = 0
            Id.AniCurFrame = Id.AniCurFrame + 1
            If Id.AniCurFrame > Id.AniNumFrames Then
                Id.AniCurFrame = 1
            End If
        End If
        
        PutSprite Destination, Source, Id.AniCurFrame, Id.X, Id.Y, Id.Width, Id.Height

End Sub
