Attribute VB_Name = "ParticleEngine"
Option Explicit

Private Type PARTICLE
    X As Single
    Y As Single
    Vx As Single
    Vy As Single
    Color As Long
    Lifetime As Long
End Type

Private Const nParticles As Long = 180

Private Type EXPLOSION
    nParticle(nParticles) As PARTICLE
    Lifetime As Single
End Type

Private BigBlast(1) As EXPLOSION
Private ShockWave() As EXPLOSION
Private Flame() As PARTICLE
Private Stars() As PARTICLE
Private Const nFlames As Long = 40
Private Const nExplosions As Long = 15
Private Const nStars As Long = 30

Public Sub InitParticleEngine()

Dim i As Long
Dim k As Long

    'Regular Explosion
    ReDim ShockWave(nExplosions - 1) As EXPLOSION
    For i = 0 To nExplosions - 1
        ShockWave(i).Lifetime = 0
    
        For k = 0 To nParticles - 1
            ShockWave(i).nParticle(k).Color = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
            ShockWave(i).nParticle(k).Vx = Cos((k * 2) * PI / 180) * Rnd * 20
            ShockWave(i).nParticle(k).Vy = Sin((k * 2) * PI / 180) * Rnd * 20
        Next k
    Next i
    
    'Big Blast Ring
    BigBlast(0).Lifetime = 0
    BigBlast(1).Lifetime = 0
    For k = 0 To nParticles - 1
        BigBlast(0).nParticle(k).Color = RGB(64, 128, 192)
        BigBlast(1).nParticle(k).Color = RGB(255, 128, 64)
'        If k Mod 2 Then
            BigBlast(0).nParticle(k).Vx = Cos((k * 2) * PI / 180) * 7
            BigBlast(0).nParticle(k).Vy = Sin((k * 2) * PI / 180) * 7
'        Else
'            BigBlast(0).nParticle(k).Vx = Cos((k * 2) * PI / 180) * 8.2
'            BigBlast(0).nParticle(k).Vy = Sin((k * 2) * PI / 180) * 8.2
'        End If
        BigBlast(1).nParticle(k).Vx = Cos((k * 2) * PI / 180) * 18
        BigBlast(1).nParticle(k).Vy = Sin((k * 2) * PI / 180) * 18
    Next k

    
    'Background Stars
    ReDim Stars(nStars - 1) As PARTICLE
    For i = 0 To nStars - 1
        Stars(i).X = Rnd * 500
        Stars(i).Y = Rnd * 500
        Stars(i).Vy = 1 + (Rnd * 2)
        Stars(i).Color = RGB(128 + Int(Rnd * 128), 128 + Int(Rnd * 128), 128 + Int(Rnd * 128))
    Next i
    
    ReDim Flame(nFlames - 1) As PARTICLE
End Sub


Public Sub RenderExplosion(Picture As Object)
Dim i As Long
Dim k As Long

    'Regular explosion
    For i = 0 To nExplosions - 1
        If ShockWave(i).Lifetime Then
            For k = 0 To nParticles - 1
                  ShockWave(i).nParticle(k).X = (ShockWave(i).nParticle(k).X + ShockWave(i).nParticle(k).Vx)
                  ShockWave(i).nParticle(k).Y = (ShockWave(i).nParticle(k).Y - ShockWave(i).nParticle(k).Vy)
                  SetPixel Picture.hdc, ShockWave(i).nParticle(k).X, ShockWave(i).nParticle(k).Y, ShockWave(i).nParticle(k).Color
            Next k
            ShockWave(i).Lifetime = ShockWave(i).Lifetime - 1
        End If
    Next i
    
    'Big Blast ring
    For i = 0 To 1
        If BigBlast(i).Lifetime Then
            For k = 0 To nParticles - 1
                  BigBlast(i).nParticle(k).X = (BigBlast(i).nParticle(k).X + BigBlast(i).nParticle(k).Vx)
                  BigBlast(i).nParticle(k).Y = (BigBlast(i).nParticle(k).Y - BigBlast(i).nParticle(k).Vy)
                  SetPixel Picture.hdc, BigBlast(i).nParticle(k).X, BigBlast(i).nParticle(k).Y, BigBlast(i).nParticle(k).Color
            Next k
            BigBlast(i).Lifetime = BigBlast(i).Lifetime - 1
        End If
    Next i
End Sub

Public Sub NewExplosion(ByVal X As Single, ByVal Y As Single, ByVal r As Long, ByVal G As Long, ByVal B As Long)
Dim i As Long
Dim k As Long

    i = FreeExplosion
    
    If i >= 0 Then
        ShockWave(i).Lifetime = 33
            For k = 0 To nParticles - 1
                ShockWave(i).nParticle(k).Color = RGB(r, G, B)
                ShockWave(i).nParticle(k).X = X
                ShockWave(i).nParticle(k).Y = Y
            Next k
    End If
End Sub


Private Function FreeExplosion() As Long
Dim i As Long

  For i = 0 To nExplosions - 1
        If ShockWave(i).Lifetime = 0 Then
            FreeExplosion = i
            Exit Function
        End If
    Next i
    
    FreeExplosion = -1
    Debug.Print "noFX!";
    
End Function


Public Sub NewBigBlast(ByVal X As Single, ByVal Y As Single)
Dim k As Long
    BigBlast(0).Lifetime = 55
    BigBlast(1).Lifetime = 40
    
    For k = 0 To nParticles - 1
        BigBlast(0).nParticle(k).X = X
        BigBlast(0).nParticle(k).Y = Y
        BigBlast(1).nParticle(k).X = X
        BigBlast(1).nParticle(k).Y = Y
    Next k
End Sub


'Stars Scroll
Public Sub RenderStars()
Dim i As Long

    For i = 0 To nStars - 1
        Stars(i).Y = Stars(i).Y + Stars(i).Vy
        
        If Stars(i).Y > 500 Then
            Stars(i).Y = 0
        End If
        SetPixel frmMain.Screen.hdc, Stars(i).X, Stars(i).Y, Stars(i).Color
    Next i
    
End Sub

'Flame
Public Sub NewFlame(ByVal X As Single, ByVal Y As Single)
Dim i As Long

    i = FreeFlame
    
    If i >= 0 Then
        Flame(i).X = X
        Flame(i).Y = Y
        Flame(i).Lifetime = 20
    End If
    
End Sub

Private Function FreeFlame() As Long
Dim i As Long

    For i = 0 To nFlames - 1
        If Flame(i).Lifetime = 0 Then
            FreeFlame = i
            Exit Function
        End If
    Next i
    
    FreeFlame = -1

End Function

Public Sub RenderFlame()
Dim i As Long

    For i = 0 To nFlames - 1
        
        If Flame(i).Lifetime Then
        
            Flame(i).Y = Flame(i).Y + 1
            Flame(i).X = Flame(i).X + Cos(i * 3)
            
            Flame(i).Lifetime = Flame(i).Lifetime - 1
        
'            SetPixel frmMain.Screen.hdc, Flame(i).X + 20, Flame(i).Y + 35, RGB(255 - (Flame(i).Lifetime * 3), Flame(i).Lifetime, i * 4)
            SetPixel frmMain.Screen.hdc, Flame(i).X + 22, Flame(i).Y + 35, RGB(150 - (Flame(i).Lifetime * 2), Flame(i).Lifetime * 6, i * 2)
            SetPixel frmMain.Screen.hdc, Flame(i).X + 22, Flame(i).Y + 36, RGB(50 - (Flame(i).Lifetime * 2), Flame(i).Lifetime * 2, i * 7)
        End If
    Next i
    
End Sub
