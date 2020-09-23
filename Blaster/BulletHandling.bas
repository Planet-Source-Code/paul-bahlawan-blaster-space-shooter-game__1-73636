Attribute VB_Name = "BulletHandling"
Option Explicit
Private Type Bullet
    X As Single
    Y As Single
    Vx As Single
    Vy As Single
    Alive As Boolean
    Damage As Long
End Type

Private Const BulletSpeed = 10
Private BulletTime As Long
Public Const MaxBullets As Long = 40
Public Bullets() As Bullet

Private Function FreeBullet() As Long 'Function to Find a "Free" Bullet
Dim i As Long

    For i = 0 To MaxBullets - 1
        If Bullets(i).Alive = False Then
            FreeBullet = i 'We found a free bullet
            Exit Function
        End If
    Next i

    FreeBullet = -1
    Debug.Print "noFB!";
    
End Function

Public Sub NewBullet(ByVal X As Single, ByVal Y As Single, Damage As Long, Frequency As Long)
Dim i As Long
    
    BulletTime = BulletTime + 1

    If BulletTime >= Frequency Then
    
        i = FreeBullet
        
        If i >= 0 Then
            Bullets(i).Alive = True
            Bullets(i).X = X
            Bullets(i).Y = Y
            Bullets(i).Damage = Damage
        End If
        
    BulletTime = 0
    End If
   
End Sub

Public Sub MoveBullets()
Dim i As Long
Dim col As Long
    For i = 0 To MaxBullets - 1
        If Bullets(i).Alive = True Then
            Bullets(i).Y = Bullets(i).Y - BulletSpeed
           
            SetPixel frmMain.Screen.hdc, Bullets(i).X, Bullets(i).Y, RGB(255, 64, 64)
            SetPixel frmMain.Screen.hdc, Bullets(i).X, Bullets(i).Y + 1, RGB(255, 64, 64)
            SetPixel frmMain.Screen.hdc, Bullets(i).X, Bullets(i).Y + 2, RGB(255, 255, 255)
            
                If Bullets(i).Y <= 0 Then
                    Bullets(i).Alive = False
                    'Bullets(i).Vy = 0
                End If
        End If
    Next i
        
End Sub

Public Sub InitBullets()
Dim i As Long

ReDim Bullets(MaxBullets - 1) As Bullet
    For i = 0 To MaxBullets - 1
        Bullets(i).Alive = False
    Next i
    
End Sub

Public Sub Shoot()

    If HasDoubleFire = True Then 'Player has Double Fire
        NewBullet CurX + 10, CurY, 10, 5
        NewBullet CurX + 32, CurY, 10, 5
    End If
    
    If HasMinigun = True Then 'Player has MiniGun
        NewBullet CurX, CurY, 15, 5 'Left Minigun
        NewBullet CurX + 45, CurY, 15, 5 'Right Minigun
    End If
    
    If HasNoPowerUp = True Then
        NewBullet CurX + 32, CurY, 10, 10
    End If
 
End Sub
