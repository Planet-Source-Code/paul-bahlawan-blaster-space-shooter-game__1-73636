Attribute VB_Name = "Enemy"
Option Explicit

Private Enum EnemyType
    Speeder = 1
    Charger = 2
    Disc = 3
    Blaster = 4
    Boss1 = 8
    Boss2 = 9
    Bullet = 10
    Bullet1 = 11
    HeatSeeker = 12
End Enum

Private Type Enemy
    X As Single
    Y As Single
    Vx As Single
    Vy As Single
    Width As Single
    Height As Single
    Alive As Boolean ' is this needed ?
    Health As Long
    ShootTime As Long
    Type As Long
    Damage As Long
    Value As Long
    AniNumFrames As Long
    AniCurFrame As Long
    AniSpeed As Long
    AniDelay As Long
End Type

Public Enemies() As Enemy
Private Squad(8) As String
Public Const nEnemies As Long = 60 '40

Public Level As Long
Private Wave As Long
Public LevelMap As String

Public Sub InitEnemies()
Dim i As Long
ReDim Enemies(nEnemies - 1) As Enemy

    For i = 0 To nEnemies - 1
        Enemies(i).Alive = False
    Next i
    
    Wave = 0

End Sub


Public Sub MoveEnemies()
Dim i As Long
Dim Theta As Single
Dim curTheta As Single
Dim cross As Single

    For i = 0 To nEnemies - 1
        If Enemies(i).Alive = True Then
            
            Select Case Enemies(i).Type
                '***********Enemy Charger IA***********
                Case EnemyType.Charger
            
                    Enemies(i).Y = Enemies(i).Y + Enemies(i).Vy
                    
                    Enemies(i).ShootTime = Enemies(i).ShootTime + 1
                     
                    If Enemies(i).ShootTime >= 60 Then 'Shoot A Bullet
                        NewEnemy Enemies(i).X + 12, Enemies(i).Y + 15, EnemyType.Bullet, 0, 6
                        Enemies(i).ShootTime = 0
                    End If
                    
                '***********Enemy Blaster IA***********
                Case EnemyType.Blaster
            
                    Enemies(i).Y = Enemies(i).Y + Enemies(i).Vy
                    Enemies(i).X = Enemies(i).X + Enemies(i).Vx
                    
                    Enemies(i).ShootTime = Enemies(i).ShootTime + 1
                     
                    If Enemies(i).ShootTime >= 60 Then 'Shoot A Bullet
                        NewEnemy Enemies(i).X + 8, Enemies(i).Y + 10, EnemyType.Bullet1
                        NewEnemy Enemies(i).X + 32, Enemies(i).Y + 10, EnemyType.Bullet1
                        Enemies(i).ShootTime = 0
                    End If
                     
                    If Enemies(i).Y > (Player1.Y - 150) Then 'When Enemy get close to player...
                    
                        'Calculate the path to colide with player
                        If Enemies(i).X > Player1.X Then
                            Enemies(i).Vx = Enemies(i).Vx - 0.22
                        Else
                            Enemies(i).Vx = Enemies(i).Vx + 0.22
                        End If
                        
                    End If
                     
                '*********Enemy Speeder IA***********
                Case EnemyType.Speeder
            
                    Enemies(i).Y = Enemies(i).Y + Enemies(i).Vy ' Move Enemy Down
                    
                '***********Enemy Disc IA***********
                Case EnemyType.Disc
            
                    Enemies(i).Y = Enemies(i).Y + Enemies(i).Vy ' Move Enemy Down
                    Enemies(i).X = Enemies(i).X + Cos(Enemies(i).Y * PI / 180)
                    
                    If Level > 1 Then 'disc can shoot on level 2 or higher
                        Enemies(i).ShootTime = Enemies(i).ShootTime + 1
                        If Enemies(i).ShootTime >= 123 Then 'Shoot A Bullet
                            NewEnemy Enemies(i).X + 17, Enemies(i).Y + 30, EnemyType.Bullet1
                            Enemies(i).ShootTime = 0
                        End If
                    End If
                
                '***********Enemy Boss1 IA***********
                Case EnemyType.Boss1
            
                    Enemies(i).Y = Enemies(i).Y + Enemies(i).Vy
                    Enemies(i).ShootTime = Enemies(i).ShootTime + 1
                     
                    If Enemies(i).ShootTime >= 16 Then 'Shoot A Bullet
                        Theta = Atan2(Player1.Y - Enemies(i).Y - 78, Player1.X - Enemies(i).X - 33)
                        NewEnemy Enemies(i).X + 55, Enemies(i).Y + 100, EnemyType.Bullet, 6 * Cos(Theta), 6 * Sin(Theta)
                        Enemies(i).ShootTime = 0
                    End If
                    
                    If Enemies(i).Health <= 0 Then
                        NewBigBlast Enemies(i).X + 60, Enemies(i).Y + 60
                        newMessage "1000", Enemies(i).X + 60, Enemies(i).Y + 60, 12, vbYellow, 100
                    End If
                
                '***********Enemy Boss2 IA***********
                Case EnemyType.Boss2
            
                    Enemies(i).Y = Enemies(i).Y + Enemies(i).Vy
                    Enemies(i).ShootTime = Enemies(i).ShootTime + 1
                     
                    If Enemies(i).ShootTime Mod 20 = 19 Then 'Shoot A Bullet
                        NewEnemy Enemies(i).X + 26, Enemies(i).Y + 84, EnemyType.Bullet, 0, 6
                        NewEnemy Enemies(i).X + 114, Enemies(i).Y + 84, EnemyType.Bullet, 0, 6
                        
                        If Enemies(i).ShootTime >= 240 Then 'Shoot a Heat Seeker
                            NewEnemy Enemies(i).X + 70, Enemies(i).Y + 100, EnemyType.HeatSeeker
                            Enemies(i).ShootTime = 0
                        End If
                    End If
                    
                    If Enemies(i).Health <= 0 Then
                        NewBigBlast Enemies(i).X + 76, Enemies(i).Y + 72
                        newMessage "2000", Enemies(i).X + 76, Enemies(i).Y + 72, 12, vbYellow, 100
                    End If
                
                '***********Enemy Bullet***********
                Case EnemyType.Bullet
            
                    Enemies(i).Y = Enemies(i).Y + Enemies(i).Vy
                    Enemies(i).X = Enemies(i).X + Enemies(i).Vx
                    If Enemies(i).X < 0 Or Enemies(i).X > 555 Or Enemies(i).Y < 0 Then
                        Enemies(i).Alive = False
                    End If
                    
                '***********Enemy Bullet1***********
                Case EnemyType.Bullet1
            
                    Enemies(i).Y = Enemies(i).Y + Enemies(i).Vy
                
                '***********Enemy Heat Seeker***********
                Case EnemyType.HeatSeeker
            
                    Enemies(i).Y = Enemies(i).Y + Enemies(i).Vy
                    Enemies(i).X = Enemies(i).X + Enemies(i).Vx
                    
                    If Enemies(i).X < -5 Or Enemies(i).X > 560 Or Enemies(i).Y < -5 Then
                        Enemies(i).Alive = False
                    End If
                    
                    'heat seeker runs out of fuel
                    Enemies(i).ShootTime = Enemies(i).ShootTime - 1
                    If Enemies(i).ShootTime = 0 Then
                        Enemies(i).Health = 0
                    End If
                    
                    'Find the current travel direction (angle) of the heat seeker
                    curTheta = Atan2(Enemies(i).Vy, Enemies(i).Vx)
                        
                    'Find the direction (angle) to the player
                    Theta = Atan2(Player1.Y - Enemies(i).Y + 16, Player1.X - Enemies(i).X + 16)
                    
                    'Find the best direction to turn (clockwise or counterclockwize)
                    cross = (Sin(Theta) * Cos(curTheta)) - (Cos(Theta) * Sin(curTheta))
                    
                    If cross < 0 Then 'Decide which way to turn
                        curTheta = curTheta - 0.04
                    Else
                        curTheta = curTheta + 0.04
                    End If
            
                    Enemies(i).Vx = 4 * Cos(curTheta) 'Convert Polar to Cartesian
                    Enemies(i).Vy = 4 * Sin(curTheta)
                
            End Select
            
            
            'Has Enemy reached Earth?
            If Enemies(i).Y > 520 Then
                Enemies(i).Alive = False
                If Enemies(i).Type < 10 Then '(bullets don't damage Earthcore)
                    EarthCore = EarthCore - 1 'EarthCore takes damage!
                End If
            End If
            
            'Has Enemy been killed?
            If Enemies(i).Health <= 0 Then 'Dead Enemy
                NewExplosion Enemies(i).X + (Enemies(i).Width / 2), Enemies(i).Y + (Enemies(i).Height / 2), Rnd * 255, Rnd * 255, Rnd * 255
                Enemies(i).Alive = False
                If Enemies(i).Type < 10 Then
                    EnemiesKilled = EnemiesKilled + 1 'One Enemy Killed!
                    Score = Score + Enemies(i).Value
                End If
            End If
            
        End If
    Next i
    
End Sub

Public Sub RenderEnemies()
Dim i As Long

    For i = 0 To nEnemies - 1
        If Enemies(i).Alive = True Then
            Select Case Enemies(i).Type
            
                Case EnemyType.Speeder
                    PutSprite frmMain.Screen, frmMain.Enemy1, 1, Enemies(i).X, Enemies(i).Y, Enemies(i).Width, Enemies(i).Height
                
                Case EnemyType.Charger
                    PutSprite frmMain.Screen, frmMain.Enemy2, 1, Enemies(i).X, Enemies(i).Y, Enemies(i).Width, Enemies(i).Height
                
                Case EnemyType.Disc
                    PutAnimatedEnmy frmMain.Screen, frmMain.Enemy3, Enemies(i)
                
                Case EnemyType.Blaster
                    PutSprite frmMain.Screen, frmMain.Enemy4, 1, Enemies(i).X, Enemies(i).Y, Enemies(i).Width, Enemies(i).Height
                
                Case EnemyType.Boss1
                    PutSprite frmMain.Screen, frmMain.Boss1, 1, Enemies(i).X, Enemies(i).Y, Enemies(i).Width, Enemies(i).Height
                
                Case EnemyType.Boss2
                    PutSprite frmMain.Screen, frmMain.Boss2, 1, Enemies(i).X, Enemies(i).Y, Enemies(i).Width, Enemies(i).Height
                
                Case EnemyType.Bullet
                    PutSprite frmMain.Screen, frmMain.Bullet, 1, Enemies(i).X, Enemies(i).Y, Enemies(i).Width, Enemies(i).Height
                
                Case EnemyType.Bullet1
                    SetPixel frmMain.Screen.hdc, Enemies(i).X, Enemies(i).Y, RGB(255, 64, 64)
                    SetPixel frmMain.Screen.hdc, Enemies(i).X, Enemies(i).Y - 1, RGB(255, 255, 255)
                    SetPixel frmMain.Screen.hdc, Enemies(i).X, Enemies(i).Y - 2, RGB(255, 255, 255)
                    SetPixel frmMain.Screen.hdc, Enemies(i).X - 1, Enemies(i).Y - 1, RGB(0, 255, 255)
                    SetPixel frmMain.Screen.hdc, Enemies(i).X + 1, Enemies(i).Y - 1, RGB(0, 255, 255)
                
                Case EnemyType.HeatSeeker
                    PutAnimatedEnmy frmMain.Screen, frmMain.Bullet, Enemies(i)
              
            End Select
        End If
    Next i

End Sub

Public Function FreeEnemy() As Long 'Function to Find a "Free" Enemy
Dim i As Long

    For i = 0 To nEnemies - 1
        If Enemies(i).Alive = False Then
            FreeEnemy = i
            Exit Function
        End If
    Next i

    FreeEnemy = -1
    Debug.Print "noFE!";
    
End Function

Public Sub NewEnemy(ByVal X As Single, ByVal Y As Single, ByVal Character As Long, Optional Vx As Single = 0, Optional Vy As Single = 0)
Dim i As Long

    i = FreeEnemy
    If i >= 0 Then
        Enemies(i).Alive = True
        Enemies(i).X = X
        Enemies(i).Y = Y
        Enemies(i).Type = Character
        
        Select Case Character
            Case EnemyType.Speeder
                Enemies(i).Vy = 3
                Enemies(i).Health = 10
                Enemies(i).Width = 40
                Enemies(i).Height = 40
                Enemies(i).Damage = 10
                Enemies(i).Value = 10
            
            Case EnemyType.Charger
                Enemies(i).Vy = 1
                Enemies(i).Health = 80
                Enemies(i).Width = 40
                Enemies(i).Height = 40
                Enemies(i).Damage = 15
                Enemies(i).Value = 20
                
            Case EnemyType.Disc
                Enemies(i).Vy = 1.5
                Enemies(i).Health = 60
                Enemies(i).Width = 35
                Enemies(i).Height = 35
                Enemies(i).Damage = 15
                Enemies(i).AniNumFrames = 7
                Enemies(i).AniCurFrame = 1
                Enemies(i).AniSpeed = 4
                Enemies(i).AniDelay = 0
                Enemies(i).Value = 30
                
            Case EnemyType.Blaster
                Enemies(i).Vx = 0
                Enemies(i).Vy = 2
                Enemies(i).Health = 30
                Enemies(i).Width = 38
                Enemies(i).Height = 31
                Enemies(i).Damage = 12
                Enemies(i).Value = 20
                
            Case EnemyType.Boss1
                Enemies(i).Vy = 0.4
                Enemies(i).Health = 3500
                Enemies(i).Width = 120
                Enemies(i).Height = 120
                Enemies(i).Damage = 15
                Enemies(i).Value = 1000
                                
            Case EnemyType.Boss2
                Enemies(i).Vy = 0.35
                Enemies(i).Health = 4000
                Enemies(i).Width = 152
                Enemies(i).Height = 124
                Enemies(i).Damage = 20
                Enemies(i).Value = 2000
            
            Case EnemyType.Bullet
                Enemies(i).Vx = Vx
                Enemies(i).Vy = Vy '6
                Enemies(i).Health = 1
                Enemies(i).Width = 17
                Enemies(i).Height = 17
                Enemies(i).Damage = 10
                
            Case EnemyType.Bullet1
                Enemies(i).Vy = 8
                Enemies(i).Health = 1
                Enemies(i).Width = 1
                Enemies(i).Height = 1
                Enemies(i).Damage = 5
            
            Case EnemyType.HeatSeeker
                Enemies(i).Vy = 4
                Enemies(i).Health = 1
                Enemies(i).Width = 17
                Enemies(i).Height = 17
                Enemies(i).Damage = 13
                Enemies(i).ShootTime = 260
                Enemies(i).AniNumFrames = 2
                Enemies(i).AniCurFrame = 1
                Enemies(i).AniSpeed = 3
                Enemies(i).AniDelay = 0
                
        End Select
                
    End If
   
End Sub


Public Sub InitSquadPositions(Formation As Long)

'1=Speeders
'2=Chargers
'3=Discs
'4=Blasters
'
'8=Boss1
'9=Boss2

    Select Case Formation
        Case 1
            Squad(1) = "     "
            Squad(2) = "     "
            Squad(3) = "     "
            Squad(4) = "1       1"
            Squad(5) = " 1     1"
            Squad(6) = "  1   1 "
            Squad(7) = "   1 1  "
            Squad(8) = "    1   "
        Case 2
            Squad(1) = "1   1"
            Squad(2) = "1   1"
            Squad(3) = " 1 1 "
            Squad(4) = "  1  "
            Squad(5) = "     "
            Squad(6) = "     "
            Squad(7) = "     "
            Squad(8) = "2 2 2"
        Case 3
            Squad(1) = "       "
            Squad(2) = "       "
            Squad(3) = "        "
            Squad(4) = "3      3"
            Squad(5) = "3      3"
            Squad(6) = "3      3"
            Squad(7) = "3      3"
            Squad(8) = "   22  "
        Case 4
            Squad(1) = "1   1"
            Squad(2) = "1   1"
            Squad(3) = " 1 1 "
            Squad(4) = "  1  "
            Squad(5) = "  1  "
            Squad(6) = "1   1"
            Squad(7) = "1   1"
            Squad(8) = "1   1"
        Case 5
            Squad(1) = "  33  "
            Squad(2) = "3    3"
            Squad(3) = "3    3"
            Squad(4) = "  33  "
            Squad(5) = "     "
            Squad(6) = "     "
            Squad(7) = "     "
            Squad(8) = "     "
            
        Case 6
            Squad(1) = "  33333"
            Squad(2) = "       "
            Squad(3) = "  4   4"
            Squad(4) = "       "
            Squad(5) = "4       4"
            Squad(6) = "       "
            Squad(7) = "     "
            Squad(8) = "     "
            
        Case 7
            Squad(1) = "   "
            Squad(2) = "   "
            Squad(3) = "   "
            Squad(4) = "   "
            Squad(5) = "     "
            Squad(6) = "  8  "
            Squad(7) = "     "
            Squad(8) = "     "
        
        Case 8
            Squad(1) = "   "
            Squad(2) = "   "
            Squad(3) = "   "
            Squad(4) = "   "
            Squad(5) = "     "
            Squad(6) = "  9  "
            Squad(7) = "     "
            Squad(8) = "     "
 
    End Select

End Sub

Public Sub NewSquad(Formation As Long, PosX As Single, PosY As Single)
Dim X As Long
Dim Y As Long
Dim t As Long
    
    If Formation = 0 Then Exit Sub
    
    InitSquadPositions Formation

    For Y = 1 To 8
        For X = 1 To Len(Squad(Y))
            t = Val(Mid$(Squad(Y), X, 1))
            If t > 0 Then
                NewEnemy X * 40 + PosX, Y * 40 + PosY, t
            End If
        Next X
    Next Y
              
End Sub

Public Sub NextWave()
    Wave = Wave + 1
    If Wave > Len(LevelMap) Then 'next level?
        NextLevel
        Wave = 1
    End If
    
    NewSquad Val(Mid$(LevelMap, Wave, 1)), Rnd * 170, -360

End Sub

Public Sub NextLevel()
    Level = Level + 1
    If Level > 2 Then 'we only have 2 levels.... for now
        Level = 1
        frmMain.EnemySpawner.Interval = frmMain.EnemySpawner.Interval - 450 'Speed it up
    End If
        
    newMessage "LEVEL " & Str$(Level), 200, 190, 18, vbGreen, 150
    
    '0 = no squad
    '1 = squad 1
    '2 = squad 2
    'etc
    Select Case Level
        Case 1
            LevelMap = "01112131451627000111454361554617070000"
        Case 2
            LevelMap = "02256132451628003546543651536180800000"
    End Select
            
End Sub

'Animate Enemy Sprites
Public Sub PutAnimatedEnmy(Destination As Object, Source As Object, ByRef Id As Enemy)
        
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
