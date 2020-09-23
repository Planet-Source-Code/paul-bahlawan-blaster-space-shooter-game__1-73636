Attribute VB_Name = "Engine"
Option Explicit

'Draw sprite with mask
Public Sub PutSprite(Destination As Object, Source As Object, Index As Long, ByVal X As Single, ByVal Y As Single, ByVal ResX As Single, ByVal ResY As Single)
    BitBlt Destination.hdc, X, Y, ResX, ResY, Source.hdc, 0, 0, vbSrcAnd
    BitBlt Destination.hdc, X, Y, ResX, ResY, Source.hdc, (Index * ResX), 0, vbSrcPaint
End Sub

'Square collision detection
Public Function Collision(ByVal Width1 As Single, ByVal Height1 As Single, ByVal Width2 As Single, ByVal Height2 As Single, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Boolean
    If ((X1 + Width1) > X2 And (Y1 + Height1) > Y2 And (X2 + Width2) > X1 And (Y2 + Height2) > Y1) Then
        Collision = True
    End If
End Function

'ATAN2 - math function
Public Function Atan2(ByVal Y As Single, ByVal X As Single) As Single
    If Y > 0 Then
        If X >= Y Then
            Atan2 = Atn(Y / X)
        ElseIf X <= -Y Then
            Atan2 = Atn(Y / X) + PI
        Else
            Atan2 = PI / 2 - Atn(X / Y)
        End If
    Else
        If X >= -Y Then
            Atan2 = Atn(Y / X)
        ElseIf X <= Y Then
            Atan2 = Atn(Y / X) - PI
        Else
            Atan2 = -Atn(X / Y) - PI / 2
        End If
    End If
End Function
 


'Public Function Dist(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
'    Dist = Sqr((X2 - X1) * (X2 - X1)) + ((Y2 - Y1) * (Y2 - Y1))
'End Function


