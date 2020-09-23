VERSION 5.00
Begin VB.Form GameOver 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   FillColor       =   &H0000FFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInitials 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   675
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "GameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private HighScore(4) As String

Private Sub Form_Activate()
    
    PutText "GAME OVER", 50, 8, 24, vbRed

    If Player1.Health = 0 Then
        PutText "You have been defeated", 67, 64, 10, RGB(255, 128, 0)
    Else
        PutText "The Earth has been destroyed", 48, 64, 10, RGB(255, 128, 0)
    End If
    
    'a new high score?
    If Score > Val(HighScore(4)) Then
        Me.Height = 4000
        PutText "New High Score - Enter your initials", 30, 140, 10, RGB(255, 255, 0)
        txtInitials.Enabled = True
    End If

End Sub

Private Sub Form_Click()
    If txtInitials.Enabled Then
        SaveNewHighScore
    End If
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Long
    Me.Height = 2400
    On Error GoTo errOut
    Open "Blaster.hi" For Input As #1
    For i = 0 To 4
        Input #1, HighScore(i)
    Next i
errOut:
    Close #1
End Sub

Private Sub SaveNewHighScore()
Dim i As Long
Dim j As Long
    For i = 0 To 4 'where does the new score go?
        If Score > Val(HighScore(i)) Then
            Exit For
        End If
    Next i
        
    If i < 4 Then 'move old scores down to make space for the new score
        For j = 4 To i + 1 Step -1
            HighScore(j) = HighScore(j - 1)
        Next j
    End If
        
    'insert the new score
    HighScore(i) = Trim(Str$(Score)) & " " & txtInitials.Text

    On Error GoTo errOut
    Open "Blaster.hi" For Output As #1
    For i = 0 To 4
        Print #1, HighScore(i)
    Next i
errOut:
    Close #1

End Sub

Private Sub PutText(txt As String, X As Long, Y As Long, Size As Long, Colr As Long)
    Me.CurrentX = X
    Me.CurrentY = Y
    Me.FontSize = Size
    Me.ForeColor = Colr
    Me.Print txt
    
End Sub

Private Sub txtInitials_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Form_Click
End Sub
