VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BLASTER"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":0000
   ScaleHeight     =   648.979
   ScaleMode       =   0  'User
   ScaleWidth      =   688.75
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PShield 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   5
      Left            =   7200
      Picture         =   "Main.frx":0F06
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   38
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PShield 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   4
      Left            =   6720
      Picture         =   "Main.frx":16B6
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   37
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PShield 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   3
      Left            =   6240
      Picture         =   "Main.frx":1E66
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   36
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PShield 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   2
      Left            =   5760
      Picture         =   "Main.frx":2616
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   35
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PShield 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   1
      Left            =   5280
      Picture         =   "Main.frx":2DC6
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   34
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PShield 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   4800
      Picture         =   "Main.frx":3574
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   33
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PlayerShooting 
      AutoRedraw      =   -1  'True
      Height          =   735
      Left            =   3480
      Picture         =   "Main.frx":3D22
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   32
      Top             =   8280
      Width           =   735
   End
   Begin VB.PictureBox Enemy5Mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   960
      Picture         =   "Main.frx":554E
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   31
      Top             =   10920
      Width           =   735
   End
   Begin VB.PictureBox Enemy5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   240
      Picture         =   "Main.frx":6D78
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   30
      Top             =   10920
      Width           =   735
   End
   Begin VB.TextBox PShieldPoints 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1080
      TabIndex        =   29
      Text            =   "Shield"
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox PHealth 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      TabIndex        =   28
      Text            =   "Health"
      Top             =   7680
      Width           =   975
   End
   Begin VB.PictureBox PlayerHit 
      AutoRedraw      =   -1  'True
      Height          =   735
      Left            =   2640
      Picture         =   "Main.frx":85A2
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   27
      Top             =   8280
      Width           =   735
   End
   Begin VB.PictureBox Enemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Index           =   6
      Left            =   6720
      Picture         =   "Main.frx":9DCE
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   26
      Top             =   10200
      Width           =   585
   End
   Begin VB.PictureBox Enemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Index           =   5
      Left            =   6120
      Picture         =   "Main.frx":ACD6
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   25
      Top             =   10200
      Width           =   585
   End
   Begin VB.PictureBox Enemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Index           =   4
      Left            =   5520
      Picture         =   "Main.frx":BBDE
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   24
      Top             =   10200
      Width           =   585
   End
   Begin VB.PictureBox Enemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Index           =   3
      Left            =   4920
      Picture         =   "Main.frx":CAE6
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   23
      Top             =   10200
      Width           =   585
   End
   Begin VB.PictureBox Enemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Index           =   2
      Left            =   4320
      Picture         =   "Main.frx":D9EE
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   22
      Top             =   10200
      Width           =   585
   End
   Begin VB.PictureBox Enemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Index           =   1
      Left            =   3720
      Picture         =   "Main.frx":E8F6
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   21
      Top             =   10200
      Width           =   585
   End
   Begin VB.PictureBox PlayerShield 
      AutoRedraw      =   -1  'True
      Height          =   735
      Left            =   1800
      Picture         =   "Main.frx":F7FE
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   20
      Top             =   8280
      Width           =   735
   End
   Begin VB.PictureBox BulletMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   2760
      Picture         =   "Main.frx":1102A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   19
      Top             =   10200
      Width           =   315
   End
   Begin VB.PictureBox Bullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   2400
      Picture         =   "Main.frx":113E0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   10200
      Width           =   315
   End
   Begin VB.PictureBox Enemy3mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   1680
      Picture         =   "Main.frx":11796
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   17
      Top             =   10200
      Width           =   660
   End
   Begin VB.PictureBox Enemy3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   960
      Picture         =   "Main.frx":12A98
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   16
      Top             =   10200
      Width           =   660
   End
   Begin VB.PictureBox MinigunMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   6840
      Picture         =   "Main.frx":13D9C
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   15
      Top             =   8280
      Width           =   660
   End
   Begin VB.PictureBox Minigun 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   2
      Left            =   6120
      Picture         =   "Main.frx":1509E
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   14
      Top             =   8280
      Width           =   660
   End
   Begin VB.PictureBox Minigun 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   1
      Left            =   5400
      Picture         =   "Main.frx":163A2
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   13
      Top             =   8280
      Width           =   660
   End
   Begin VB.PictureBox Minigun 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Index           =   0
      Left            =   4680
      Picture         =   "Main.frx":176A6
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   12
      Top             =   8280
      Width           =   660
   End
   Begin VB.PictureBox Enemy2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   240
      Picture         =   "Main.frx":189AA
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   11
      Top             =   10200
      Width           =   660
   End
   Begin VB.PictureBox PDoubleFireMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   3360
      Picture         =   "Main.frx":19CAE
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   10
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PDoubleFire 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   4
      Left            =   2760
      Picture         =   "Main.frx":1A45C
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   9
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PDoubleFire 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   3
      Left            =   2160
      Picture         =   "Main.frx":1AC0C
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   8
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PDoubleFire 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   2
      Left            =   1560
      Picture         =   "Main.frx":1B3BC
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   7
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PDoubleFire 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   1
      Left            =   960
      Picture         =   "Main.frx":1BB6C
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   6
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PDoubleFire 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   360
      Picture         =   "Main.frx":1C31C
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   5
      Top             =   9240
      Width           =   435
   End
   Begin VB.PictureBox PlayerMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   960
      Picture         =   "Main.frx":1CACC
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   8280
      Width           =   735
   End
   Begin VB.PictureBox Player 
      AutoRedraw      =   -1  'True
      Height          =   735
      Left            =   120
      Picture         =   "Main.frx":1E2F6
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   8280
      Width           =   735
   End
   Begin VB.PictureBox EnemyMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Left            =   7440
      Picture         =   "Main.frx":1FB22
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   2
      Top             =   10200
      Width           =   585
   End
   Begin VB.PictureBox Enemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Index           =   0
      Left            =   3120
      Picture         =   "Main.frx":20A28
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   1
      Top             =   10200
      Width           =   585
   End
   Begin VB.PictureBox Screen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   6  'Cross
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   503
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   551
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.Timer Shooter 
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
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   6720
         Top             =   4920
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Single
Dim Y As Single
Dim time As Single
Dim ShieldSpawnTime As Integer


Private Sub EnemySpawner_Timer()
ShieldSpawnTime = ShieldSpawnTime + 1

Randomize
NewSquad Rnd * 6, Rnd * 200, Rnd * -500

'Spawn a shield every X seconds
If ShieldSpawnTime = 5 Then
    NewShieldPower Rnd * 400, Rnd * 200
    ShieldSpawnTime = 0
End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If
End Sub


Private Sub Form_Load()
ShowCursor False
GameInit
End Sub



Private Sub Screen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If
End Sub



Private Sub Screen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shooter.Enabled = True
Shooting = True
End Sub

Private Sub Screen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CurX = X
CurY = Y
Player1.X = X
Player1.Y = Y
NewFlame Player1.X, Player1.Y
End Sub


Private Sub Screen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shooter.Enabled = False
Shooting = False
End Sub

Private Sub Shooter_Timer()
Shoot
End Sub

Private Sub Timer1_Timer()
'Render All!
Screen.Cls
RenderStars
RenderFlame
RenderPlayer
RenderPowerUps
MoveEnemies
RenderEnemies
MoveBullets
CheckCollisions
CheckEnemiesKilled
RenderExplosion Form1.Screen
End Sub
