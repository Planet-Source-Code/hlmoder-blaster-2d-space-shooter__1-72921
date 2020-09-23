VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BLASTER"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   1455
      Left            =   600
      Picture         =   "Menu.frx":0000
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   3
      Top             =   8640
      Width           =   4935
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1455
      Left            =   600
      Picture         =   "Menu.frx":15556
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   2
      Top             =   6840
      Width           =   5055
   End
   Begin VB.Timer Timer3 
      Interval        =   200
      Left            =   3720
      Top             =   3480
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1320
      Top             =   3480
   End
   Begin VB.Label Credits 
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label NewGame 
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Exit 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   4560
      Width           =   495
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private musicsound As Long
Private musicchannel As Long

Private Sub Credits_Click()
MsgBox "Programed by HLMODER"
End Sub

Private Sub Exit_Click()
UnloadAll
End Sub

Private Sub Form_Load()
Unload Form1
ShowCursor True
InitParticleEngine
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Intro
Unload Me
End Sub

Private Sub NewGame_Click()
Load Form1
Form1.Timer1.Enabled = True
Form1.Show
ShowCursor False
Unload Me
Unload Intro
End Sub

Private Sub Timer2_Timer()
Menu.Cls
RenderExplosion Menu
PutSprite Menu, Picture1, Picture2, 20, 0, 350, 83
End Sub

Private Sub Timer3_Timer()
NewExplosion Rnd * 350, Rnd * 350, Rnd * 255, Rnd * 255, Rnd * 255
End Sub
