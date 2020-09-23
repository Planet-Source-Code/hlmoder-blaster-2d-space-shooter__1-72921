VERSION 5.00
Begin VB.Form Intro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Blaster"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   LinkTopic       =   "BLASTER"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Intro.frx":0000
      Top             =   120
      Width           =   6615
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6600
      Top             =   1440
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   3165
      Left            =   7560
      Picture         =   "Intro.frx":003B
      ScaleHeight     =   3105
      ScaleWidth      =   7455
      TabIndex        =   1
      Top             =   6960
      Width           =   7515
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   3165
      Left            =   0
      Picture         =   "Intro.frx":4B6E9
      ScaleHeight     =   3105
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   7200
      Width           =   7515
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, Y2 As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
    Unload Me
    Load Menu
    Menu.Show
End If

End Sub

Private Sub Form_Load()
Unload Form1
Unload Menu

Y = 270
Y2 = 500
End Sub

Private Sub Timer1_Timer()
Intro.Cls
Y = Y + 1
Y2 = Y2 - 2
PutSprite Intro, Picture1, Picture2, 0, Y, 497, 207

'Stars
SetPixel Intro.hdc, 100, Y - 270, RGB(255, 255, 255)
SetPixel Intro.hdc, 200, Y - 270, RGB(255, 255, 255)
SetPixel Intro.hdc, 100, Y - 200, RGB(255, 255, 255)
SetPixel Intro.hdc, 50, Y - 150, RGB(255, 255, 255)
SetPixel Intro.hdc, 300, Y - 180, RGB(255, 255, 255)
SetPixel Intro.hdc, 350, Y - 100, RGB(255, 255, 255)
SetPixel Intro.hdc, 220, Y - 50, RGB(255, 255, 255)
SetPixel Intro.hdc, 90, Y - 20, RGB(255, 255, 255)
SetPixel Intro.hdc, 400, Y - 100, RGB(255, 255, 255)
SetPixel Intro.hdc, 450, Y - 200, RGB(255, 255, 255)

If Y > 400 Then
    Text1.Text = "DEFEND THE EARTH AT ALL COST!"
    PutSprite Intro, Form1.Player, Form1.PlayerMask, 200, Y2, 45, 45
End If

If Y > 450 Then
    Unload Me
    Load Menu
    Menu.Show
End If

End Sub
