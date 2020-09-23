Attribute VB_Name = "Player"

Public Type Player
    X As Single
    Y As Single
    Width As Integer
    Height As Integer
    Health As Integer
    Shield As Integer
    Hited As Boolean
    HitTime As Integer
End Type

Public Player1 As Player
Public Shooting As Boolean


Public Sub RenderPlayer()


 PutSprite Form1.Screen, Form1.Player, Form1.PlayerMask, Player1.X, Player1.Y, Player1.Width, Player1.Height
 
If Player1.Hited = True Then

     Player1.HitTime = Player1.HitTime + 1
    
     PutSprite Form1.Screen, Form1.PlayerHit, Form1.PlayerMask, Player1.X, Player1.Y, Player1.Width, Player1.Height
     
     If Player1.HitTime > 5 Then
        Player1.HitTime = 0
        Player1.Hited = False
    End If
End If

If Shooting = True Then
    
   PutSprite Form1.Screen, Form1.PlayerShooting, Form1.PlayerMask, Player1.X, Player1.Y, Player1.Width, Player1.Height
   
End If
 
 If HasMinigun = True Then 'Render MiniGun
 
    If Shooting = True Then
        Minigun.i = Minigun.i + 1
        
        If Minigun.i > Minigun.Speed Then
            Minigun.Animation = Minigun.Animation + 1
            Minigun.i = 0
        End If
        
    End If
        'Render 2 Miniguns
         PutSprite Form1.Screen, Form1.Minigun(Minigun.Animation), Form1.MinigunMask, Player1.X - 20, Player1.Y, 40, 40
         PutSprite Form1.Screen, Form1.Minigun(Minigun.Animation), Form1.MinigunMask, Player1.X + 25, Player1.Y, 40, 40
        
        If Minigun.Animation >= Minigun.Frames - 1 Then
            Minigun.Animation = 0
        End If
 End If
 
 'Render Player With Shield
 If HasShield = True Then
     PutSprite Form1.Screen, Form1.PlayerShield, Form1.PlayerMask, Player1.X, Player1.Y, Player1.Width, Player1.Height
 End If
    
 

End Sub

Public Sub InitPlayer()
Player1.Y = 400
Player1.Width = 45
Player1.Height = 45
Player1.Health = 100
Player1.Shield = 0
Form1.PHealth = "Health: " & Player1.Health
Form1.PShieldPoints = "Shield: " & Player1.Shield
End Sub
