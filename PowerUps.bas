Attribute VB_Name = "PowerUps"
Option Explicit

Public Type POWERUP
    X As Single
    Y As Single
    Alive As Boolean
    Lifetime As Integer
    Animation As Single
    Frames As Integer
    Speed As Integer
    Width As Integer
    Height As Integer
    Sound As Long
    SoundChannel As Long
    i As Integer
End Type



Public DoubleFire As POWERUP '2 Bullets
Public Minigun As POWERUP 'Minigun
Public Shield As POWERUP


Public HasDoubleFire As Boolean 'Has player 2 bullets?
Public HasMinigun As Boolean
Public HasShield As Boolean
Public HasNoPowerUp As Boolean



Public Sub InitPowerUps()
    DoubleFire.Alive = False
    DoubleFire.Lifetime = 500
    DoubleFire.Animation = 0
    DoubleFire.Frames = 5
    DoubleFire.Speed = 3
    DoubleFire.X = 350
    DoubleFire.Y = 350
    DoubleFire.Width = 25
    DoubleFire.Height = 25
    
    Minigun.Alive = False
    Minigun.Lifetime = 500
    Minigun.Animation = 0
    Minigun.Frames = 3
    Minigun.Speed = 3
    Minigun.X = 350
    Minigun.Y = 350
    Minigun.Width = 40
    Minigun.Height = 40
    
    
    Shield.Alive = False
    Shield.Lifetime = 200
    Shield.Animation = 0
    Shield.Frames = 5
    Shield.Speed = 3
    Shield.X = 350
    Shield.Y = 350
    Shield.Width = 25
    Shield.Height = 25

    
    HasNoPowerUp = True
    HasMinigun = False
    HasDoubleFire = False

    
End Sub


Public Sub RenderPowerUps()

'Double Fire
If DoubleFire.Alive = True Then

    DoubleFire.i = DoubleFire.i + 1
    
    If DoubleFire.i > DoubleFire.Speed Then
        DoubleFire.Animation = DoubleFire.Animation + 1
        DoubleFire.i = 0
    End If
    
    If DoubleFire.Animation >= DoubleFire.Frames - 1 Then
        DoubleFire.Animation = 0
    End If
       
       PutSprite Form1.Screen, Form1.PDoubleFire(DoubleFire.Animation), Form1.PDoubleFireMask, DoubleFire.X, DoubleFire.Y, 25, 25
       DoubleFire.Lifetime = DoubleFire.Lifetime - 1
       
    If DoubleFire.Lifetime <= 0 Then
        DoubleFire.Lifetime = 500
        DoubleFire.Alive = False
    End If
    
End If

'Minigun
If Minigun.Alive = True Then

       PutSprite Form1.Screen, Form1.Minigun(0), Form1.MinigunMask, Minigun.X, Minigun.Y, 40, 40
       Minigun.Lifetime = Minigun.Lifetime - 1
       
    If Minigun.Lifetime <= 0 Then
        Minigun.Lifetime = 500
        Minigun.Alive = False
    End If
    
End If

'Shield
If Shield.Alive = True Then

    Shield.i = Shield.i + 1
    
    If Shield.i > Shield.Speed Then
        Shield.Animation = Shield.Animation + 1
        Shield.i = 0
    End If
    
    If Shield.Animation >= Shield.Frames - 1 Then
        Shield.Animation = 0
    End If
       
       PutSprite Form1.Screen, Form1.PShield(Shield.Animation), Form1.PDoubleFireMask, Shield.X, Shield.Y, 25, 25
       Shield.Lifetime = Shield.Lifetime - 1
       
    If Shield.Lifetime <= 0 Then
        Shield.Lifetime = 200
        Shield.Alive = False
    End If
    
End If

End Sub

Public Sub NewDoubleFirePower(ByVal X As Single, ByVal Y As Single)
     DoubleFire.Alive = True
     DoubleFire.X = X
     DoubleFire.Y = Y
End Sub


Public Sub NewMinigunPower(ByVal X As Single, ByVal Y As Single)
     Minigun.Alive = True
     Minigun.X = X
     Minigun.Y = Y
End Sub


Public Sub NewShieldPower(ByVal X As Single, ByVal Y As Single)
     Shield.Alive = True
     Shield.X = X
     Shield.Y = Y
End Sub

        
            
            
       
        
            
        


