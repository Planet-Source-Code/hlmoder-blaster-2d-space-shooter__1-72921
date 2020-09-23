Attribute VB_Name = "BulletHandling"

Public Type Bullet
    X As Single
    Y As Single
    Vx As Single
    Vy As Single
    Alive As Boolean
    Damage As Integer
End Type

Public BulletSpeed As Integer
Private BulletSound As Long
Private time As Integer
Private BulletChannel As Long
Public RightCannon As Boolean
Public LeftCannon As Boolean
Public Const MaxBullets As Integer = 100
Public Bullets() As Bullet

Public Function FreeBullet() As Integer 'Function to Find a "Free" Bullet
Dim i As Integer

    For i = 0 To MaxBullets - 1
        If Bullets(i).Alive = False Then
            FreeBullet = i 'We found a free bullet
            Exit Function
        End If
    Next i

FreeBullet = -1
    
End Function

Public Sub NewBullet(ByVal X As Single, ByVal Y As Single, Damage As Integer, Frecuence As Integer)
Dim i As Integer
time = time + 1

    If time >= Frecuence Then
    
        i = FreeBullet
        
        If FreeBullet <> -1 Then
            Bullets(i).Alive = True
            Bullets(i).X = X
            Bullets(i).Y = Y
            Bullets(i).Damage = Damage
        End If
        
    time = 0
    End If
    
    
   
End Sub

Public Sub MoveBullets()
Dim i As Integer
Dim col As Integer
Dim distance As Single
    For i = 0 To MaxBullets - 1
        If Bullets(i).Alive = True Then
            Bullets(i).Y = Bullets(i).Y - BulletSpeed
           
            
            
            SetPixel Form1.Screen.hdc, Bullets(i).X, Bullets(i).Y, RGB(255, 125, 125)
            SetPixel Form1.Screen.hdc, Bullets(i).X, Bullets(i).Y + 1, RGB(200, 250, 125)
            SetPixel Form1.Screen.hdc, Bullets(i).X, Bullets(i).Y + 2, RGB(150, 125, 250)
            
                If Bullets(i).Y <= 0 Then
                    Bullets(i).Alive = False
                    Bullets(i).Vy = 0
                End If
        End If
    Next i
            
        
End Sub

Public Sub InitBullets()
Dim i As Integer
BulletSpeed = 10

ReDim Bullets(MaxBullets - 1) As Bullet
    For i = 0 To MaxBullets - 1
        Bullets(i).Alive = False
    Next i
End Sub

Public Sub Shoot()

        If HasDoubleFire = True Then 'Player has Double Fire
            NewBullet CurX + 10, CurY, 5, 5
            NewBullet CurX + 32, CurY, 5, 5
        End If
        
        If HasMinigun = True Then 'Player has MiniGun
            NewBullet CurX, CurY, 20, 5 'Left Minigun
            NewBullet CurX + 45, CurY, 10, 5 'Right Minigun
        End If
        
        If HasNoPowerUp = True Then
            NewBullet CurX + 32, CurY, 10, 10
        End If
 
End Sub

