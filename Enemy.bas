Attribute VB_Name = "Enemy"
Option Explicit

Public Type EnemyType
    Blaster As Integer
    Charger As Integer
    Bullet As Integer
    Bullet1 As Integer
    Speeder As Integer
    Disc As Integer
End Type

Public Type Enemy
    X As Single
    Y As Single
    Vx As Single
    Vy As Single
    Width As Single
    Height As Single
    SpeedX As Single
    SpeedY As Single
    Alive As Boolean
    Health As Integer
    ShootTime As Integer
    i As Integer
    Animation As Integer
    Type As EnemyType
End Type


Public Enemies() As Enemy
Public Squad(10) As String
Public Const nEnemies As Integer = 20

Public Sub InitEnemies()
Dim i As Integer
ReDim Enemies(nEnemies - 1) As Enemy


For i = 0 To nEnemies - 1
    Enemies(i).Alive = False
Next i



End Sub


Public Sub MoveEnemies()
Dim i As Integer
Dim distance As Single

For i = 0 To nEnemies - 1
    If Enemies(i).Alive = True Then
    
        If Enemies(i).Type.Charger = 1 Then '***********Enemy Charger IA***********
        
            Enemies(i).Y = Enemies(i).Y + 1
            
             Enemies(i).ShootTime = Enemies(i).ShootTime + 1
             
             If Enemies(i).ShootTime >= 60 Then 'Shoot A Bullet
                NewEnemy Enemies(i).X + 15, Enemies(i).Y, "Bullet"
                Enemies(i).ShootTime = 0
             End If
             
             If Enemies(i).Y > 450 Then 'When Enemy has reached Y limit
                EarthCore = EarthCore - 1 'EarthCore takes damage!
                Enemies(i).Alive = False
            End If
            
            If Enemies(i).Health <= 0 Then 'Dead Enemy
                NewExplosion Enemies(i).X, Enemies(i).Y, Rnd * 255, Rnd * 255, Rnd * 255
                EnemiesKilled = EnemiesKilled + 1 'One Enemy Killed!
                Enemies(i).Alive = False
            End If
            
        ElseIf Enemies(i).Type.Blaster = 1 Then '***********Enemy Blaster IA***********
        
            Enemies(i).Y = Enemies(i).Y + 2 ' Move Enemy Down
            Enemies(i).X = Enemies(i).X + Enemies(i).Vx
            
             Enemies(i).ShootTime = Enemies(i).ShootTime + 1
             
             If Enemies(i).ShootTime >= 60 Then 'Shoot A Bullet
                NewEnemy Enemies(i).X + 8, Enemies(i).Y, "Bullet1"
                NewEnemy Enemies(i).X + 32, Enemies(i).Y, "Bullet1"
                Enemies(i).ShootTime = 0
             End If
             
            If Enemies(i).Y > Player1.Y - 50 Then 'When Enemy has passed player
            
                'Calculate the path to colide with player
                If Enemies(i).X > Player1.X Then
                    Enemies(i).Vx = Enemies(i).Vx - 0.3
                Else
                    Enemies(i).Vx = Enemies(i).Vx + 0.3
                End If
                
            End If
             
            If Enemies(i).Y > 450 Then 'When Enemy has reached Y limit
                EarthCore = EarthCore - 1 'EarthCore takes damage!
                Enemies(i).Alive = False
                Enemies(i).Vx = 0 'Reset Vx Speed
            End If
            
             If Enemies(i).Health <= 0 Then 'Dead Enemy
                NewExplosion Enemies(i).X, Enemies(i).Y, Rnd * 255, Rnd * 255, Rnd * 255
                EnemiesKilled = EnemiesKilled + 1 'One Enemy Killed!
                Enemies(i).Alive = False
            End If
            
            
        ElseIf Enemies(i).Type.Speeder = 1 Then '*********Enemy Speeder IA***********
        
            Enemies(i).Y = Enemies(i).Y + 3 ' Move Enemy Down
            
            If Enemies(i).Y > 450 Then 'When Enemy has reached Y limit
                 EarthCore = EarthCore - 1 'EarthCore takes damage!
                Enemies(i).Alive = False
            End If
            
             If Enemies(i).Health <= 0 Then 'Dead Enemy
                NewExplosion Enemies(i).X, Enemies(i).Y, Rnd * 255, Rnd * 255, Rnd * 255
                EnemiesKilled = EnemiesKilled + 1 'One Enemy Killed!
                Enemies(i).Alive = False
            End If
            
        ElseIf Enemies(i).Type.Disc = 1 Then '***********Enemy Disc IA***********
        
            Enemies(i).Y = Enemies(i).Y + 1.5 ' Move Enemy Down
            Enemies(i).X = Enemies(i).X + Cos(Enemies(i).Y * PI / 180)
            
            If Enemies(i).Y > 450 Then 'When Enemy has reached Y limit
                EarthCore = EarthCore - 1 'EarthCore takes damage!
                Enemies(i).Alive = False
            End If
            
             If Enemies(i).Health <= 0 Then 'Dead Enemy
                NewExplosion Enemies(i).X, Enemies(i).Y, Rnd * 255, Rnd * 255, Rnd * 255
                EnemiesKilled = EnemiesKilled + 1 'One Enemy Killed!
                Enemies(i).Alive = False
            End If
            
            
        ElseIf Enemies(i).Type.Bullet = 1 Then '***********Enemy Bullet***********
        
            Enemies(i).Y = Enemies(i).Y + 5 ' Move Enemy Down
            
            If Enemies(i).Y > 450 Then 'When Enemy has reached Y limit
                Enemies(i).Alive = False
            End If
            
            
      ElseIf Enemies(i).Type.Bullet1 = 1 Then '***********Enemy Bullet1***********
        
            Enemies(i).Y = Enemies(i).Y + 8 ' Move Enemy Down
            
            If Enemies(i).Y > 450 Then 'When Enemy has reached Y limit
                Enemies(i).Alive = False
            End If
            
            
        End If
        
    End If ' End IF Alive
Next i

End Sub

Public Sub RenderEnemies()
Dim i As Integer

For i = 0 To nEnemies - 1
    If Enemies(i).Alive = True Then
    
        If Enemies(i).Type.Blaster = 1 Then
            PutSprite Form1.Screen, Form1.Enemy5, Form1.Enemy5Mask, Enemies(i).X, Enemies(i).Y, 45, 45
            
        ElseIf Enemies(i).Type.Charger = 1 Then
            PutSprite Form1.Screen, Form1.Enemy2, Form1.PlayerMask, Enemies(i).X, Enemies(i).Y, 40, 40
            
        ElseIf Enemies(i).Type.Speeder = 1 Then
            PutSprite Form1.Screen, Form1.Enemy3, Form1.Enemy3mask, Enemies(i).X, Enemies(i).Y, 40, 40
            
        ElseIf Enemies(i).Type.Disc = 1 Then
        
            '6 Frames of animation
            Enemies(i).Animation = Enemies(i).Animation + 1
            
            If Enemies(i).Animation >= 5 Then
                Enemies(i).Animation = 0
                Enemies(i).i = Enemies(i).i + 1
            End If
            
            If Enemies(i).i >= 6 Then
                Enemies(i).i = 0
            End If
            
                PutSprite Form1.Screen, Form1.Enemy(Enemies(i).i), Form1.EnemyMask, Enemies(i).X, Enemies(i).Y, 35, 35
            
        ElseIf Enemies(i).Type.Bullet = 1 Then
            PutSprite Form1.Screen, Form1.Bullet, Form1.BulletMask, Enemies(i).X, Enemies(i).Y, 17, 17
            
         ElseIf Enemies(i).Type.Bullet1 = 1 Then
            SetPixel Form1.Screen.hdc, Enemies(i).X, Enemies(i).Y, RGB(0, 255, 255)
            
        End If
        
    End If
Next i

End Sub

Public Function FreeEnemy() As Integer 'Function to Find a "Free" Enemy
Dim i As Integer

    For i = 0 To nEnemies - 1
        If Enemies(i).Alive = False Then
            FreeEnemy = i
            Exit Function
        End If
    Next i

FreeEnemy = -1
    
End Function

Public Sub NewEnemy(ByVal X As Single, ByVal Y As Single, Caracter As String)
Dim i As Integer

i = FreeEnemy
    If FreeEnemy <> -1 Then
        Enemies(i).Alive = True
        Enemies(i).X = X
        Enemies(i).Y = Y
        
            If Caracter = "Blaster" Then
                Enemies(i).Type.Blaster = 1
                Enemies(i).Type.Charger = 0
                Enemies(i).Type.Speeder = 0
                Enemies(i).Type.Disc = 0
                Enemies(i).Health = 10
            ElseIf Caracter = "Charger" Then
                 Enemies(i).Type.Blaster = 0
                 Enemies(i).Type.Charger = 1
                 Enemies(i).Type.Speeder = 0
                Enemies(i).Type.Disc = 0
                 Enemies(i).Health = 80
                 Enemies(i).Width = 40
                 Enemies(i).Height = 40
            ElseIf Caracter = "Speeder" Then
                 Enemies(i).Type.Speeder = 1
                 Enemies(i).Type.Blaster = 0
                 Enemies(i).Type.Charger = 0
                 Enemies(i).Health = 10
                 Enemies(i).Width = 25
                 Enemies(i).Height = 23
                 Enemies(i).Width = 40
                 Enemies(i).Height = 40
            ElseIf Caracter = "Disc" Then
                 Enemies(i).Type.Speeder = 0
                 Enemies(i).Type.Blaster = 0
                 Enemies(i).Type.Charger = 0
                 Enemies(i).Type.Disc = 1
                 Enemies(i).Health = 60
                 Enemies(i).Width = 35
                 Enemies(i).Height = 35
            ElseIf Caracter = "Bullet" Then
                 Enemies(i).Type.Speeder = 0
                 Enemies(i).Type.Blaster = 0
                 Enemies(i).Type.Charger = 0
                 Enemies(i).Type.Disc = 0
                 Enemies(i).Type.Bullet = 1
                 Enemies(i).Health = 500
                 Enemies(i).Width = 1
                 Enemies(i).Height = 1
             ElseIf Caracter = "Bullet1" Then
                 Enemies(i).Type.Speeder = 0
                 Enemies(i).Type.Blaster = 0
                 Enemies(i).Type.Charger = 0
                 Enemies(i).Type.Disc = 0
                 Enemies(i).Type.Bullet1 = 1
                 Enemies(i).Type.Bullet = 0
                 Enemies(i).Health = 500
                 Enemies(i).Width = 1
                 Enemies(i).Height = 1
            End If
                
    End If
    
   
End Sub



Public Sub InitSquadPositions(Formation As Integer)

'S=Speeders
'C=Chargers
'D=Discs
'B=Blasters

Select Case Formation
    Case 1
        Squad(1) = "     "
        Squad(2) = "S   S"
        Squad(3) = " S S "
        Squad(4) = "  S  "
    Case 2
        Squad(1) = "SCCCS"
        Squad(2) = "S   S"
        Squad(3) = " S S "
        Squad(4) = "  S  "
     Case 3
        Squad(1) = "S S S"
        Squad(2) = "S C S"
        Squad(3) = " S S "
        Squad(4) = "C S C"
     Case 4
        Squad(1) = "S   S"
        Squad(2) = "S   S"
        Squad(3) = " S S "
        Squad(4) = "  S  "
        Squad(5) = "  S  "
        Squad(6) = "S   S"
        Squad(7) = "S   S"
        Squad(8) = "S   S"
     Case 5
        Squad(1) = "  D  "
        Squad(2) = "D   D"
        Squad(3) = "D   D"
        Squad(4) = "     "
        
     Case 6
        Squad(1) = "  D  "
        Squad(2) = "D   D"
        Squad(3) = "B   B"
        Squad(4) = "S   S"
End Select
    

End Sub

Public Sub NewSquad(Formation As Integer, PosX As Single, PosY As Single)
Dim X As Integer
Dim Y As Integer

InitSquadPositions Formation

    For X = 1 To 9
        For Y = 1 To 9
           If VBA.Mid$(Squad(Y), X, 1) = "S" Then
                Call NewEnemy(X * 40 + PosX, Y * 40 + PosY, "Speeder")
            ElseIf VBA.Mid$(Squad(Y), X, 1) = "C" Then
                Call NewEnemy(X * 40 + PosX, Y * 40 + PosY, "Charger")
            ElseIf VBA.Mid$(Squad(Y), X, 1) = "D" Then
                Call NewEnemy(X * 40 + PosX, Y * 40 + PosY, "Disc")
             ElseIf VBA.Mid$(Squad(Y), X, 1) = "B" Then
                Call NewEnemy(X * 40 + PosX, Y * 40 + PosY, "Blaster")
            End If
        Next Y
    Next X
                
              
End Sub
