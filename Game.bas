Attribute VB_Name = "Game"
Public EnemiesKilled As Long 'How many Emeies has the player killed
Public EarthCore As Integer  'EarthCore that player must Defend!!!

Public CurX, CurY As Integer



Public Sub GameInit()
EarthCore = 40 '40 health points to EarthCore
InitParticleEngine
InitEnemies
InitBullets
InitPowerUps
InitPlayer
End Sub


Public Sub CheckCollisions()
Dim i, k As Integer
Dim col As Integer
    
    For i = 0 To nEnemies - 1 'Loop all enemies array
        If Enemies(i).Alive = True Then 'Are they Alive?
                For k = 0 To MaxBullets - 1 'Loop all player Bullets
                    If Bullets(k).Alive = True Then 'Are they Alive?
                        col = Collision(1, 1, Enemies(i).Width, Enemies(i).Height, Bullets(k).X, Bullets(k).Y, Enemies(i).X, Enemies(i).Y)
                        If col = 1 Then
                            Enemies(i).Health = Enemies(i).Health - Bullets(k).Damage 'Enemy Loses Health
                            If Enemies(i).Type.Bullet = 1 Then
                            Else
                            Bullets(k).Alive = False
                            End If
                            Bullets(k).Vy = 0
                        End If
                    End If
                Next k
                
                'Enemy Colision With Player
                col = Collision(Player1.Width, Player1.Height, Enemies(i).Width, Enemies(i).Height, Player1.X, Player1.Y, Enemies(i).X, Enemies(i).Y)
                If col = 1 Then
                    Enemies(i).Alive = False
                    Enemies(i).Health = -1
                    If HasShield = True Then 'If the player has the shield powerup
                        Player1.Shield = Player1.Shield - 10
                    Else
                        Player1.Health = Player1.Health - 10
                    End If
                    Player1.Hited = True
                    Form1.PHealth = "Health: " & Player1.Health
                    Form1.PShieldPoints = "Shield: " & Player1.Shield
                End If
                
            End If
    Next i
    
    If DoubleFire.Alive = True Then 'Colision with DoubleFire Power Up
        col = Collision(45, 45, DoubleFire.Width, DoubleFire.Height, Player1.X, Player1.Y, DoubleFire.X, DoubleFire.Y)
        If col = 1 Then
            DoubleFire.Alive = False
            HasDoubleFire = True
            HasNoPowerUp = False
        End If
    End If
    
    
    If Minigun.Alive = True Then 'Colision with minigun Power Up
        col = Collision(45, 45, Minigun.Width, Minigun.Height, Player1.X, Player1.Y, Minigun.X, Minigun.Y)
        If col = 1 Then
            Minigun.Alive = False
            HasMinigun = True
            HasNoPowerUp = False
        End If
    End If
    
    If Shield.Alive = True Then 'Colision with Shield Power Up
        col = Collision(25, 25, Shield.Width, Shield.Height, Player1.X, Player1.Y, Shield.X, Shield.Y)
        If col = 1 Then
            Shield.Alive = False
            HasShield = True
            Player1.Shield = Player1.Shield + 50
            Form1.PShieldPoints = "Shield: " & Player1.Shield
        End If
    End If
    
    'Check Players Health
    If Player1.Health <= 0 Then
        EndGame
    End If
    
    'Check Players Shield
    If Player1.Shield <= 0 Then
        HasShield = False
    End If
 
    'If Earth is Destroyed the game is over!
    If EarthCore <= 0 Then
       EndGame
    End If
            
End Sub

Public Sub CheckEnemiesKilled()

If HasDoubleFire = False Then
    If EnemiesKilled = 16 Then 'If player has killed 16 enemies put double fire powerup
        NewDoubleFirePower Rnd * 400, Rnd * 300
        EnemiesKilled = 17
    End If
End If


If HasMinigun = False Then
    If EnemiesKilled = 70 Then 'If player has killed 71 enemies put Minigun!
        NewMinigunPower Rnd * 400, Rnd * 300
        EnemiesKilled = 71
    End If
End If

End Sub

Public Sub EndGame()
HasMinigun = False
HasDoubleFire = False
HasNoPowerUp = True
MsgBox "Game Over! " & "Your Score: " & EnemiesKilled * 10
EnemiesKilled = 0
Form1.Timer1.Enabled = False
Unload Form1
Menu.Show
ShowCursor True
End Sub

Public Sub UnloadAll()
Unload Form1
Unload Intro
Unload Menu
End Sub
