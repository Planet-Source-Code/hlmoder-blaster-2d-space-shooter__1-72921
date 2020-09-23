Attribute VB_Name = "ParticleEngine"
Option Explicit
Private Declare Function GetTickCount& Lib "kernel32" ()

Public Enum PARTICLE_STATUS
    Alive = 0
    Dead = 1
End Enum


Public Type PARTICLE
    X As Single
    Y As Single
    Vx As Single
    Vy As Single
    Color As Long
    Lifetime As Integer
    Alive As Boolean
End Type

Public Const nParticles As Integer = 180

Public Type EXPLOSION
    PosX As Single
    PosY As Single
    nParticle(nParticles) As PARTICLE
    Lifetime As Single
    Alive As Boolean
End Type




Public ShockWave() As EXPLOSION
Public Flame() As PARTICLE
Public Stars() As PARTICLE
Public Const nFlames As Integer = 40
Public Const nExplosions As Integer = 20
Public Const nStars As Integer = 20
Public Const YVariation As Integer = 5
Public Const MinusYVariation As Integer = -5
Public Const XVariation As Integer = 5
Public Const Lifetime As Integer = 600
Dim m As Single



Public Sub InitParticleEngine()
ReDim ShockWave(nExplosions - 1) As EXPLOSION


Dim i As Integer
Dim k As Integer

For i = 0 To nExplosions - 1
    ShockWave(i).PosX = 500
    ShockWave(i).PosY = 200
    ShockWave(i).Alive = False
    ShockWave(i).Lifetime = 7000

    For k = 0 To nParticles - 1
        Randomize
        ShockWave(i).nParticle(k).Color = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        ShockWave(i).nParticle(k).X = ShockWave(i).PosX
        ShockWave(i).nParticle(k).Y = ShockWave(i).PosY
            If nParticles = 180 Then
                ShockWave(i).nParticle(k).Vx = Cos((k * 2) * PI / 180) * Rnd * 20
                ShockWave(i).nParticle(k).Vy = Sin((k * 2) * PI / 180) * Rnd * 20
            ElseIf nParticles = 90 Then
                ShockWave(i).nParticle(k).Vx = Cos((k * 4) * PI / 180) * Rnd * 20
                ShockWave(i).nParticle(k).Vy = Sin((k * 4) * PI / 180) * Rnd * 20
            Else
                ShockWave(i).nParticle(k).Vx = Cos(k * PI / 180) * Rnd * 20
                ShockWave(i).nParticle(k).Vy = Sin(k * PI / 180) * Rnd * 20
            End If
    Next k
Next i

ReDim Stars(nStars - 1) As PARTICLE

For i = 0 To nStars - 1
    Stars(i).X = Rnd * 500
    Stars(i).Y = Rnd * 500
    Stars(i).Vy = 1
    Stars(i).Color = RGB(255, 255, 255)
Next i

ReDim Flame(nFlames - 1) As PARTICLE
End Sub


Public Sub RenderExplosion(Picture As Object)
Dim i As Integer
Dim k As Integer


For i = 0 To nExplosions - 1

    If ShockWave(i).Alive = True Then
        ShockWave(i).Lifetime = ShockWave(i).Lifetime - 200
        
        For k = 0 To nParticles - 1
              ShockWave(i).nParticle(k).X = (ShockWave(i).nParticle(k).X + ShockWave(i).nParticle(k).Vx)
              ShockWave(i).nParticle(k).Vx = ShockWave(i).nParticle(k).Vx
              ShockWave(i).nParticle(k).Y = (ShockWave(i).nParticle(k).Y - ShockWave(i).nParticle(k).Vy)
              ShockWave(i).nParticle(k).Vy = ShockWave(i).nParticle(k).Vy
              
              'Draw the pixels
              SetPixel Picture.hdc, ShockWave(i).nParticle(k).X, ShockWave(i).nParticle(k).Y, ShockWave(i).nParticle(k).Color
             
        Next k
        
        If ShockWave(i).Lifetime <= 0 Then
            ShockWave(i).Alive = False
            ShockWave(i).Lifetime = 7000
        End If
End If
    
Next i

End Sub

Public Sub NewExplosion(ByVal X As Single, ByVal Y As Single, ByVal r As Integer, ByVal G As Integer, ByVal B As Integer)
Dim i As Integer
Dim k As Integer

i = FreeExplosion
    If FreeExplosion <> -1 Then
        ShockWave(i).Alive = True
        ShockWave(i).PosX = X
        ShockWave(i).PosY = Y
        
            For k = 0 To nParticles - 1
                ShockWave(i).nParticle(k).Color = RGB(r, G, B)
                ShockWave(i).nParticle(k).X = ShockWave(i).PosX
                ShockWave(i).nParticle(k).Y = ShockWave(i).PosY
            Next k
                
    End If
    
End Sub


Public Function FreeExplosion() As Integer
Dim i As Integer

  For i = 0 To nExplosions - 1
        If ShockWave(i).Alive = False Then
            FreeExplosion = i
            Exit Function
        End If
    Next i
    
FreeExplosion = -1
    
End Function

'Stars Scroll
Public Sub RenderStars()
Dim i As Integer

For i = 0 To nStars - 1
    Stars(i).Y = Stars(i).Y + Stars(i).Vy
    
    If Stars(i).Y > 500 Then
        Stars(i).Y = 0
    End If
    SetPixel Form1.Screen.hdc, Stars(i).X, Stars(i).Y, RGB(255, 255, 255)
Next i
    
End Sub

'Flame
Public Sub NewFlame(ByVal X As Single, ByVal Y As Single)
Dim i As Integer
Dim k As Integer

i = FreeFlame
    If FreeFlame <> -1 Then
        Flame(i).Alive = True
        Flame(i).X = X
        Flame(i).Y = Y
    End If
    
End Sub


Public Function FreeFlame() As Integer
Dim i As Integer

  For i = 0 To nFlames - 1
        If Flame(i).Alive = False Then
            FreeFlame = i
            Exit Function
        End If
    Next i
    
FreeFlame = -1

End Function
Public Sub RenderFlame()
Dim i As Integer

    For i = 0 To nFlames - 1
        
        If Flame(i).Alive = True Then
        
            Flame(i).Lifetime = Flame(i).Lifetime + 1
            Flame(i).Y = Flame(i).Y + 1
            Flame(i).X = Flame(i).X + Cos(i * 3)
            
            If Flame(i).Lifetime = 20 Then
                Flame(i).Alive = False
                Flame(i).Lifetime = 0
            End If
            
        
            SetPixel Form1.Screen.hdc, Flame(i).X + 20, Flame(i).Y + 35, RGB(255 - (Flame(i).Lifetime * 3), Flame(i).Lifetime, i * 4)
        End If
    Next i
    
End Sub
