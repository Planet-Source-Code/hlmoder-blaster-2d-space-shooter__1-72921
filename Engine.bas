Attribute VB_Name = "Engine"
Option Explicit

Public Sub PutSprite(Control2 As Object, Picture As Object, PictureMask As Object, ByVal X As Single, ByVal Y As Single, ByVal ResX As Single, ByVal ResY As Single)
BitBlt Control2.hdc, X, Y, ResX, ResY, PictureMask.hdc, 0, 0, vbSrcAnd
BitBlt Control2.hdc, X, Y, ResX, ResY, Picture.hdc, 0, 0, vbSrcPaint
End Sub

'Square collision detection
Public Function Collision(ByVal Width1 As Single, ByVal Height1 As Single, ByVal Width2 As Single, ByVal Height2 As Single, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Integer
If ((X1 + Width1) > X2 And (Y1 + Height1) > Y2 And (X2 + Width2) > X1 And (Y2 + Height2) > Y1) Then
    Collision = 1
Else
    Collision = 0
End If
End Function

Public Function Dist(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
Dist = Sqr((X2 - X1) * (X2 - X1)) + ((Y2 - Y1) * (Y2 - Y1))
End Function


