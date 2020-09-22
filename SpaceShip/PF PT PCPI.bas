Attribute VB_Name = "PFPTPCPI"
Option Explicit 'good for debugging
Public Const PI As Double = 3.14159265358979
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer 'press-key function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Custom Type------------------------------------
'///////////////////////////////////////////////
Public Type WEAPON
    Angle As Double
    StartX As Integer
    StartY As Integer
    X As Integer
    Y As Integer
End Type

Public Type SPACECRAFT
    Angle As Double
    Bullet As WEAPON
    CanShoot As Byte
    Color As Long
    CenterX As Integer
    CenterY As Integer
    MoveRadius As Integer
    Speed As Integer
    StartX As Integer
    StartY As Integer
    'X coordinates
    X1 As Integer
    X2 As Integer
    X3 As Integer
    'Y coordinates
    Y1 As Integer
    Y2 As Integer
    Y3 As Integer
End Type
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Functions--------------------------------------
'///////////////////////////////////////////////
Public Function Deg(Radian As Double) As Double
    Deg = (Radian * 180) / PI
End Function

Public Function Rad(Degree As Double) As Double
    Rad = (Degree * PI) / 180
End Function

Public Function ArcSin(Number As Double) As Double
    'i'm still not sure if my math is correct for this one
    ArcSin = Atn(Number / Sqr(Number * -Number + 1))
End Function

Public Function KeyPressed(KeyCode As Long) As Boolean
    'returns 'true' when specified key is hit
    If GetKeyState(KeyCode) < -125 Then
        KeyPressed = True
    End If
End Function
