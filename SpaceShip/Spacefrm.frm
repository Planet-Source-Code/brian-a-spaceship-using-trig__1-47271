VERSION 5.00
Begin VB.Form Spacefrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Spacefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'good for debugging
Private Ship As SPACECRAFT 'spacecraft is my own custom object type
Private Const Radius As Integer = 200 'the radius will determine the size of my ship
Private ShotRadius As Integer 'distance of bullet from ship

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Start-Up------------------------------------------
'//////////////////////////////////////////////////
Private Sub Form_Load()
    'App.PrevInstance keeps the program from running
    'while its already running
    If App.PrevInstance = True Then End
    'the timer will control ship movement and speed
    'and the bullet movement
    Timer1.Interval = 1 'set timer to as fast as possible
    FillStyle = vbFSSolid 'makes circles solid
    'calls to the sub-procedures i made
    InitializeSpaceShip
    ShowShip
End Sub

Private Sub InitializeSpaceShip()
    With Ship
        'My spaceship is really a circle so to move it
        'i must move its center. All calculations are based
        'on my ship-circle's radius and center.
        .CenterX = Screen.Width / 2
        .CenterY = Screen.Height / 2
        .StartX = .CenterX
        .StartY = .CenterY
        .MoveRadius = 0
        'set angle to up
        .Angle = 270# 'the # sign appears when a double has .0 after the integer
        'positioning vertices with trig
        .X1 = ((Radius * Cos(Rad(.Angle + 90))) + .CenterX)
        .X2 = ((Radius * Cos(Rad(.Angle))) + .CenterX)
        .X3 = ((Radius * Cos(Rad(.Angle - 90))) + .CenterX)
        .Y1 = ((Radius * Sin(Rad(.Angle + 90))) + .CenterY)
        .Y2 = ((Radius * Sin(Rad(.Angle))) + .CenterY)
        .Y3 = ((Radius * Sin(Rad(.Angle - 90))) + .CenterY)
        ShotRadius = 0
        'set color
        .Color = RGB(0, 255, 0)
        'set speed
        .Speed = 0
    End With
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Controls------------------------------------------
'//////////////////////////////////////////////////
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then End 'escape key
    'This will allow my ship to shoot one bullet at a time
    If KeyAscii = 13 Then 'this is the enter key
        'erase any old bullets
        ShipShoot RGB(0, 0, 0)
        'make ship ready to fire
        With Ship
            .CanShoot = 1 '1 will mean yes
            .Bullet.Angle = .Angle 'set shot tragectory to same as ship when fired
            .Bullet.X = .X2 'the bullet will start at the
            .Bullet.Y = .Y2 'center of the top of the shipcircle
            .Bullet.StartX = .X2 'required for bullet to fire
            .Bullet.StartY = .Y2 'required for bullet to fire
            ShotRadius = 0
        End With
        'call to the shooting sub-procedure i made
        ShipShoot RGB(255, 0, 0) 'show the bullet
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Ship.Color = Me.BackColor 'set ship color to background
    ShowShip 'redraw ship to cheaply erase it
    If Ship.Angle < 0 Then 'Setup lower bound for .angle
        Ship.Angle = Ship.Angle + 360
    End If
    If Ship.Angle > 360 Then 'Setup upper bound for .angle
        Ship.Angle = Ship.Angle - 360
    End If
    With Ship
        If KeyCode = vbKeyLeft Then 'Turn counter-clockwise
            Ship.Angle = Ship.Angle - 6
            .StartX = .CenterX
            .StartY = .CenterY
            .MoveRadius = 0
        End If
        If KeyCode = vbKeyRight Then 'Turn clockwise
            Ship.Angle = Ship.Angle + 6
            .StartX = .CenterX
            .StartY = .CenterY
            .MoveRadius = 0
        End If
        .Color = RGB(0, 255, 0) 'reset ship color
    End With
    ShowShip 'redraw ship
End Sub

Private Sub Timer1_Timer()
    With Ship
        'the timer is going to control speed movement
        'and bullet movement
        If KeyPressed(vbKeyUp) Then
            .Speed = .Speed + 1 'speed up
        End If
        If KeyPressed(vbKeyDown) Then
            If .Speed > -2 Then .Speed = .Speed - 2 'slow
            If .Speed < 2 Then .Speed = .Speed + 2
        End If
        If .Speed < 0 Then .Speed = .Speed + 1 'stop reverse
        If .Speed > 50 Then .Speed = 50 'speed limit
        .Color = Me.BackColor
        ShowShip
        'setup boundaries
        If .CenterX + Radius > Screen.Width Or .CenterY + Radius > Screen.Height Or .CenterX - Radius < 0 Or .CenterY - Radius < 0 Then .Speed = -.Speed 'bounce off walls
        .MoveRadius = .MoveRadius + .Speed
        .CenterX = ((.MoveRadius * Cos(Rad(.Angle))) + .StartX)
        .CenterY = ((.MoveRadius * Sin(Rad(.Angle))) + .StartY)
        .Color = RGB(0, 255, 0)
        ShowShip
                
        'Bullet Stuff----------------------------------
        
        'bullet limits
        If .Bullet.X > Screen.Width Or .Bullet.Y > Screen.Width Or .Bullet.X < 0 Or .Bullet.Y < 0 Then
            'with this 'if statement' the bullet cannot travel outside of the screen
            .CanShoot = 254
            ShipShoot RGB(0, 0, 0)
        End If
        If .CanShoot = 1 Then
            With .Bullet
                ShipShoot RGB(0, 0, 0)
                'to move the bullet i will simply place
                'it on a circle according to its cos and sin
                'of its angle of travel and then increase
                'the radius of the circle, it works very well
                ShotRadius = ShotRadius + 65
                .X = ((ShotRadius * Cos(Rad(.Angle))) + .StartX)
                .Y = ((ShotRadius * Sin(Rad(.Angle))) + .StartY)
                ShipShoot RGB(255, 0, 0)
            End With
        End If
    End With
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Graphics------------------------------------------
'//////////////////////////////////////////////////
Private Sub ShowShip()
    Me.Show 'refresh screen
    'useing the 'with statement' saves on keystrokes
    With Ship 'Fun trig --V
        .X1 = ((Radius * Cos(Rad(.Angle + 90))) + .CenterX)
        .X2 = ((Radius * Cos(Rad(.Angle))) + .CenterX)
        .X3 = ((Radius * Cos(Rad(.Angle - 90))) + .CenterX)
        .Y1 = ((Radius * Sin(Rad(.Angle + 90))) + .CenterY)
        .Y2 = ((Radius * Sin(Rad(.Angle))) + .CenterY)
        .Y3 = ((Radius * Sin(Rad(.Angle - 90))) + .CenterY)
        'ship lines
        Line (.X1, .Y1)-(.X2, .Y2), .Color
        Line (.X2, .Y2)-(.X3, .Y3), .Color
        Line (.X3, .Y3)-(.X1, .Y1), .Color
    End With
End Sub

Private Sub ShipShoot(BulletColor As Long)
    With Ship
        If .CanShoot <> 1 Then
            Exit Sub 'if the ship is not able to shoot, abort current procedure
        End If
        FillColor = BulletColor
        Circle (.Bullet.X, .Bullet.Y), 15, BulletColor
    End With
End Sub
