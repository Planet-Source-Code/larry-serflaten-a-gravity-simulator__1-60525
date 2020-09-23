VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   975
   ClientTop       =   2130
   ClientWidth     =   4305
   ClipControls    =   0   'False
   DrawMode        =   7  'Invert
   DrawStyle       =   2  'Dot
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4305
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Gravity Simulator - © 2005 Larry Serflaten [PUBLIC DOMAIN]

' This program allows the user to set balls in motion where
' the balls are attracted to eachother based on the physics of
' gravity.  Size, direction, and speed are adjustable when
' the balls are created.

' Hold any mouse button down, then drag to create a new ball.

' Groups common X, Y pairs
Private Type Vector
  X As Double
  Y As Double
End Type

' Defines a ball
Private Type ball
  Position As Vector
  Delta As Vector
  Gravity As Vector
  Weight As Single
  Radius As Single
  Color As Long
End Type

' Indicates what the user is doing
Private Enum UserState
  Pause
  Build
  Orbit
End Enum

#Const ShowDistance = False     ' Debugging aid to see distance between
#If ShowDistance Then           ' last two balls
Private DistValue As Double
#End If

Private Balls(0 To 100) As ball ' Collection of balls
Private NextFreeBall As Long    ' Next available
Private Mouse As Vector         ' Mouse Down point
Private User As Vector          ' Mouse Up point
Private TimerState As UserState ' Build / Orbit flag
Private Colors                  ' Progression based on weight
                                ' Red, Blue, White, Yellow (max)

Private BuildFactor As Double   ' Build speed (mouse button indicator)

' Constants provide easy adjustment of program parameters:

Const ZONE = 10                 ' 10      Screen border buffer (in pixels)
Const WMAX = 999999             ' 999999  Maximum weight of a ball
Const VACC = 0.999998           ' .999998 Vaccum  :  0 = solid  <...>   1 = total vaccum
Const GSCALE = 40               ' 40      Gravity : .1 = low    <...> 100 = very high
                                '         -(value) = negative gravity!




Private Sub Form_Load()
    ' Set up form
    ScaleMode = vbPixels
    AutoRedraw = True
    WindowState = vbMaximized
    BackColor = vbBlack
    ' Set colors used for different sizes
    Colors = Array(&H97CDFC, &H81ADF7, &H7384ED, &H795DD9, _
                   &H9748BD, &HBA4CA6, &HDB639B, &HF086A1, _
                   &HFBB8BE, &HFDDAD8, &HFBDCDD, &HFFFFFF)
                   
    FragileEarth
End Sub


'Mouse routines are used to set initial Weight and Direction (Delta)
'of a new ball.

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' If a button is hit, and we're not already building a ball....
    If (Button > 0) And (TimerState <> UserState.Build) Then
      ' Set build speed, save click point
      BuildFactor = 1 + (0.125 * Button)
      Mouse.X = X
      Mouse.Y = Y
      User = Mouse
      ' Start new ball, switch to build mode
      InitNewBall
      TimerState = UserState.Build
      DrawMode = vbXorPen
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If TimerState = UserState.Build Then
      ' Erase old line, track mouse, draw new line (XOR draw mode)
      Line (Mouse.X, Mouse.Y)-(User.X, User.Y), vbWhite
      User.X = X
      User.Y = Y
      Line (Mouse.X, Mouse.Y)-(User.X, User.Y), vbWhite
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If TimerState = UserState.Build Then
      ' Erase old line, fill in direction data
      Line (Mouse.X, Mouse.Y)-(User.X, User.Y), vbWhite
      With Balls(NextFreeBall)
        .Delta.X = ((Abs(GSCALE) ^ 0.8) * (Mouse.X - X)) / 100
        .Delta.Y = ((Abs(GSCALE) ^ 0.8) * (Mouse.Y - Y)) / 100
      End With
      ' Increment ball count, switch to draw mode
      NextFreeBall = NextFreeBall - (NextFreeBall < UBound(Balls))
      DrawMode = vbCopyPen
      TimerState = UserState.Orbit
    End If
End Sub


' Timer is used to provide animation
Private Sub Timer1_Timer()
    If TimerState = Build Then
      BuildBall
    Else
      Gravitation
      DrawOrbit
      ' let user see how many balls are still active
      If Caption <> Str(NextFreeBall) Then Caption = NextFreeBall
    End If
End Sub



' Gravitation calculates all the effects of gravity in the entire
' universe (the screen's representation of the universe, anyway  ;-)

Private Sub Gravitation()
Dim pull As Vector
Dim dist As Vector
Dim b1 As Long, b2 As Long
Dim wgt As Long, last As Long
Dim grv As Double, tot As Double, prop As Double

    If NextFreeBall > 1 Then
    
       last = NextFreeBall - 1
      
      ' The loops compare one ball against all the others, one time only.
      ' Each ball is compared to every other ball to see what effects
      ' gravity will have between them.  With the gravity calculated,
      ' it must be applied to the two balls such that the lighter ball
      ' gets more of the effect, while the heavier ball sees less of
      ' the effect.  The effect of all the interactions is summed up
      ' and stored (in Gravity) before being applied to any of the balls.
      
      
      For b1 = 0 To last - 1
          For b2 = b1 + 1 To last
            
            ' First the total weight involved in the interaction is
            ' calculated,
            wgt = Balls(b1).Weight + Balls(b2).Weight
            
            ' Next the distance between them is found (sign intact)
            dist.X = Balls(b2).Position.X - Balls(b1).Position.X
            dist.Y = Balls(b2).Position.Y - Balls(b1).Position.Y
            ' Distance = Sqr((x * x) + (y * y)) (tot is left squared)
            tot = (dist.X * dist.X) + (dist.Y * dist.Y)
            
            ' Weight divided by distance squared equals the total effect
            ' of gravity between the current two balls.  (Sqr and Pi are
            ' used to help scale the effects to the size of the screen)
            grv = Sqr(wgt / tot) / 3.141592653
            
            ' Now it must be divided up into its x and y componants.
            ' Sign is intact in dist, so sign is also intact for pull.
            ' GSCALE allows for scaling the gravity effect.
            pull.X = grv * (dist.X / tot) * GSCALE
            pull.Y = grv * (dist.Y / tot) * GSCALE
            
            
            ' Finally the X and Y componants are added to the cumulative
            ' gravity effects for the two balls.  If b1 is 10lbs, and b2
            ' is 5lbs then the total weight is 15lbs.  To divide the gravity
            ' effect up proportionately, b1 is given 5/15 and b2; 10/15.
            ' (b2 is the lighter ball so it would get more of the effect
            ' because it is easier to move.)
            prop = Balls(b2).Weight / wgt
            Balls(b1).Gravity.X = Balls(b1).Gravity.X + (pull.X * prop)
            Balls(b1).Gravity.Y = Balls(b1).Gravity.Y + (pull.Y * prop)
            
            ' With the sign intact for pull, it is accurate for the b1 ball.
            ' The b1 ball is being pulled toward the b2 ball, but the effect is
            ' exactly opposite when looking at the b2 ball being pulled toward
            ' the b1 ball.  To translate the effect to the b2 ball, the X and Y
            ' components must be subtracted, instead of added.
            prop = Balls(b1).Weight / wgt
            Balls(b2).Gravity.X = Balls(b2).Gravity.X - (pull.X * prop)
            Balls(b2).Gravity.Y = Balls(b2).Gravity.Y - (pull.Y * prop)
            
          Next b2
      Next b1
      
#If ShowDistance Then
  DistValue = tot
#End If
      
      pull.X = 0
      pull.Y = 0
      ' With all the interaction calculated, the summed effect is applied to
      ' the current motion of the individual balls. VACC is used to allow for
      ' adjusting the amount of friction (or drag) there appears to be.
      ' Gravity is set to 0 in preparation for the next time this routine is
      ' called.
      For b1 = 0 To last
        With Balls(b1)
          .Delta.X = VACC * .Delta.X + .Gravity.X
          .Delta.Y = VACC * .Delta.Y + .Gravity.Y
          .Gravity = pull
        End With
      Next
    End If

End Sub


' Orbit routines handle animating the balls

' DrawOrbit applies the animation and draws all the balls
Private Sub DrawOrbit()
Dim nxt As Long
    Cls   ' Comment out for neat effect
    Do While nxt < NextFreeBall
      CheckBounds nxt
      With Balls(nxt)
        .Position.X = .Position.X + .Delta.X
        .Position.Y = .Position.Y + .Delta.Y
        FillColor = .Color
        Circle (.Position.X, .Position.Y), .Radius, vbBlack
      End With
      nxt = nxt + 1
    Loop
#If ShowDistance Then
  ForeColor = vbWhite
  Print Format(DistValue, "0.00")
#End If
End Sub

' CheckBounds determines if the ball has flown off the screen
' The actual area used is the size of the screen plus a ZONE buffer
' on each side
Private Sub CheckBounds(ByVal Index As Long)
Dim swap As Boolean
Dim last As Long

    ' Check boundries
    With Balls(Index)
      swap = (.Position.X + .Radius) < -ZONE
      swap = swap Or (.Position.X - .Radius) > (ScaleWidth + ZONE)
      swap = swap Or (.Position.Y + .Radius) < -ZONE
      swap = swap Or (.Position.Y - .Radius) > (ScaleHeight + ZONE)
    End With
    
    ' If this ball is out of bounds, then the last ball in the group
    ' is moved into its slot and the last one is made available for use.
    ' That saves from having to move the remaining group to fill the
    ' hole from a removed ball.
    ' If the last ball was also out of bounds, it will be caught in
    ' the next draw cycle.
    
    If swap Then
      last = NextFreeBall - 1
      If last >= 0 Then
        Beep
        Balls(Index) = Balls(last)
        NextFreeBall = last
      End If
    End If
End Sub



' Build routines handle creating a new ball

' InitNewBall readies ball for building
Private Sub InitNewBall()
Dim v As Vector
    With Balls(NextFreeBall)
      .Color = Colors(0)
      .Weight = 30
      .Position = Mouse
      .Delta = v
      .Gravity = v
    End With
End Sub


' BuildBall increases weight until user lets go of mouse button
Private Sub BuildBall()
Dim klr As Long
Dim skip As Boolean

    With Balls(NextFreeBall)
    
      .Weight = .Weight * BuildFactor
      
      If .Weight > WMAX Then
        .Weight = WMAX
        .Color = vbYellow
      Else
        ' Reduce weight to 0-11 range for colors
        klr = (Log(.Weight) * 0.82) - 2
        If klr < 0 Then klr = 0
        If klr > 11 Then klr = 11
        ' Only draw if color is changing
        skip = (.Color = Colors(klr))
        .Color = Colors(klr)
        'Ball size is dependant on weight
        .Radius = Log(.Weight ^ 0.6) + 1
      End If
      ' Draw the ball
      If Not skip Then
        FillColor = .Color
        Circle (.Position.X, .Position.Y), .Radius, .Color
      End If
      ' Let the user see the current weight
      Caption = .Weight
    End With
  
End Sub


' FragileEarth sets up two balls orbiting a larger ball.
' It will run properly when the original constant values
' are used.  If you send in a comet and knock the Earth
' into a more oval orbit, you should know that means its
' going to get awfully cold while the earth is away from
' the sun (think Ice Age....)

Private Sub FragileEarth()

  Show
  Refresh
  
  ' Sun
  With Balls(0)
    .Color = vbYellow
    .Position.X = ScaleWidth / 2
    .Position.Y = ScaleHeight / 2
    .Delta.X = -0.001
    .Delta.Y = 0.02
    .Weight = WMAX
    .Radius = 15
  End With
  
  ' Earth
  With Balls(1)
    .Color = Colors(8)
    .Position = Balls(0).Position
    .Position.Y = .Position.Y - 80
    .Delta.X = 12.5
    .Weight = 10
    .Radius = 4
  End With
  
  ' Jupiter
  With Balls(2)
    .Color = Colors(0)
    .Position = Balls(0).Position
    .Position.X = .Position.X - 350
    .Delta.Y = -5.6
    .Weight = 3500
    .Radius = 6
  End With
  
  NextFreeBall = 3
  TimerState = UserState.Orbit
  
End Sub
