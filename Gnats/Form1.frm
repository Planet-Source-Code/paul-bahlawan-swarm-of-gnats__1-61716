VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Gnats"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   FillColor       =   &H000000FF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   480
      Top             =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Swarm of Gnats
'Paul Bahlawan
'Dec 5, 2004
'
'update Jul 15, 2005
'update Mar 11, 2010 - Implement Atan2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Const qty As Long = 300 'Amount of gnats
Const PI As Single = 3.14159265
Const PI2 As Single = PI * 2

Private Type gnat 'Each gnat has a location, a direction, speed, and turning ability
    X As Single
    Y As Single
    dir As Single
    mag As Single
    ability As Single
End Type

Dim swarm(qty) As gnat
Dim Trails As Boolean

Private Sub Form_Load()
Dim a As Long
Randomize
    'Set the bugs up with random values
    For a = 0 To qty
        With swarm(a)
            .X = Rnd * ScaleWidth
            .Y = Rnd * ScaleHeight
            .dir = -PI * Rnd * PI2
            .mag = 3 + Rnd * 3
            .ability = 0.02 + Rnd * 0.2
        End With
    Next
    Timer1.Enabled = True
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then 'Put the female here
        swarm(0).X = X
        swarm(0).Y = Y
        swarm(0).mag = 0
    End If
    
    If Button = 2 Then 'Toggle trails
        Trails = Not Trails
    End If
End Sub


Private Sub Timer1_Timer()
Dim goal As Single
Dim cross As Single
Dim a As Long
    If Not Trails Then Cls
    For a = 0 To qty
        With swarm(a)
            If a > 0 Then 'All the male's seek the female
            
                'Find the direction (angle) to the female (goal)
                goal = Atan2(swarm(0).Y - .Y, swarm(0).X - .X)
                
                'Cross product(?) helps us find the best direction to turn
                cross = (Sin(goal) * Cos(.dir)) - (Cos(goal) * Sin(.dir))
                
                If cross < 0 Then 'Decide which way to turn
                    .dir = .dir - .ability
                Else
                    .dir = .dir + .ability
                End If
                
                If .dir < -PI Then .dir = .dir + PI2 'Keep our values between -PI and PI
                If .dir > PI Then .dir = .dir - PI2
            
            Else
            'The female (gnat # 0) usually travels in a straight line
            'but sometimes changes direction and speed
                If Int(Rnd * 100) = 1 Then
                    .dir = -PI * Rnd * PI2
                    .mag = 2 + Rnd * 5
                End If
            End If
            
            
            'Everyone is moving
            .X = .X + .mag * Cos(.dir) 'Convert Polar to Cartesian
            .Y = .Y + .mag * Sin(.dir)
                  
            If .X > ScaleWidth Then .X = .X - ScaleWidth 'Don't let em get away
            If .Y > ScaleHeight Then .Y = .Y - ScaleHeight
            If .X < 0 Then .X = .X + ScaleWidth
            If .Y < 0 Then .Y = .Y + ScaleHeight
            
            If a > 0 Then
                PSet (.X, .Y), vbWhite
            Else
                PSet (.X, .Y), vbRed
                PSet (.X + 1, .Y), vbRed
                PSet (.X - 1, .Y), vbRed
                PSet (.X, .Y + 1), vbRed
                PSet (.X, .Y - 1), vbRed
            End If
        End With
    Next
End Sub


'Atan2 - it's twice as good as Atn
'returns a value between -PI and PI
Private Function Atan2(ByVal Y As Single, ByVal X As Single) As Single
    If Y > 0 Then
        If X >= Y Then
            Atan2 = Atn(Y / X)
        ElseIf X <= -Y Then
            Atan2 = Atn(Y / X) + PI
        Else
            Atan2 = PI / 2 - Atn(X / Y)
        End If
    Else
        If X >= -Y Then
            Atan2 = Atn(Y / X)
        ElseIf X <= Y Then
            Atan2 = Atn(Y / X) - PI
        Else
            Atan2 = -Atn(X / Y) - PI / 2
        End If
    End If
End Function
 
