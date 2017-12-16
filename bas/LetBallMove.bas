Attribute VB_Name = "LetBallMove"
Option Explicit

Public Function MoveThings(Things As Object, X As Integer, Y As Integer) As String
Dim a As Long
  'Things.Left = X + GamePlaceWidth / 2 + 30
  'Things.Top = -Y + GamePlaceHeight / 2 + 30
  Things.Left = X
  Things.Top = Y
  MoveThings = ""
End Function

Public Function MoveBalls(Things As Object, X As Integer, Y As Integer) As String
  Dim Temp As String
  Dim i As Integer
  Dim Coll As Boolean
  Temp = MoveThings(Things, X - Things.Width / 2 + 30, Y + Things.Height / 2)
  'Temp = MoveThings(Things, 0, 0)
  'MsgBox (Things.Top)
  i = 0
  Coll = False
'  Do While i <= 3 And Not (Coll)
'    If X >= SliderState(i).X - BallR And SliderState(i).X + SliderState(i).W >= X Then
'      If SliderState(i).Y + BallR >= Y And Y >= SliderState(i).Y - SliderState(i).H Then
'        Collision (i)
'        MoveBalls = "c," + Str(i)
'        Coll = True
'      End If
'    End If
'    i = i + 1
'  Loop
  
  If X - BallR <= -GamePlaceWidth / 2 + OutH And BallState.vX < 0 Then
    If SliderState(2).Y >= Y And SliderState(2).Y - SliderState(2).H <= Y Then Collision (2) Else Failed 'leftout
  End If
    
  If X + BallR >= GamePlaceWidth / 2 - OutH And BallState.vX > 0 Then
    If SliderState(3).Y >= Y And SliderState(3).Y - SliderState(3).H <= Y Then Collision (3) Else Failed 'rightout
  End If
  
  If Y + BallR >= GamePlaceHeight / 2 - OutH And BallState.vY > 0 Then
    If SliderState(0).X <= X And SliderState(0).X + SliderState(0).W >= X Then Collision (0) Else Failed 'upout
  End If
  
  If Y - BallR <= -GamePlaceHeight / 2 + OutH And BallState.vY < 0 Then
    If SliderState(1).X <= X And SliderState(1).X + SliderState(1).W >= X Then Collision (1) Else Failed  'downout
  End If
 
  'If BallState.vX < 1 Then BallState.vX = (BallState.vX + 10) * 1.5
  'If BallState.vY < 1 Then BallState.vY = (BallState.vY + 10) * 1.5
  If Not (Coll) Then MoveBalls = ""
End Function

Public Function RunBall() As String
  Dim s As String
  BallState.X = BallState.X + BallState.vX
  BallState.Y = BallState.Y + BallState.vY
  s = MoveBalls(Ball, BallState.X, BallState.Y)
End Function

Public Function Collision(Sli As Integer)
  CollisionTime = CollisionTime + 1
  Form_Main.Label1.Caption = CollisionTime
  If Sli = 0 Or Sli = 1 Then BallState.vY = -BallState.vY
  If Sli = 2 Or Sli = 3 Then BallState.vX = -BallState.vX
  
  If Sli = 0 Then BallState.vY = BallState.vY * 0.8
  If Sli = 1 Then BallState.vY = BallState.vY * 1
  If Sli = 2 Then BallState.vX = BallState.vX * 1.1
  If Sli = 3 Then BallState.vX = BallState.vX * 1.2
End Function

Public Function Failed()
  Form_Main.Timer1.Enabled = False
  MsgBox ("Failed!")
End Function

