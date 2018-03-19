Attribute VB_Name = "LetBallMove"
Option Explicit

Public Function MoveThings(Things As Object, X As Long, Y As Long) As String
Dim a As Long
  'Things.Left = X + GamePlaceWidth / 2 + 30
  'Things.Top = -Y + GamePlaceHeight / 2 + 30
  Things.Left = X
  Things.Top = Y
  MoveThings = ""
End Function

Public Function MoveBalls(Things As Object, X As Long, Y As Long) As String
  Dim Temp As String
  Dim i As Long
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
    If SliderState(2).Y >= Y And SliderState(2).Y - SliderState(2).H <= Y Or SliderState(2).NPC Then Collision (2) Else Failed 'leftout
  End If
    
  If X + BallR >= GamePlaceWidth / 2 - OutH And BallState.vX > 0 Then
    If SliderState(3).Y >= Y And SliderState(3).Y - SliderState(3).H <= Y Or SliderState(3).NPC Then Collision (3) Else Failed 'rightout
  End If
  
  If Y + BallR >= GamePlaceHeight / 2 - OutH And BallState.vY > 0 Then
    If SliderState(0).X <= X And SliderState(0).X + SliderState(0).W >= X Or SliderState(0).NPC Then Collision (0) Else Failed 'upout
  End If
  
  If Y - BallR <= -GamePlaceHeight / 2 + OutH And BallState.vY < 0 Then
    If SliderState(1).X <= X And SliderState(1).X + SliderState(1).W >= X Or SliderState(1).NPC Then Collision (1) Else Failed  'downout
  End If
 
  'If BallState.vX < 1 Then BallState.vX = (BallState.vX + 10) * 1.5
  'If BallState.vY < 1 Then BallState.vY = (BallState.vY + 10) * 1.5
  If Not (Coll) Then MoveBalls = ""
End Function

Public Function RunBall() As String
  Dim s As String
  BallState.X = BallState.X + BallState.vX * vProp
  BallState.Y = BallState.Y + BallState.vY * vProp
  s = MoveBalls(Ball, BallState.X, BallState.Y)
End Function

Public Function Collision(Sli As Long)
  Dim MoreY As Long, MoreX As Long
  CollisionTime = CollisionTime + 1
  Form_Main.Label1.Caption = CollisionTime
  Form_Main.Label9.Caption = CollisionTime
  If Sli = 0 Or Sli = 1 Then BallState.vY = -BallState.vY
  If Sli = 2 Or Sli = 3 Then BallState.vX = -BallState.vX
  
  MoreY = 8
  MoreX = 16
  Randomize
  MoreX = MoreX * Rnd
  MoreY = MoreY * Rnd
  If Sli = 0 Then BallState.vY = BallState.vY - MoreY
  If Sli = 1 Then BallState.vY = BallState.vY + MoreY
  If Sli = 2 Then BallState.vX = BallState.vX + MoreX
  If Sli = 3 Then BallState.vX = BallState.vX - MoreX
  
  Form_Main.Label2.Caption = "vX:" + Str(BallState.vX) + " vY:" + Str(BallState.vY)
End Function

Public Function Failed()
  Form_Main.Timer1.Enabled = False
  Form_Main.Label1.Caption = "Falied"
  Form_Main.Label7.Caption = "×îÖÕµÃ·Ö" + Str(CollisionTime)
  CollisionTime = 0
  Form_Main.Label9.Caption = 0
  Form_Main.StopMode.Visible = True
  Form_Main.Label9.Visible = False
  GameBegin = False
End Function

Public Sub StartBall(X, Y)
  Dim s As String
  Dim Angle As Single
  CollisionTime = 0
  BallState.X = X
  BallState.Y = Y
  s = MoveBalls(Ball, BallState.X, BallState.Y)
  Randomize
  Angle = Rnd * 360 - 180
  Do While Not (Abs(BallSreedInit * Cos(Angle)) >= 80 And Abs(BallSreedInit * Sin(Angle)) >= 60)
    Randomize
    Angle = Rnd * 360 - 180
  Loop
  Form_Main.Label3.Caption = Str(Angle)
  BallState.vX = Int(BallSreedInit * Cos(Angle))
  BallState.vY = Int(BallSreedInit * Sin(Angle))
  Form_Main.TheBall.Visible = True
  Form_Main.Label9.Visible = True
  Form_Main.Timer1.Enabled = True
  GameBegin = True
End Sub

