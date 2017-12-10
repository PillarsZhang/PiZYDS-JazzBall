Attribute VB_Name = "LetBallMove"
Option Explicit

Public Function MoveThings(Things As Object, X As Integer, Y As Integer) As String
  'Things.Left = X + GamePlaceWidth / 2
  'Things.Top = Y + GamePlaceHeight / 2
  Things.Left = X
  Things.Top = Y
  MoveThings = ""
End Function

Public Function MoveBalls(Things As Object, X As Integer, Y As Integer) As String
  Dim Temp As String
  Temp = MoveThings(Things, X - Things.Width / 2 + Things.Width / 10, Y + Things.Height / 2)
  MoveBalls = ""
End Function

Public Function RunBall() As String
  Dim s As String
  BallState.X = BallState.X + BallState.vX
  BallState.Y = BallState.Y + BallState.vY
  s = MoveBalls(Ball, BallState.X, BallState.Y)
End Function

