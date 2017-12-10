Attribute VB_Name = "Initialization"
Option Explicit

Public Debugging As Boolean

Public AppZhName As String
Public AppEnName As String
Public AppName As String
Public AppVersion As String

Public GamePlaceWidth As Integer
Public GamePlaceHeight As Integer

Public Ball As Object
Public BallR As Integer
Public BallD As Integer
Public Ballprop As Long

Public Type BallStateRec
  vX As Single
  vY As Single
  X As Integer
  Y As Integer
End Type
Public BallState As BallStateRec

Public FrameTime As Integer

Public Function Init() As String

  Debugging = True
  
  AppZhName = "¾ôÊ¿µ¯Çò"
  AppEnName = "JazzBall"
  AppVersion = "V1.0.0"
  AppName = "PiZYDS-" & AppEnName & "-" & AppZhName & " " & AppVersion
  Form_Main.Caption = AppName
  
  GamePlaceWidth = 8000
  GamePlaceHeight = 5000
  
  Set Ball = Form_Main.TheBall
  BallR = 125
  BallD = BallR * 2
  Ball.Width = BallD
  Ball.Height = BallD
  'Set Ball = Form_Main.Command1
  'Ball.Visible = False
  Ballprop = 16
  MakeBallRound
  initBallState
  
  FrameTime = 60
  Form_Main.Timer1.Interval = 1000 \ FrameTime
  
  Init = "AllRight"
End Function

Public Function Init_2() As String
  Form_Main.Scale (-GamePlaceWidth / 2, GamePlaceHeight / 2)-(GamePlaceWidth / 2, -GamePlaceHeight / 2)
  
  If Debugging Then
    Form_Main.Line (-GamePlaceWidth / 2, 0)-(GamePlaceWidth / 2, 0)
    Form_Main.Line (0, -GamePlaceHeight / 2)-(0, GamePlaceHeight / 2)
  End If
  Init_2 = "AllRight"
End Function

Public Sub MakeBallRound()
  Dim crgn As Long
  Dim ret As Long
  'crgn = CreateEllipticRgn(0, 0, BallD, BallD)
  crgn = CreateEllipticRgn(0, 0, Ball.Width / Ballprop, Ball.Height / Ballprop)
  ret = SetWindowRgn(Ball.hwnd, crgn, True)
  'MsgBox (Ball.Width / 16 & " " & Ball.Height / 16)
  Ball.Visible = True
End Sub

Public Sub initBallState()
  BallState.vX = 0
  BallState.vY = 0
  BallState.X = 0
  BallState.Y = 0
End Sub
  

