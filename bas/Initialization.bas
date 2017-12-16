Attribute VB_Name = "Initialization"
Option Explicit

Public Debugging As Boolean

Public AppZhName As String
Public AppEnName As String
Public AppName As String
Public AppVersion As String

Public GamePlaceWidth As Long
Public GamePlaceHeight As Long

Public Ball As Object
Public BallR As Long
Public BallD As Long
Public Ballprop As Long
Public BallSreedInit As Long

Public Type BallStateRec
  vX As Single
  vY As Single
  X As Long
  Y As Long
End Type
Public BallState As BallStateRec

Public Type SliderStateRec
  X As Long
  Y As Long
  W As Long
  H As Long
  NPC As Boolean
End Type
Public SliderState(0 To 3) As SliderStateRec
Public Sliders(0 To 3) As Object
Public OutH As Long

Public CollisionTime As Long
Public FrameTime As Long
Public vProp As Single

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
  BallSreedInit = 180
  MakeBallRound
  initBallState
  'MsgBox (Ball.Left)
  FrameTime = 30
  vProp = 30 / FrameTime
  Form_Main.Timer1.Interval = 1000 \ FrameTime
  
  initSliderState
  
  CollisionTime = 0
  Init = "AllRight"
End Function

Public Function Init_2() As String
  'Form_Main.Scale (-GamePlaceWidth / 2, GamePlaceHeight / 2)-(GamePlaceWidth / 2, -GamePlaceHeight / 2)
  
  If Debugging Then
    Form_Main.Line (-GamePlaceWidth / 2, 0)-(GamePlaceWidth / 2, 0)
    Form_Main.Line (0, -GamePlaceHeight / 2)-(0, GamePlaceHeight / 2)
    
    Form_Main.Line (-GamePlaceWidth / 2, GamePlaceHeight / 2 - 250)-(GamePlaceWidth / 2, GamePlaceHeight / 2 - 250)
    Form_Main.Line (-GamePlaceWidth / 2, -GamePlaceHeight / 2 + 250)-(GamePlaceWidth / 2, -GamePlaceHeight / 2 + 250)
    Form_Main.Line (-GamePlaceWidth / 2 + 250, GamePlaceHeight / 2)-(-GamePlaceWidth / 2 + 250, -GamePlaceHeight / 2)
    Form_Main.Line (GamePlaceWidth / 2 - 250, GamePlaceHeight / 2)-(GamePlaceWidth / 2 - 250, -GamePlaceHeight / 2)
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
  Dim s As String
  BallState.vX = 0
  BallState.vY = 0
  BallState.X = 0
  BallState.Y = 0
  s = MoveBalls(Ball, BallState.X, BallState.Y)
End Sub
  
Public Sub initSliderState()
  Dim i As Long
  Dim s As String
  Dim Thick As Long
  OutH = 250
  Thick = 300
  
  SliderState(0).H = Thick
  SliderState(0).W = 2000
  SliderState(0).X = -SliderState(0).W / 2
  SliderState(0).Y = GamePlaceHeight / 2 + SliderState(0).H - OutH
  'UP
  
  SliderState(1).H = Thick
  SliderState(1).W = 2000
  SliderState(1).X = -SliderState(1).W / 2
  SliderState(1).Y = -GamePlaceHeight / 2 + OutH
  'DOWN
  
  SliderState(2).H = 1500
  SliderState(2).W = Thick
  SliderState(2).X = -GamePlaceWidth / 2 - SliderState(2).W + OutH
  SliderState(2).Y = SliderState(2).H / 2
  'LEFT
  
  SliderState(3).H = 1500
  SliderState(3).W = Thick
  SliderState(3).X = GamePlaceWidth / 2 - OutH
  SliderState(3).Y = SliderState(3).H / 2
  'RIGHT
  
  For i = 0 To 3
    SliderState(i).NPC = True
    Set Sliders(i) = Form_Main.Slider(i)
    Sliders(i).Width = SliderState(i).W
    Sliders(i).Height = SliderState(i).H
    s = MoveThings(Sliders(i), SliderState(i).X, SliderState(i).Y)
  Next i
  SliderState(1).NPC = False
  'SliderState(0).NPC = False
End Sub

