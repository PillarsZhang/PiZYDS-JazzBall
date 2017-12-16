VERSION 5.00
Begin VB.Form Form_Main 
   AutoRedraw      =   -1  'True
   Caption         =   "PiZYDS-JazzBall V1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   -5000
   ScaleLeft       =   -4000
   ScaleMode       =   0  'User
   ScaleTop        =   2500
   ScaleWidth      =   8000
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Slider 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   3
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Slider 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   2
      Left            =   7560
      ScaleHeight     =   3495
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox Slider 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   480
      ScaleHeight     =   255
      ScaleWidth      =   6975
      TabIndex        =   2
      Top             =   4200
      Width           =   6975
   End
   Begin VB.PictureBox Slider 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   360
      ScaleHeight     =   255
      ScaleWidth      =   7335
      TabIndex        =   1
      Top             =   0
      Width           =   7335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox TheBall 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   3840
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   2280
      Width           =   250
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  BallState.vX = 220
  BallState.vY = -200
  Timer1.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp    '上箭头键
        PlayerEventKey (0)
    Case vbKeyDown  '下箭头键
        PlayerEventKey (1)
    Case vbKeyLeft  '左箭头键
        PlayerEventKey (2)
    Case vbKeyRight '右箭头键
        PlayerEventKey (3)
    End Select
End Sub

Private Sub Form_Load()
  Dim InitMsg As String
  InitMsg = Init()
  If InitMsg <> "AllRight" Then
    MsgBox (InitMsg)
    End
  End If
  Me.KeyPreview = True
End Sub

Private Sub Form_Activate()
  Dim InitMsg As String
  InitMsg = Init_2()
  If InitMsg <> "AllRight" Then
    MsgBox (InitMsg)
    End
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call PlayerEventMou(X, Y)
  Debug.Print (X & " " & Y)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  BallState.X = X
  BallState.Y = Y
  s = MoveBalls(Ball, BallState.X, BallState.Y)
  BallState.vX = 100
  BallState.vY = -200
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  OneFrame
End Sub

