VERSION 5.00
Begin VB.Form Form_Main 
   AutoRedraw      =   -1  'True
   Caption         =   "PiZYDS-JazzBall V1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7875
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox TheBall 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   2760
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   1560
      Width           =   250
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  BallState.X = -2000
  BallState.Y = -2000
  s = MoveBalls(Ball, BallState.X, BallState.Y)
  BallState.vX = 40
  BallState.vY = 90
  Timer1.Enabled = True
End Sub

Private Sub Form_Load()
  Dim InitMsg As String
  InitMsg = Init()
  If InitMsg <> "AllRight" Then
    MsgBox (InitMsg)
    End
  End If
End Sub

Private Sub Form_Activate()
  Dim InitMsg As String
  InitMsg = Init_2()
  If InitMsg <> "AllRight" Then
    MsgBox (InitMsg)
    End
  End If
End Sub

Private Sub Timer1_Timer()
  OneFrame
End Sub
