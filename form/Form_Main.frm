VERSION 5.00
Begin VB.Form Form_Main 
   Caption         =   "PiZYDS-JazzBall V1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7875
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox TheBall 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   3840
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   2040
      Width           =   250
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

