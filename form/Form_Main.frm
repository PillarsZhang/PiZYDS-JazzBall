VERSION 5.00
Begin VB.Form Form_Main 
   Caption         =   "PiZYDS-JazzBall V1"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   8565
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
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
