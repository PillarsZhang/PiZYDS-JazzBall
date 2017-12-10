Attribute VB_Name = "Initialization"
Option Explicit

Public AppZhName As String
Public AppEnName As String
Public AppName As String
Public AppVersion As String

Public Function Init() As String
  
  AppZhName = "¾ôÊ¿µ¯Çò"
  AppEnName = "JazzBall"
  AppVersion = "V1.0.0"
  
  AppName = "PiZYDS-" & AppEnName & "-" & AppZhName & " " & AppVersion
  
  Form_Main.Caption = AppName
  Init = "AllRight"
End Function
