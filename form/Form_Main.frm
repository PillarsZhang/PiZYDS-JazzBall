VERSION 5.00
Begin VB.Form Form_Main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PiZYDS-JazzBall V1"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7875
   Icon            =   "Form_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Form_Main.frx":4C4A
   ScaleHeight     =   -5000
   ScaleLeft       =   -4000
   ScaleMode       =   0  'User
   ScaleTop        =   2500
   ScaleWidth      =   8000
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   480
      Top             =   3600
   End
   Begin VB.Frame StopMode 
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Caption         =   "Stop"
      Height          =   2535
      Left            =   1080
      TabIndex        =   8
      Top             =   960
      Width           =   5775
      Begin VB.Label Label8 
         BackColor       =   &H00FF80FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "¾ôÊ¿µ¯Çò by ÕÂÓãDS"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   960
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   3720
         MousePointer    =   2  'Cross
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF80FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Come on!!!"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF80FF&
         BackStyle       =   0  'Transparent
         Caption         =   " Jazz Ball"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   36
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         Caption         =   "   µã»÷ÒÔ¿ªÊ¼"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   2040
         Width           =   1335
      End
   End
   Begin VB.PictureBox TheBall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox Slider 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   3
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Slider 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   2
      Left            =   7680
      ScaleHeight     =   3495
      ScaleWidth      =   255
      TabIndex        =   2
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
      TabIndex        =   1
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
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp    'ÉÏ¼ýÍ·¼ü
        PlayerEventKey (0)
    Case vbKeyDown  'ÏÂ¼ýÍ·¼ü
        PlayerEventKey (1)
    Case vbKeyLeft  '×ó¼ýÍ·¼ü
        PlayerEventKey (2)
    Case vbKeyRight 'ÓÒ¼ýÍ·¼ü
        PlayerEventKey (3)
    End Select
End Sub

Private Sub Form_Load()
  Dim InitMsg As String
  Dim i As Long
  Dim t As Long
  GameBegin = False
  InitMsg = Init()
  If InitMsg <> "AllRight" Then
    MsgBox (InitMsg)
    End
  End If
  Me.KeyPreview = True
  TheBall.Picture = LoadResPicture(101, vbResBitmap)
  Image1.Picture = LoadResPicture(104, vbResBitmap)
  Me.Picture = LoadResPicture(102, vbResBitmap)
  'Me.PaintPicture Me.Picture, 0, 0, Me.Width, Me.Height
  'Me.PaintPicture Me.Picture, -Me.Width / 2, Me.Height / 2, Me.Width / 2, -Me.Height / 2
  TheBall.PaintPicture TheBall.Picture, 0, 0, TheBall.Width, TheBall.Height
  't = Rotation(Picture1.hDC, TheBall.hDC, TheBall.Width, TheBall.Height, TheBall.Width, TheBall.Height, 90)
  
  For i = 0 To 3
    Slider(i).Picture = LoadResPicture(103, vbResBitmap)
    'Slider(i).PaintPicture Slider(i).Picture, 0, 0, Slider(i).Width, Slider(i).Height
  Next i
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
  MouX = X
  MouY = Y
  If GameBegin Then Call PlayerEventMou(MouX, MouY)
  'Debug.Print (X & " " & Y)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Failed
End Sub

Private Sub Label6_Click()
  StopMode.Visible = False
  Call StartBall(0, 0)
End Sub

Private Sub Timer1_Timer()
  OneFrame
End Sub

Private Sub Image1_Click()
On Error GoTo Err
 Call ShellExecute(hwnd, "open", "http://www.pizyds.com/", vbNullString, vbNullString, &H0)
Err:
End Sub

Private Sub Timer2_Timer()
  Debug.Print (FPS)
  FPS = 0
End Sub
