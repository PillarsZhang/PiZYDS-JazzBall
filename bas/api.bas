Attribute VB_Name = "api"
Option Explicit

Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Type POINTAPI
  X As Long
  Y As Long
End Type

Private Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, _
   ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long
   
Public Function Rotation(ByVal DesHdc As Long, ByVal SrcHdc As Long, ByVal DesX As Long, ByVal DesY As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal Angle As Single) As Long
Dim PI As Double
PI = 3.1415926
Dim lpPoint(1 To 3) As POINTAPI
Dim Rad As Double
Rad = Angle * PI / 180
lpPoint(3).X = DesX
lpPoint(3).Y = DesX
lpPoint(1).X = lpPoint(3).X + nHeight * Sin(Rad)
lpPoint(1).Y = lpPoint(3).Y - nHeight * Cos(Rad)
lpPoint(2).X = lpPoint(1).X + nWidth * Cos(Rad)
lpPoint(2).Y = lpPoint(1).Y + nWidth * Sin(Rad)
PlgBlt DesHdc, lpPoint(1), SrcHdc, 0, 0, nWidth, nHeight, 0, 0, 0
End Function
