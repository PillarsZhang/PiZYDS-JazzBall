Attribute VB_Name = "OneFrameEvents"
Option Explicit

Public Sub OneFrame()
  Dim s As String
  s = RunBall
  SliderNPCEvent
  'Call PlayerEventMou(MouX, MouY)
  FPS = FPS + 1
End Sub
