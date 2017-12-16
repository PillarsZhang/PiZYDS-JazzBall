Attribute VB_Name = "SliderPlayer"
Option Explicit

Public Sub PlayerEventKey(Key As Long)
  Dim s As String
  If Key = 2 Then
    If SliderState(1).X > -GamePlaceWidth / 2 Then
      SliderState(1).X = SliderState(1).X - 100
      s = MoveThings(Sliders(1), SliderState(1).X, SliderState(1).Y)
    End If
  End If
  
  If Key = 3 Then
    If SliderState(1).X + SliderState(1).W < GamePlaceWidth / 2 Then
      SliderState(1).X = SliderState(1).X + 100
      s = MoveThings(Sliders(1), SliderState(1).X, SliderState(1).Y)
    End If
  End If
End Sub

Public Sub PlayerEventMou(X As Single, Y As Single)
  Dim s As String
  If SliderState(1).NPC = False Then
    SliderState(1).X = X - SliderState(1).W / 2
  End If
  s = MoveThings(Sliders(1), SliderState(1).X, SliderState(1).Y)
  'DOWN
  
  If SliderState(0).NPC = False Then
    SliderState(0).X = X - SliderState(0).W / 2
  End If
  s = MoveThings(Sliders(0), SliderState(0).X, SliderState(0).Y)
  'UP
  
  If SliderState(2).NPC = False Then
    SliderState(2).Y = Y + SliderState(2).H / 2
  End If
  s = MoveThings(Sliders(2), SliderState(2).X, SliderState(2).Y)
  'LEFT
  
  If SliderState(3).NPC = False Then
    SliderState(3).Y = Y + SliderState(3).H / 2
  End If
  s = MoveThings(Sliders(3), SliderState(3).X, SliderState(3).Y)
  'RIGHT

End Sub

Public Sub SliderNPCEvent()
  Dim s As String
  If SliderState(1).NPC = True Then
    SliderState(1).X = BallState.X - SliderState(1).W / 2
  End If
  s = MoveThings(Sliders(1), SliderState(1).X, SliderState(1).Y)
  'DOWN
  
  If SliderState(0).NPC = True Then
    SliderState(0).X = BallState.X - SliderState(0).W / 2
  End If
  s = MoveThings(Sliders(0), SliderState(0).X, SliderState(0).Y)
  'UP
  
  If SliderState(2).NPC = True Then
    SliderState(2).Y = BallState.Y + SliderState(2).H / 2
  End If
  s = MoveThings(Sliders(2), SliderState(2).X, SliderState(2).Y)
  'LEFT
  
  If SliderState(3).NPC = True Then
    SliderState(3).Y = BallState.Y + SliderState(3).H / 2
  End If
  s = MoveThings(Sliders(3), SliderState(3).X, SliderState(3).Y)
  'RIGHT
End Sub
