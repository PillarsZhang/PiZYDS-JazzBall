Attribute VB_Name = "SliderPlayer"
Option Explicit

Public Sub PlayerEventKey(Key As Integer)
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
  SliderState(1).X = X - SliderState(1).W / 2
  s = MoveThings(Sliders(1), SliderState(1).X, SliderState(1).Y)
End Sub
