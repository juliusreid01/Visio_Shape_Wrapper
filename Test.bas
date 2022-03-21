Attribute VB_Name = "Test"

Public Sub Tester()
  Dim Line as Shape
  Dim Rect as Shape
  Dim wrapLine As vw_shape_wrapper_c
  Dim wrapRect As vw_shape_wrapper_c

  Set Line = ActivePage.DrawLine(1, 10, 4, 10)
  Set Rect = ActivePage.DrawRectangle(1, 9, 4, 9.5)

  Set wrapLine = New vw_shape_wrapper_c
  Set wrapRect = New vw_shape_wrapper_c
  Set wrapLine.vsoShape = Line
  Set wrapRect.vsoShape = Rect

  wrapLine.Width = 2
  wrapRect.Height = 0.25

End Sub