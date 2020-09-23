Attribute VB_Name = "backgrounds"
Global tColor As Long

Public Sub makegray(who As Form, tColor As Long)

    who.Picture = mainchat.Image1.Picture
    who.Cls
    If tColor < 1 Then Exit Sub
    who.BackColor = tColor
    'who.DrawMode = 9
    'who.ForeColor = tColor
    'who.Line (0, 0)-(iWidth, iHeight), , BF

End Sub

Public Sub GradientFill(who As Form)

  'This code is simply for a cool background effect.
  'I did not write this code. For any further info
  'or to post a question about it go to
  'http://www.planet-source-code.com/xq/ASP/txtCodeId.13788/lngWId.1/qx/vb/scripts/ShowCode.htm

    who.Refresh
    who.DrawMode = 13
  Dim i As Long
  Dim c As Integer
  Dim r As Double
    r = who.ScaleHeight / mgradient '3.142
    'Hint: Multiplying r by differnt values
    '     give different effects (try 2.3)

    For i = 0 To who.ScaleHeight
        c = Abs(220 * Sin(i / r))
        'Hint: Changing sin to cos reverses rang
        '     e
        who.Line (0, i)-(who.ScaleWidth, i), RGB(c, c, c + mcolor) '30)
        'Hint: Notice the bias to blue. You can
        '     be more subtle by reducing this number (
        '     try 10). Try other colours too.
    Next

End Sub

Public Sub backgroundupdate(who As Form)

    If background = 0 Then Call GradientFill(who)
    If background = 2 Then Call makegray(who, tColor)

End Sub

