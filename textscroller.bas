Attribute VB_Name = "textscroller"
Public Declare Function SendMessage Lib _
              "user32" Alias "SendMessageA" _
              (ByVal hWnd As Long, _
              ByVal wMsg As Long, _
              ByVal wParam As Long, _
              lParam As Any) As Long

Public Declare Function PutFocus Lib "user32" Alias "SetFocus" _
              (ByVal hWnd As Long) As Long
Public Declare Function SendMessageLong Lib _
              "user32" Alias "SendMessageA" _
              (ByVal hWnd As Long, _
              ByVal wMsg As Long, _
              ByVal wParam As Long, _
              ByVal lParam As Long) As Long
Declare Function LockWindowUpdate Lib "user32" _
        (ByVal hWnd As Long) As Long
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINESCROLL = &HB6

Function ScrollText(TextBox As Control, vLines As Integer) As Long

  Dim Success As Long
  Dim SavedWnd As Long
  Dim moveLines As Long

    'save the window handle of the control that currently has focus
    SavedWnd = Screen.ActiveControl.hWnd
    moveLines = vLines

    'Set the focus to the passed control (text control)
    TextBox.SetFocus

    'Scroll the lines.
    Success = SendMessage(TextBox.hWnd, EM_LINESCROLL, 0, ByVal moveLines)

    'Restore the focus to the original control
    Call PutFocus(SavedWnd)

    'Return the number of lines actually scrolled
    ScrollText = Success

End Function
