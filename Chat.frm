VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form chat 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Chat coded by Navarchy"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "Chat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox NewText 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Timer weednames 
      Left            =   2400
      Top             =   1680
   End
   Begin VB.Timer myupdate 
      Left            =   1440
      Top             =   1320
   End
   Begin VB.Timer newstuff 
      Left            =   240
      Top             =   960
   End
   Begin VB.Timer backupdate 
      Left            =   120
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "send"
      Default         =   -1  'True
      Height          =   195
      Left            =   4560
      TabIndex        =   0
      Top             =   9990
      Width           =   855
   End
   Begin VB.ListBox lstNames 
      Height          =   2205
      Left            =   4080
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox chatbox 
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4048
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Chat.frx":0ECA
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   2850
      Width           =   450
   End
   Begin VB.Image imgsend 
      Height          =   300
      Index           =   0
      Left            =   4245
      Picture         =   "Chat.frx":0F78
      Top             =   2805
      Width           =   1425
   End
   Begin VB.Image imgsend 
      Height          =   300
      Index           =   1
      Left            =   4245
      Picture         =   "Chat.frx":171A
      Top             =   2805
      Width           =   1425
   End
End
Attribute VB_Name = "chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldback, oldtcolor, oldgrad, oldcol As Integer
Dim num As Integer
Dim mynumber As Integer
Dim myroom As String
Dim onornot As Boolean

Public Sub send()

    On Error Resume Next
      Open thefullpath & "data\room" & myroom & ".num" For Input As #1
      Line Input #1, newnum
      Close #1
      Open thefullpath & "data\room" & myroom For Append As #2
      Print #2, person & ": " & NewText.Text
      NewText.Text = ""
      Close #2
      If newnum > 500 Then
          newnum = newnum - 500
          Open thefullpath & "data\room" & myroom For Input As #3
          Open "room" & myroom For Output As #4
          Do While asf < 500
              Line Input #3, asdf
              asf = asf + 1
          Loop
          Do Until EOF(3)
              Line Input #1, asdf
              Print #4, asdf
          Loop
          Close #3
          Close #4
          FileCopy "room" & myroom, thefullpath & "data\room" & myroom
          Kill "room" & myroom
      End If
      Open thefullpath & "data\room" & myroom & ".num" For Output As #5
      Print #5, newnum + 1
      Close #5

End Sub

Private Sub backupdate_Timer()

    On Error Resume Next
      If sweety = True Then Exit Sub
      If oldback <> background Then
          Call backgroundupdate(Me)
      End If
      oldback = background

      If background = 0 Then
          If oldgrad <> mgradient Or oldcol <> mcolor Then Call GradientFill(Me)
          oldgrad = mgradient
          oldcol = mcolor
      End If

      If tColor <> oldtcolor Then
          If background = 2 Then Call makegray(Me, tColor)
      End If
      oldtcolor = tColor

End Sub

Private Sub Command1_Click()

    On Error Resume Next
      Call send
      imgsend(0).Visible = False
      c = Timer
      Do While Timer < c + 0.5
          DoEvents
      Loop
      imgsend(0).Visible = True

End Sub

Private Sub Form_Load()

    On Error Resume Next
      myroom = room
      If Dir$(thefullpath & "data\person" & myroom & person) = "" Then
          Open thefullpath & "data\person" & myroom & person For Output As #6
          Print #6, 0
          Close #6
      End If
      Me.ScaleMode = 3
      Me.AutoRedraw = True
      mynumber = chatnumber
      Label2.Caption = "Welcome to " & myroom
      Open thefullpath & "data\room" & myroom & ".num" For Input As #7
      Line Input #7, a
      If IsNumeric(a) = True Then num = a
      Close #7
      newname = Dir$(thefullpath & "data\person" & myroom & "*")
      Do While newname <> ""
          'MsgBox Right$(newname, Len(newname) - Len("person" & myroom))
          lstNames.AddItem Right$(newname, Len(newname) - Len("person" & myroom))
          newname = Dir$()
      Loop
      backupdate.Interval = 100
      myupdate.Interval = 50
      newstuff.Interval = 100
      weednames.Interval = 150

End Sub

Private Sub Form_LostFocus()

    Call backgroundupdate(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call boom

End Sub

Public Sub boom()

    On Error Resume Next
      chatroom(mynumber) = ""
      chatcount(mynumber) = False
      Kill "person" & myroom & person

End Sub

Private Sub imgsend_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgsend(0).Visible = False

End Sub

Private Sub imgsend_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgsend(0).Visible = True
    Call send

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgsend(0).Visible = False

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgsend(0).Visible = True
    Call send

End Sub

Private Sub myupdate_Timer()

    If closedforgood = True Then
        myupdate.Interval = 0
        weednames.Interval = 0
        newstuff.Interval = 0
        backupdate.Interval = 0
        Exit Sub
    End If
    If chatcount(mynumber) = False Then Exit Sub
    If Dir$(thefullpath & "data\kill" & myroom) <> "" Then
    Kill thefullpath & "data\person" & myroom & person
    Unload Me
    Exit Sub
    End If
    If onornot = False Then
        NewText.SetFocus
        onornot = True
    End If
    On Error GoTo makenew
    If myroom = "" Then Exit Sub
    Open thefullpath & "data\person" & myroom & person For Input As #8
    Line Input #8, oldnum
    On Error Resume Next
      Close #8
      If oldnum > 9999 Then oldnum = -9999
      Open thefullpath & "data\person" & myroom & person For Output As #9
      Print #9, oldnum + 1
      Close #9
  
Exit Sub

makenew:
      Close #8
      Open thefullpath & "data\person" & myroom & person For Output As #10
      Print #10, 0
      Close #10

End Sub

Private Sub newstuff_Timer()

    On Error Resume Next
      Open thefullpath & "data\room" & myroom & ".num" For Input As #11
      Line Input #11, newnum
      Close #11
      If num > newnum Then
          num = num - 500
      End If
      If newnum > num Then
          Open thefullpath & "data\room" & myroom For Input As #12
          Do While xt < num
              Line Input #12, a
              xt = xt + 1
          Loop
          Do While newnum > num
              Line Input #12, newline
              LockWindowUpdate Me.hWnd
              chatlen = Len(chatbox.Text)
              'ChatBox.Text = ChatBox.Text & vbCrLf & newline
              Where = InStr(newline, ":")   ' Find string in text.
              If Where Then   ' If found,
                  chatbox.SelStart = chatlen + 2   ' set selection start and
                  If Left$(newline, Len(person) + 1) = person & ":" Then
                      chatbox.SelColor = vbRed
                    Else
                      chatbox.SelColor = vbBlue
                  End If
                  'evan1 = Mid$(newline, InStr(newline, ":"), Len(newline))
                  'evan2 = Mid$(newline, InStr(newline, ":") + 1, Len(newline))
                  'evan3 = Replace$(newline, evan1, "")
                  chatbox.SelBold = True
                  chatbox.SelText = Left$(newline, Where + 1)   'evan3 & ":"
                  chatbox.SelBold = False
                  chatbox.SelColor = vbBlack
                  chatbox.SelText = Right$(newline, Len(newline) - Where - 1) & vbCrLf  'evan2 & vbCrLf
                  chatbox.SelStart = 0
                  chatbox.SelLength = 0
                Else
                  chatbox.Text = chatbox.Text & newline & vbCrLf
              End If
              Call ScrollText(chatbox, SendMessageLong(chatbox.hWnd, EM_GETLINECOUNT, 0&, 0&) - 12)
              LockWindowUpdate 0
              num = num + 1
          Loop
      End If
      Close #12

End Sub

Private Sub weednames_Timer()

    On Error Resume Next
      newname = Dir$(thefullpath & "data\person" & myroom & "*")

      Do While newname <> ""
          listnum = 0
          w = False
          Do While listnum < lstNames.ListCount
              If lstNames.List(listnum) = Right$(newname, Len(newname) - Len("person" & myroom)) Then w = True
              listnum = listnum + 1
          Loop
          If w = False Then lstNames.AddItem Right$(newname, Len(newname) - Len("person" & myroom))
          newname = Dir$()
      Loop

      listnum = 0
      Do While listnum < lstNames.ListCount
          w = False
          newname = Dir$(thefullpath & "data\person" & myroom & "*")
          Do While newname <> ""
              If Right$(newname, Len(newname) - Len("person" & myroom)) = lstNames.List(listnum) Then w = True
              newname = Dir$()
          Loop
          If w = False Then lstNames.RemoveItem (listnum)
          listnum = listnum + 1
      Loop
      Call killpeople(Me.lstNames, myroom)

End Sub
