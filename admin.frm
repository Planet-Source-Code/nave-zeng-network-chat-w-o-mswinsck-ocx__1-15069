VERSION 5.00
Begin VB.Form admin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator Options coded by Navarchy"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "admin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer onoroff 
      Interval        =   200
      Left            =   2880
      Top             =   360
   End
   Begin VB.Timer newroom 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer newpeople 
      Interval        =   150
      Left            =   960
      Top             =   360
   End
   Begin VB.ListBox List3 
      Height          =   2205
      Left            =   4200
      MultiSelect     =   2  'Extended
      TabIndex        =   15
      ToolTipText     =   """selected people"" refers to the selected people in this list"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   2160
      TabIndex        =   14
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   13
      ToolTipText     =   """selected rooms"" refers to the selected rooms in this list"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "boot selected people"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton Command6 
      Caption         =   "delete selected rooms"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "create new room"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "broadcast to selected rooms"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PM selected people"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1920
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "restore all chat programs"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "close all chat programs"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "People"
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      ToolTipText     =   """selected people"" refers to the selected people in this list"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "People in selected rooms"
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rooms"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   """selected rooms"" refers to the selected rooms in this list"
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   6015
   End
   Begin VB.Line Line6 
      X1              =   6240
      X2              =   0
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line7 
      X1              =   3120
      X2              =   6240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line3 
      X1              =   3120
      X2              =   6240
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line5 
      X1              =   3120
      X2              =   6240
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line4 
      X1              =   3120
      X2              =   3120
      Y1              =   0
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3120
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Chat Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Open thefullpath & "data\shutdownnet" For Output As #1
    Print #1, ""
    Close #1

End Sub

Private Sub Command2_Click()

    On Error Resume Next
      Kill thefullpath & "data\shutdownnet"

End Sub

Private Sub Command3_Click()
    If Left$(Text5.Text, 4) = "boot" Then MsgBox "Invalid Room Name", vbalert, "": Exit Sub
    If Left$(Text5.Text, 4) = "room" Then MsgBox "Invalid Room Name", vbalert, "": Exit Sub
    If Left$(Text5.Text, 6) = "person" Then MsgBox "Invalid Room Name", vbalert, "": Exit Sub
    If Right$(Text5.Text, 4) = ".num" Then MsgBox "Invalid Room Name", vbalert, "": Exit Sub
    If Text5.Text <> "" Then
        Open thefullpath & "data\room" & Text5.Text For Output As #16
        Print #16, ""
        Close #16
        Open thefullpath & "data\room" & Text5.Text & ".num" For Output As #17
        Print #17, 1
        Close #17
    End If

End Sub

Private Sub Command4_Click()

    Do While a < List2.ListCount
        If List3.Selected(a) = True Then Shell "net send " & List3.List(a) & " " & Text2.Text, vbHide
        a = a + 1
    Loop

End Sub

Private Sub Command5_Click()

    On Error Resume Next
      Do While a < List1.ListCount
          If List1.Selected(a) = True Then
              Open thefullpath & "data\room" & List1.List(a) & ".num" For Input As #1
              Line Input #1, newnum
              Close #1
              Open thefullpath & "data\room" & List1.List(a) For Append As #2
              Print #2, Text3.Text & Text4.Text
              Close #2
              Open thefullpath & "data\room" & List1.List(a) & ".num" For Output As #3
              Print #3, newnum + 1
              Close #3
          End If
          a = a + 1
      Loop

End Sub

Private Sub Command6_Click()

    'On Error Resume Next
      Do While a < List1.ListCount
          If List1.Selected(a) = True Then
              Open thefullpath & "data\kill" & List1.List(a) For Output As #1
              Print #1, ""
              Close #1
              Do While Dir$(thefullpath & "data\person" & List1.List(a) & "*") <> ""
                  DoEvents
              Loop
              Kill thefullpath & "data\kill" & List1.List(a)
              Kill thefullpath & "data\room" & List1.List(a)
              Kill thefullpath & "data\room" & List1.List(a) & ".num"
              Label1.Caption = ""
              List2.Clear
          End If
          a = a + 1
      Loop
End Sub

Private Sub Command7_Click()

    Do While a < List3.ListCount
        If List3.Selected(a) = True Then
            Open thefullpath & "data\boot" & List3.List(a) For Output As #1
            Print #1, ""
            Close #1
        End If
        a = a + 1
    Loop

End Sub

Private Sub Form_Load()

    adminon = True
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Text1.Text = person
    Text3.Text = person & ": "
    Call backgroundupdate(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    adminon = False

End Sub

Private Sub newpeople_Timer()

    Do While a < List1.ListCount
        If List1.Selected(a) = True Then Call newpeoplemaker(List1.List(a), Me.List2, True)
        Call newpeoplemaker(List1.List(a), Me.List3, False)
        a = a + 1
    Loop
    Call removeoldpeople

End Sub

Public Sub removeoldpeople()

    listnum = 0
    Do While listnum < List3.ListCount
        w = False
        a = 0
        Do While a < List1.ListCount
            newname = Dir$(thefullpath & "data\person" & List1.List(a) & "*")
            Do While newname <> ""
                If Right$(newname, Len(newname) - Len("person" & List1.List(a))) = List3.List(listnum) Then w = True
                newname = Dir$()
            Loop
            a = a + 1
        Loop
        If w = False Then List3.RemoveItem (listnum)
        listnum = listnum + 1
    Loop

End Sub

Public Sub newpeoplemaker(list1text As String, list2double As ListBox, crap As Boolean)

    On Error Resume Next

      If list1text = "" Then Exit Sub

      newname = Dir$(thefullpath & "data\person" & list1text & "*")

      Do While newname <> ""
          listnum = 0
          w = False
          Do While listnum < list2double.ListCount
              If list2double.List(listnum) = Right$(newname, Len(newname) - Len("person" & list1text)) Then w = True
              listnum = listnum + 1
          Loop
          If w = False Then list2double.AddItem Right$(newname, Len(newname) - Len("person" & list1text))
          newname = Dir$()
      Loop
      If crap = True Then
          listnum = 0
          Do While listnum < list2double.ListCount
              w = False
              a = 0
              Do While a < List1.ListCount
                  If List1.Selected(a) = True Then
                      newname = Dir$(thefullpath & "data\person" & List1.List(a) & "*")
                      Do While newname <> ""
                          If Right$(newname, Len(newname) - Len("person" & List1.List(a))) = list2double.List(listnum) Then w = True
                          newname = Dir$()
                      Loop
                  End If
                  a = a + 1
              Loop
              If w = False Then list2double.RemoveItem (listnum)
              listnum = listnum + 1
          Loop
          final = list2double.ListCount & " users in "
          Do While s < List1.ListCount
              If final <> list2double.ListCount & " users in " Then g = ", "
              If List1.Selected(s) = True Then final = final & g & List1.List(s)
              s = s + 1
          Loop
          If Label1.Caption <> final Then Label1.Caption = final
      End If
      Call killpeople(list2double, list1text)

End Sub

Private Sub newroom_Timer()

    On Error Resume Next
      If List1.Text = "" Then
          If Label1.Caption <> "0 users in no room selected" Then Label1.Caption = "0 users in no room selected"
          If List2.ListCount <> 0 Then List2.Clear
          If List3.ListCount <> 0 Then List3.Clear
      End If
      a = Dir$(thefullpath & "data\room*")
      Do While a <> ""
          a = Right$(a, Len(a) - 4)
          listnum = 0
          w = False
          Do While listnum < List1.ListCount
              If List1.List(listnum) = a Then w = True
              listnum = listnum + 1
          Loop
          If w = False Then
              If Right$(a, 4) <> ".num" Then List1.AddItem a
          End If
          a = Dir$()
      Loop
      listnum = 0
      Do While listnum < List1.ListCount
          If Dir$(thefullpath & "data\room" & List1.List(listnum)) = "" Then List1.RemoveItem (listnum)
          listnum = listnum + 1
      Loop

End Sub

Private Sub onoroff_Timer()

    If Dir$(thefullpath & "data\shutdownnet") <> "" Then
        If Command1.Enabled = True Then
            Command1.Enabled = False
            Command2.Enabled = True
        End If
      Else
        If Command2.Enabled = True Then
            Command2.Enabled = False
            Command1.Enabled = True
        End If
    End If

End Sub

Private Sub Text1_Change()
If Text1.Text = person Then Exit Sub
    Do While dfg < 10
        If chatcount(dfg) = True Then w = True
        dfg = dfg + 1
    Loop
    If w = False Then
        a = Dir$(thefullpath & "data\person*" & person)
        Do While a <> ""
            Kill thefullpath & "data\" & a
            a = Dir$()
        Loop
        If Text3.Text = person & ": " Then Text3.Text = Text1.Text & ": "
        person = Text1.Text
      Else
        MsgBox "You must close all active chat windows", vbInformation, ""
        Text1.Text = person
    End If

End Sub

Private Sub Text3_Change()
If Right(Text3.Text, 2) <> ": " Then Text3.Text = Text3.Text & ": "
End Sub
