VERSION 5.00
Begin VB.Form mainchat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Chat coded by Navarchy"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   Icon            =   "mainchat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&A"
      Height          =   195
      Left            =   9999
      TabIndex        =   18
      Top             =   0
      Width           =   135
   End
   Begin VB.Timer newpeople 
      Interval        =   150
      Left            =   840
      Top             =   840
   End
   Begin VB.Timer newroom 
      Interval        =   50
      Left            =   840
      Top             =   360
   End
   Begin VB.Frame Frame2 
      Caption         =   "Administrator Options"
      Height          =   3135
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   4125
      Begin VB.TextBox Text3 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   19
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Left            =   3840
         TabIndex        =   21
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Confirmation Password"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Background Editor"
      Height          =   5055
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   4125
      Begin VB.PictureBox picColors 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3015
         Left            =   120
         Picture         =   "mainchat.frx":0ECA
         ScaleHeight     =   2955
         ScaleWidth      =   2955
         TabIndex        =   20
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox col 
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Text            =   "30"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox grad 
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Text            =   "3.142"
         Top             =   960
         Width           =   1575
      End
      Begin VB.PictureBox picColorBox 
         AutoRedraw      =   -1  'True
         Height          =   3015
         Left            =   3240
         ScaleHeight     =   2955
         ScaleWidth      =   705
         TabIndex        =   10
         Top             =   1920
         Width           =   765
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Metallic"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Plain"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "color"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "gradient"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Left            =   3840
         TabIndex        =   15
         ToolTipText     =   "Exit background editor"
         Top             =   120
         Width           =   210
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Text            =   "Type message here..."
      Top             =   2170
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
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
      Left            =   2660
      TabIndex        =   4
      Top             =   2670
      Width           =   810
   End
   Begin VB.Image imgmsg 
      Height          =   300
      Index           =   0
      Left            =   2350
      Picture         =   "mainchat.frx":1B0A
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   15
      Left            =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
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
      Left            =   3880
      TabIndex        =   6
      ToolTipText     =   "Background Editor"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Go To Room"
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
      Left            =   470
      TabIndex        =   3
      Top             =   2670
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image imgmsg 
      Height          =   300
      Index           =   1
      Left            =   2350
      Picture         =   "mainchat.frx":22AC
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Image imggoroom 
      Height          =   300
      Index           =   0
      Left            =   300
      Picture         =   "mainchat.frx":2C7C
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Image imgsend 
      Height          =   300
      Index           =   1
      Left            =   300
      Picture         =   "mainchat.frx":341E
      Top             =   2640
      Width           =   1425
   End
End
Attribute VB_Name = "mainchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.10.32 generated this copy of mainchat.frm on Thu 1/25/01 @ 4:29 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  imgImages_Click               Label10_Click                                           *
'******************************************************************************************

'These are the declarations used for the sub
'CreateNewDirectory(), see the comments in that sub
'for more details
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Dim chatclone(9) As New chat
Dim started As Boolean
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Sub CreateNewDirectory(NewDirectory As String)

  'This is the code used to create the directories
  'needed by the program, such as people and  room
  'folders. I could have created them and the run
  'the program, but I want it to be as self contained
  'as possible.
  '
  'I did not write this code
  'http://www.vb-world.net/tips/tip45.html

  Dim sDirTest As String
  Dim SecAttrib As SECURITY_ATTRIBUTES
  Dim bSuccess As Boolean
  Dim sPath As String
  Dim iCounter As Integer
  Dim sTempDir As String
  Dim iFlag

    iFlag = 0
    sPath = NewDirectory + "\x"

    If Right$(sPath, Len(sPath)) <> "" Then
        sPath = sPath & ""
    End If

    iCounter = 1

    Do Until InStr(iCounter, sPath, "") = Len(sPath)
        iCounter = InStr(iCounter, sPath, "")
        sTempDir = Left$(sPath, iCounter)
        sDirTest = Dir$(sTempDir)
        iCounter = iCounter + 1
        If Right$(sTempDir, 1) <> "\" Then GoTo Missit
        'create directory
        SecAttrib.lpSecurityDescriptor = &O0
        SecAttrib.bInheritHandle = False
        SecAttrib.nLength = Len(SecAttrib)
        bSuccess = CreateDirectory(sTempDir, SecAttrib)
Missit:
    Loop

End Sub

Private Sub col_Change()

    If IsNumeric(col.Text) = False Or col.Text = 0 Then Exit Sub
    mcolor = col.Text
    Call backgroundupdate(Me)

End Sub

Private Sub Command1_Click()

    If Frame2.Visible = False Then
        Frame1.Visible = False
        Me.Height = 3495
        Frame2.Visible = True
        Call PlayRESSound(3, False)
        Text3.SetFocus
      Else
        Frame2.Visible = False
        Me.SetFocus
    End If

End Sub

Private Sub Form_Load()

    On Error Resume Next
      'thefullpath = "\\rocscurr2\jr-common\"
      If App.PrevInstance = True Then MsgBox "Sorry, but you are already running this program", , "": Unload Me: End: Exit Sub
      mcolor = 30
      mgradient = 3.142
      Me.ScaleMode = 3
      Me.AutoRedraw = True
      mytop = Me.Top
      myleft = Me.Left
      '    Call backgroundupdate(Me)
    Dim s As String
    Dim cnt As Long
    Dim dl As Long
      cnt = 199
      s = String$(200, 0)
      dl = GetUserName(s, cnt)
      If dl <> 0 Then person = Left$(s, cnt - 1)
      If Left$(person, 4) = "boot" Then MsgBox "Invalid Username", vbalert, "": Unload Me: End
      If Left$(person, 4) = "room" Then MsgBox "Invalid Username", vbalert, "": Unload Me: End
      If Left$(person, 6) = "person" Then MsgBox "Invalid Username", vbalert, "": Unload Me: End
      If Right$(person, 4) = ".num" Then MsgBox "Invalid Username", vbalert, "": Unload Me: End
      If person = "" Then MsgBox "It appears as though you are not logged on", vbalert, "": Unload Me: End
      If Dir$(thefullpath & "data", vbDirectory) = "" Then Call CreateNewDirectory(thefullpath & "data")
      Me.Top = Screen.Height / 2 - Me.Height / 2
      Me.Left = Screen.Width / 2 - Me.Width / 2
      a = Dir$(thefullpath & "data\person*" & person)
      Do While a <> ""
          Kill thefullpath & "data\" & a
          a = Dir$()
      Loop
      started = True
      Call PlayRESSound(2, False)
      syspath = GetSystemPath()
    
    Dim FileNumber As Integer
    Dim DllBuffer() As Byte

      DllBuffer = LoadResData(5, "CUSTOM")
      FileNumber = FreeFile
      If Dir$(syspath & "\" & "richtx32.ocx") <> "" Then Exit Sub
      Open syspath & "\" & "richtx32.ocx" For Binary Access Write As #FileNumber
      Put #FileNumber, , DllBuffer
      Close #FileNumber

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call boom

End Sub

Public Sub boom()

    On Error Resume Next
      If App.PrevInstance = True Then Exit Sub
      admin.onoroff.Interval = 0
      admin.newroom.Interval = 0
      admin.newpeople.Interval = 0
      newroom.Interval = 0
      newpeople.Interval = 0
      closedforgood = True
      If started = False Then
      Do While Y < 10
      Unload chatclone(Y)
      Y = Y + 1
      Loop
      Unload Me
      Unload admin
      End
      Exit Sub
      End If
      Call PlayRESSound(4, False)
      'Me.Hide
      a = Me.Height / 10
      b = Me.Width / 10
      c = 323
      d = 619
      e = admin.Height / 10
      f = admin.Width / 10
      For X = 0 To 9
          cur = Timer
          Do While Timer < cur + 0.25
              DoEvents
          Loop
          t = 0
          For t = 0 To 9
              If chatcount(t) = True Then
                  chatclone(t).Width = chatclone(t).Width - d
                  chatclone(t).Height = chatclone(t).Height - c
              End If
          Next t
          If adminon = True Then
              admin.Height = admin.Height - e
              admin.Width = admin.Width - f
          End If
          Me.Height = Me.Height - a
          Me.Width = Me.Width - b
      Next X
      Do While Y < 10
          Unload chatclone(Y)
          Y = Y + 1
      Loop
      Unload Me
      Unload admin
      End

End Sub

Private Sub Form_Resize()

    Call backgroundupdate(Me)

End Sub

Private Sub Form_Terminate()

    Call boom

End Sub

Private Sub grad_Change()

    If IsNumeric(grad.Text) = False Or grad.Text = 0 Then Exit Sub
    mgradient = grad.Text
    Call backgroundupdate(Me)

End Sub

Private Sub imggoroom_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    imggoroom(0).Visible = False

End Sub

Private Sub imggoroom_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    imggoroom(0).Visible = True
    Call gotoroom

End Sub

Private Sub imgmsg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgmsg(0).Visible = False

End Sub

Private Sub imgmsg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call sendmsg

End Sub

Private Sub Label11_Click()

    Frame2.Visible = False

End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imggoroom(0).Visible = False

End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imggoroom(0).Visible = True
    Call gotoroom

End Sub

Public Sub gotoroom()

    On Error Resume Next
      If List1.Text = "" Then MsgBox "You must select a room", vbInformation, "": Exit Sub
      room = List1.Text
      For c = 0 To 9
          If room = chatroom(c) Then
              MsgBox "You are already chatting in this room", vbExclamation, ""
              Exit Sub
          End If
      Next c
      Do While a = False
          If chatcount(b) = False Then
              chatnumber = b
              chatclone(b).Show
              chatcount(b) = True
              If background = 0 Then Call GradientFill(chatclone(b))
              If background = 2 Then Call makegray(chatclone(b), tColor)
              a = True
              chatroom(b) = room
          End If
          If b = 10 Then MsgBox "Sorry, but you already have ten rooms open", vbExclamation, "": a = True
          b = b + 1
      Loop

End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgmsg(0).Visible = False

End Sub

Public Sub sendmsg()

    imgmsg(0).Visible = True
    If List2.Text = "" Then MsgBox "You need to select a name", vbInformation, "": Exit Sub
    Shell "net send " & List2.Text & " " & Text1.Text, vbHide

End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call sendmsg

End Sub

Private Sub Label6_Click()

    Frame1.Visible = True
    Me.Height = 5430

End Sub

Private Sub Label9_Click()

    Frame1.Visible = False
    Me.Height = 3495

End Sub

Private Sub List1_Click()

    List2.Clear
    a = Dir$(thefullpath & "data\person" & List1.Text & "*")
    Do While a <> ""
        List2.AddItem Right$(a, Len(a) - Len("person" & List1.Text))
        a = Dir$()
    Loop

End Sub

Private Sub List1_DblClick()

    Call gotoroom

End Sub

Private Sub newpeople_Timer()

    On Error Resume Next

      If List1.Text = "" Then Exit Sub

      newname = Dir$(thefullpath & "data\person" & List1.Text & "*")

      Do While newname <> ""
          listnum = 0
          w = False
          Do While listnum < List2.ListCount
              If List2.List(listnum) = Right$(newname, Len(newname) - Len("person" & List1.Text)) Then w = True
              listnum = listnum + 1
          Loop
          If w = False Then List2.AddItem Right$(newname, Len(newname) - Len("person" & List1.Text))
          newname = Dir$()
      Loop

      listnum = 0
      Do While listnum < List2.ListCount
          w = False
          newname = Dir$(thefullpath & "data\person" & List1.Text & "*")
          Do While newname <> ""
              If Right$(newname, Len(newname) - Len("person" & List1.Text)) = List2.List(listnum) Then w = True
              newname = Dir$()
          Loop
          If w = False Then List2.RemoveItem (listnum)
          listnum = listnum + 1
      Loop
      If Label1.Caption <> List2.ListCount & " users in " & List1.Text Then Label1.Caption = List2.ListCount & " users in " & List1.Text
      Call killpeople(Me.List2, List1.Text)

End Sub

Private Sub newroom_Timer()

    If Dir$(thefullpath & "data\boot" & person) <> "" Then
        Kill thefullpath & "data\boot" & person
        Unload Me
    End If
    If Dir$(thefullpath & "data\shutdownnet") <> "" Then
        If adminon = False Then
            Unload Me
        End If
    End If

    On Error Resume Next
      If List1.Text = "" Then
          If Label1.Caption <> "0 users in no room selected" Then Label1.Caption = "0 users in no room selected"
          If List2.ListCount <> 0 Then List2.Clear
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

Private Sub Option1_Click()

    picColorBox.BackColor = &H8000000F
    tColor = 0
    background = 2
    Call makegray(Me, tColor)

End Sub

Private Sub Option3_Click()

  'Me.Picture = Image1.Picture

    Me.Cls
    background = 0
    Call GradientFill(Me)

End Sub

Private Sub Text1_Change()

    If Text1.Text = "Type message here..." Then Text1.Text = ""

End Sub

Private Sub Text1_Click()

    If Text1.Text = "Type message here..." Then Text1.Text = ""

End Sub

Private Sub Text3_Change()

    If Text3.Text = "scow" Then
        LockWindowUpdate Me.hWnd
        Text3.Text = ""
        Frame2.Visible = False
        admin.Show
        LockWindowUpdate 0
    End If

End Sub

Private Sub picColors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Local Error Resume Next

      If Button = vbLeftButton Then
          picColorBox.BackColor = picColors.Point(X, Y)
      End If

End Sub

Private Sub picColors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tColor = picColors.Point(X, Y)
    If background = 2 Then
        Call makegray(Me, tColor)
    End If

End Sub

