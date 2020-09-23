Attribute VB_Name = "basgeneral"
Global thefullpath As String
'set thefullpath to where you want the data folder
'to be placed. If you leave it empty then the data
'folder is placed in the same as the exe.
'If you want each user to have a copy of the exe
'then set thefullpath equal to something.
'If you want the exe to stay on a common computer
'that can be accessed by any computer in the network
'then leave it empty.
'Also make sure it ends with a "\"
Global mgradient, mcolor As Long
'I added these variables to the metallic form code
'so that the user could customize it.
Global person, room As String
Global background As Integer
' What the background is
Global chatcount(9), closedforgood, adminon As Boolean
Global chatroom(9) As String
Global chatnumber As Integer

Public Sub killpeople(lstNames As ListBox, myroom As String)

    On Error Resume Next
      Do While listnum < lstNames.ListCount
          thename = "person" & myroom & lstNames.List(listnum)
          If Dir$(thefullpath & "data\" & thename) <> "" Then
              Open thefullpath & "data\" & thename For Input As #13
              Line Input #13, curnuma
              Close #13
          End If
          curtime = Timer
          Do While Timer < curtime + 0.3
              DoEvents
              If closedforgood = True Then Exit Sub
          Loop
          If Dir$(thefullpath & "data\" & thename) <> "" Then
              Open thefullpath & "data\" & thename For Input As #14
              Line Input #14, curnumb
              Close #14
          End If
          curtime = Timer
          Do While Timer < curtime + 0.3
              DoEvents
              If closedforgood = True Then Exit Sub
          Loop
          If Dir$(thefullpath & "data\" & thename) <> "" Then
              Open thefullpath & "data\" & thename For Input As #15
              Line Input #15, curnumc
              Close #15
          End If
          If curnuma = curnumb And curnuma = curnumc Then Kill thefullpath & "data\" & thename
          listnum = listnum + 1
      Loop

End Sub

