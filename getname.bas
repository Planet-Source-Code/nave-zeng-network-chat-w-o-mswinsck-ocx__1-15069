Attribute VB_Name = "getname"
'This code is used to get the current user logged
'onto the system which doubles as the chat name.
'This code was found at the website
'http://www.vb-world.net/tips/tip20.html
Declare Function GetUserName Lib "advapi32.dll" Alias _
        "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) _
        As Long

