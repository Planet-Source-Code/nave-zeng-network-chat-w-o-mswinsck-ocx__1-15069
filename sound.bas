Attribute VB_Name = "sound"
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0         '  play synchronously  (default)
Public Const SND_ASYNC = &H1        '  play asynchronously (the program does not stop for the WAV)
Public Const SND_MEMORY = &H4       '  play from memory    (from a String)

' I got this sub from Kalani
Public Sub PlayRESSound(iIndex As Integer, Optional bWait As Variant)

    On Error Resume Next                '  just in case
    Dim lFlags&
      If bWait = True Then
          lFlags = SND_SYNC Or SND_MEMORY
        Else
          lFlags = SND_ASYNC Or SND_MEMORY
      End If
    Dim vAddress$
      ' the next line does all the work.
      vAddress = StrConv(LoadResData(iIndex, "SOUND"), vbUnicode)
      sndPlaySound vAddress, lFlags

End Sub

