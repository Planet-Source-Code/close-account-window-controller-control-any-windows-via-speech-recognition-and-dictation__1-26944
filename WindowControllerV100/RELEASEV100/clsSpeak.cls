VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpeak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'prjVoice: TTS&Dictation.
'Made for use in:
'  ___________________    ________________________________
' /                   \  /                                \
' | Window controller |--| By David Fiala djf1010@aol.com |
' \___________________/  \________________________________/
'
' Version 1.00   Released date: Sept. 03 2001

Option Explicit

Private vceSpeaker As SpVoice

Private Sub Class_Initialize()
    Set vceSpeaker = New SpVoice
End Sub

Public Sub Speak(ByVal strWhatToSay As String)
    vceSpeaker.Speak strWhatToSay
End Sub
