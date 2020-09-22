VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecognize"
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

Public Event SpeakRecognized(ByVal strRecognizedText As String)

Private WithEvents vceRecognizer As SpSharedRecoContext
Attribute vceRecognizer.VB_VarHelpID = -1
Private grmGrammar As ISpeechRecoGrammar

Public Sub StartRecognition()
    If vceRecognizer Is Nothing Then
        Set vceRecognizer = New SpSharedRecoContext
        Set grmGrammar = vceRecognizer.CreateGrammar(1)
        grmGrammar.DictationLoad
    End If
    
    grmGrammar.DictationSetState SGDSActive

End Sub

Public Sub StopRecognition()
    grmGrammar.DictationSetState SGDSInactive
End Sub

Private Sub vceRecognizer_Recognition(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal RecognitionType As SpeechLib.SpeechRecognitionType, ByVal Result As SpeechLib.ISpeechRecoResult)
    RaiseEvent SpeakRecognized(Result.PhraseInfo.GetText)
End Sub
