VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Controller"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSpacing 
      Caption         =   "Spacing"
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   4455
      Begin VB.CheckBox chkEnableSpacing 
         Caption         =   "Enable"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   $"frmMain.frx":0000
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame fraMode 
      Caption         =   "Mode"
      Height          =   975
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
      Begin VB.OptionButton optManual 
         Caption         =   "Manual"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optAutomatic 
         Caption         =   "Automatic(Recommended)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdRefreshParents 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ComboBox cmoWindows 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   4455
   End
   Begin VB.CommandButton cmdControlWindow 
      Caption         =   "Control"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtHWND 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "Version 1.00 by David Fiala - Djf1010@aol.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   4035
   End
   Begin VB.Label lblWindowHWND 
      AutoSize        =   -1  'True
      Caption         =   "Window HWND:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ___________________    ________________________________
' /                   \  /                                \
' | Window controller |--| By David Fiala djf1010@aol.com |
' \___________________/  \________________________________/
'
' Version 1.00   Released date: Sept. 03 2001

Option Explicit

Private WithEvents spchRecognizer As prjVoice.clsRecognize
Attribute spchRecognizer.VB_VarHelpID = -1
Private spchSpeaker As prjVoice.clsSpeak

Private blnMode As Boolean '0 - Control commands mode
                           '1 - Dictation mode

Private Sub cmdControlWindow_Click()
    Call modEnumChildren.DoChildEnum(frmMain.txtHWND.Text)
End Sub

Private Sub cmdRefreshParents_Click()
    Dim astrParents() As String
    Dim i As Long
    
    Call DoParentEnum
    
    Call GiveMeParents(astrParents)
    
    With frmMain.cmoWindows
        .Clear
        For i = 0 To UBound(astrParents)
            .AddItem astrParents(i)
        Next
    End With
End Sub

Private Sub cmoWindows_Click()
    frmMain.txtHWND.Text = ParentHWNDFromText(frmMain.cmoWindows.Text)
End Sub

Private Sub Form_Load()
    Set spchRecognizer = New prjVoice.clsRecognize
    Set spchSpeaker = New prjVoice.clsSpeak
    spchRecognizer.StartRecognition
    Me.Show
    Me.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmUnloading.Show
    frmUnloading.Refresh
    spchRecognizer.StopRecognition
    Set spchRecognizer = Nothing
    Set spchSpeaker = Nothing
    Unload frmUnloading
End Sub

Private Sub spchRecognizer_SpeakRecognized(ByVal strRecognizedText As String)
    Debug.Print "new: " & strRecognizedText
    
    If frmMain.optAutomatic.Value = True Then
    
    'Automatic mode
    
        If ClickIt(strRecognizedText) = 1 Then Exit Sub
        If frmMain.chkEnableSpacing.Value = 1 Then strRecognizedText = strRecognizedText & " "
        SendKeys strRecognizedText
        Exit Sub
        
    ElseIf frmMain.optManual.Value = True Then
    
    'Manual mode
    
        Select Case LCase(strRecognizedText)
            Case "control"
                blnMode = False ' Mode 0 - Control commands mode
                MsgBox "Control mode activated.", vbInformation + vbSystemModal, "Window controller"
            Case "dictation"
                blnMode = True ' Mode 1 - Dictionation mode
                MsgBox "Dictation mode activated.", vbInformation + vbSystemModal, "Window controller"
            Case Else
                Select Case blnMode
                    Case 0 ' Commands mode
                        Call ClickIt(strRecognizedText)
                    Case 1 ' Dictation mode
                        If frmMain.chkEnableSpacing.Value = 1 Then strRecognizedText = strRecognizedText & " "
                        SendKeys strRecognizedText 'Type it out to whatever is active.
                End Select
        End Select
        
    End If
End Sub
