VERSION 5.00
Begin VB.Form frmLoading 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading Window Controller"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgClock 
      Height          =   480
      Left            =   390
      Picture         =   "frmLoading.frx":0000
      Top             =   157
      Width           =   480
   End
   Begin VB.Label lblUnloading 
      AutoSize        =   -1  'True
      Caption         =   "Please wait while the application loads."
      Height          =   195
      Index           =   0
      Left            =   990
      TabIndex        =   1
      Top             =   157
      Width           =   2760
   End
   Begin VB.Label lblUnloading 
      AutoSize        =   -1  'True
      Caption         =   "This may take a moment..."
      Height          =   195
      Index           =   1
      Left            =   990
      TabIndex        =   0
      Top             =   397
      Width           =   1860
   End
End
Attribute VB_Name = "frmLoading"
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

Private Sub Form_Load()
    Me.Show
    Me.Refresh
    frmMain.Show
    Unload Me
End Sub
