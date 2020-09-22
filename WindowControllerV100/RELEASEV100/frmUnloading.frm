VERSION 5.00
Begin VB.Form frmUnloading 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unloading Window Controller"
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
   Begin VB.Label lblUnloading 
      AutoSize        =   -1  'True
      Caption         =   "This may take a moment..."
      Height          =   195
      Index           =   1
      Left            =   900
      TabIndex        =   1
      Top             =   397
      Width           =   1860
   End
   Begin VB.Label lblUnloading 
      AutoSize        =   -1  'True
      Caption         =   "Please wait while the application unloads."
      Height          =   195
      Index           =   0
      Left            =   900
      TabIndex        =   0
      Top             =   157
      Width           =   2940
   End
   Begin VB.Image imgClock 
      Height          =   480
      Left            =   300
      Picture         =   "frmUnloading.frx":0000
      Top             =   157
      Width           =   480
   End
End
Attribute VB_Name = "frmUnloading"
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
