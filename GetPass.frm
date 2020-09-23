VERSION 5.00
Begin VB.Form GetPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter private key password"
   ClientHeight    =   1005
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "GetPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   593.787
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "getpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Entered As Boolean
Private Sub cmdCancel_Click()
Entered = False
Me.Hide
End Sub
Private Sub cmdOK_Click()
gotpass = txtPassword
txtPassword = ""
Entered = True
Select Case runwhat
Case "import"
Main.ImportKeypair2
Case "gen"
Main.MakeKeypair2
End Select
Me.Hide
End Sub

