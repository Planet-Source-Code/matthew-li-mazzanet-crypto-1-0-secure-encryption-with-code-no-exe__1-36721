VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Wipe 
   Caption         =   "Wipe Wizard"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "Wipe.frx":0000
   LinkTopic       =   "Wipe"
   ScaleHeight     =   2670
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton exitwipe 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtLoops 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Text            =   "27"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton loadfile 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   1080
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton GoWipe 
      Caption         =   "GO"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "C:\Windows\Desktop\"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label pass 
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   $"Wipe.frx":0442
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Passes"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   240
   End
End
Attribute VB_Name = "Wipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub SFSWipeFile(sFile As String, Delete As Boolean)
Dim sLen As Long, X As Integer, FF As Byte
Dim Rand As Byte, Char As String
Randomize
For X = 0 To 35
If X = 17 Then PB.Value = PB.Value + 5
FF = FreeFile
sLen = FileLen(sFile)
Open sFile For Binary As FF
Rand = Int(255 * Rnd) + 1
Char = String(sLen, Chr(Rand))
Put #FF, , Char
Close FF
Reset
Refresh
Next X
If Delete = True Then Kill sFile
End Sub

Public Sub WipeFile(sFile As String, Loops As Integer, Delete As Boolean)
Dim sLen As Long, X As Integer, FF As Byte
Dim Rand As Byte, Char As String
Randomize
For X = 0 To Loops
If X = CInt(Loops / 2) Then PB.Value = PB.Value + 5
FF = FreeFile
sLen = FileLen(sFile)
Open sFile For Binary As FF
Rand = Int(255 * Rnd) + 1
Char = String(sLen, Chr(Rand))
Put #FF, , Char
Close FF
Reset
Refresh
Next X
If Delete = True Then Kill sFile
End Sub

Public Sub DoBlankWipe(sFile As String, Delete As Boolean)
Dim sLen As Long, X As Integer, FF As Byte
Dim Char As String, Num
Randomize
For X = 1 To 3
sLen = FileLen(sFile)
FF = FreeFile
Open sFile For Binary As FF
Select Case X
Case 0
Num = 0
Case 1
Num = 255
Case 2
Num = 0
End Select
Char = String(sLen, Chr$(Num))
Put #FF, , Char
Close FF
Reset
Refresh
Next X
If Delete = True Then Kill sFile
End Sub

Private Sub exitwipe_Click()
Wipe.Hide
End Sub

Private Sub gowipe_Click()
Dim origtimer
Dim choice
choice = InputBox("If you are certain about destroying: " & txtFile.text & " Then type 'YES' in the box below.", "Are you certain?")
If choice = "YES" Then
origtimer = Timer
exitwipe.Enabled = False
GoWipe.Enabled = False
txtLoops.Enabled = False
loadfile.Enabled = False
For i = 1 To txtLoops.text
pass.Caption = i & "/" & txtLoops.text
PB.Value = 0
Call WipeFile(txtFile.text, 1, False)
PB.Value = 20
Call DoBlankWipe(txtFile.text, False)
PB.Value = 40
Call SFSWipeFile(txtFile.text, False)
PB.Value = 60
Call WipeFile(txtFile.text, 20, False)
PB.Value = 80
Call WipeFile(txtFile.text, txtLoops.text, False)
PB.Value = 100
Next i
Kill txtFile.text
MsgBox "Successfully destroyed " & txtFile.text & " in " & Timer - origtimer & " seconds!", , "Successful destroy!"
exitwipe.Enabled = True
GoWipe.Enabled = True
txtLoops.Enabled = True
loadfile.Enabled = True
End If
End Sub

Private Sub loadfile_Click()
CommonDialog1.Filter = ""
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
txtFile.text = CommonDialog1.FileName
End Sub
