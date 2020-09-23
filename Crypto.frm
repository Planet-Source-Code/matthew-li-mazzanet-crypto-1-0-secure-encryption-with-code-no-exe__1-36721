VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crypto"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   Icon            =   "Crypto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crypto.frx":0442
            Key             =   "GenKeypair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crypto.frx":0896
            Key             =   "ImportKeypair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crypto.frx":0CEA
            Key             =   "EncryptFile"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crypto.frx":113E
            Key             =   "DecryptFile"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Crypto.frx":1592
            Key             =   "WipeFile"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   1588
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "GenKeypair"
            Object.ToolTipText     =   "Generate New Keypair"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ImportKeypair"
            Object.ToolTipText     =   "Import a keypair"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "EncryptFile"
            Object.ToolTipText     =   "Encrypt a file"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "DecryptFile"
            Object.ToolTipText     =   "Decrypt a file"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "WipeFile"
            Object.ToolTipText     =   "Wipe a file"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "c:"
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   0
      X2              =   6360
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objEncryption As CryptoCls
Public strKey As String
Public keyname As String

Private Sub FileDecrypt()
Dim sngTim As Single
Dim strFileAndPathName As String
Toolbar.Buttons(4).Enabled = False
CommonDialog1.Filter = ""
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
strFileAndPathName = CommonDialog1.FileName
If strFileAndPathName = "" Then GoTo doexit
sngTim = Timer
objEncryption.DecryptFile_KeyPair strFileAndPathName, strFileAndPathName & ".DECRYPTED"
MsgBox "Finished decrypting the file.  Remove the '.DECRYPTED' file extension to open the decrypted file. " & vbNewLine & "Time elapsed: " & Timer - sngTim
doexit:
Toolbar.Buttons(4).Enabled = True
Exit Sub
End Sub

Private Sub FileEncrypt()
Dim sngTim As Single
Dim strFileAndPathName As String
Toolbar.Buttons(3).Enabled = False
CommonDialog1.Filter = ""
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
strFileAndPathName = CommonDialog1.FileName
If strFileAndPathName = "" Then GoTo doexit
sngTim = Timer
objEncryption.EncryptFile_KeyPair strFileAndPathName, strFileAndPathName & ".ENCRYPTED"
MsgBox "Finished encrypting the file.  Remove the '.ENCRYPTED' file extension to open the encrypted file. " & vbNewLine & "Time elapsed: " & Timer - sngTim
doexit:
Toolbar.Buttons(3).Enabled = True
Exit Sub
End Sub

Private Sub Form_Load()
MsgBox "Welcome to Crypto 1.0!" & vbCrLf & "Copyright (c) Matthew Li 2002"
Set objEncryption = New CryptoCls
objEncryption.SessionStart
End Sub

Private Sub Form_Unload(Cancel As Integer)
objEncryption.SessionEnd
Set objEncryption = Nothing
Unload getpass
Unload Wipe
End Sub

Private Sub ImportKeypair()
Dim intNextFreeFile As Integer
Dim strFileAndPathName As String
CommonDialog1.Filter = "Crypto Keys (.prk; .pbk;)|*.prk;*.pbk;"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
strFileAndPathName = CommonDialog1.FileName
If strFileAndPathName = "" Then Exit Sub
intNextFreeFile = FreeFile
Open strFileAndPathName For Binary As #intNextFreeFile
strKey = String(LOF(intNextFreeFile), vbNullChar)
Get #intNextFreeFile, , strKey
Close #intNextFreeFile
Select Case Right(strFileAndPathName, 4)
Case ".prk"
objEncryption.ValuePublicPrivateKey = String(Len(strKey), vbNullChar)
objEncryption.ValuePublicPrivateKey = strKey
runwhat = "import"
getpass.Show
Case ".pbk"
objEncryption.ValuePublicKey = String(Len(strKey), vbNullChar)
objEncryption.ValuePublicKey = strKey
objEncryption.Import_KeyPair
MsgBox "Imported the Public key."
Toolbar.Buttons(3).Enabled = True
Toolbar.Buttons(4).Enabled = False
Main.Caption = "Crypto - " & CommonDialog1.FileName
Case Else
MsgBox "Not a key file.  Did not import a key."
End Select
End Sub

Private Sub MakeKeypair()
Dim intNextFreeFile As Integer
Dim strFileAndPathName As String
objEncryption.Generate_KeyPair
runwhat = "gen"
getpass.Show
End Sub
Public Sub MakeKeypair2()
Dim intNextFreeFile As Integer
objEncryption.Export_KeyPair gotpass
intNextFreeFile = FreeFile
keyname = InputBox("Enter a name for your keypair.")
Open "C:\" & keyname & ".prk" For Binary Access Write As #intNextFreeFile
Put #intNextFreeFile, , objEncryption.ValuePublicPrivateKey
Close #intNextFreeFile
intNextFreeFile = FreeFile
Open "C:\" & keyname & ".pbk" For Binary Access Write As #intNextFreeFile
Put #intNextFreeFile, , objEncryption.ValuePublicKey
Close #intNextFreeFile
MsgBox "Generated new key pair and exported the keys to C:\" & keyname & ".prk and C:\" & keyname & ".pbk"
Toolbar.Buttons(3).Enabled = True
Toolbar.Buttons(4).Enabled = True
Main.Caption = "Crypto - " & "C:\" & keyname & ".prk"
End Sub
Public Sub ImportKeypair2()
objEncryption.Import_KeyPair gotpass
MsgBox "Imported the Private key."
Toolbar.Buttons(3).Enabled = True
Toolbar.Buttons(4).Enabled = True
Main.Caption = "Crypto - " & CommonDialog1.FileName
End Sub

Private Sub WipeFile()
Wipe.Show
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "GenKeypair"
Call MakeKeypair
Case "ImportKeypair"
Call ImportKeypair
Case "EncryptFile"
Call FileEncrypt
Case "DecryptFile"
Call FileDecrypt
Case "WipeFile"
Call WipeFile
End Select
End Sub
