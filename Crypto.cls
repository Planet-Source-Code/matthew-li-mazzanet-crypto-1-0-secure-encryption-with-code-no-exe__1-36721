VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptoCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hSessionKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef hSessionKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hSessionKey As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptGetKeyParam Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptSetKeyParam Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal dwParam As Long, ByVal pbData As String, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long
Private Declare Function CryptGenKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal dwFlags As Long, ByRef hSessionKey As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptSetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByVal dwFlags As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptGetUserKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwKeySpec As Long, ByVal phUserKey As Long) As Long
Private Declare Function CryptSignHash Lib "advapi32.dll" Alias "CryptSignHashA" (ByVal hHash As Long, ByVal dwKeySpec As Long, ByVal sDescription As String, ByVal dwFlags As Long, ByVal pbSignature As String, pdwSigLen As Long) As Long
Private Declare Function CryptVerifySignature Lib "advapi32.dll" Alias "CryptVerifySignatureA" (ByVal hHash As Long, ByVal pbSignature As String, ByVal dwSigLen As Long, ByVal hPubKey As Long, ByVal sDescription As String, ByVal dwFlags As Long) As Long
Private Const SERVICE_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0" & vbNullChar
Private Const KEY_CONTAINER As String = "Crypto" & vbNullChar
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const CRYPT_DELETEKEYSET As Long = 16
Private Const CRYPT_CREATE_SALT As Long = 4
Private Const CRYPT_EXPORTABLE As Long = 1
Private Const KP_SALT As Long = 2
Private Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Private Const ALG_CLASS_HASH As Long = 32768
Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_TYPE_STREAM As Long = 2048
Private Const ALG_SID_RC4 As Long = 1
Private Const ALG_SID_MD5 As Long = 3
Private Const ALG_SID_SHA As Long = 4
Private Const CALG_MD5 As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC4 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
Private Const CALG_SHA As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA)
Private Const AT_KEYEXCHANGE As Long = 1
Private Const AT_SIGNATURE As Long = 2
Private Const HP_HASHVAL As Long = 2
Private Const SIMPLEBLOB = 1
Private Const PUBLICKEYBLOB As Long = 6
Private Const PRIVATEKEYBLOB As Long = 7
Private Const CRYPT_NO_SALT = 16
Private Const NTE_BAD_SIGNATURE As Long = -2146893818
Dim hCryptProv As Long
Dim strHash As String
Dim hSessionKey As Long
Dim hKeyPair As Long
Dim strPublicPrivateBlob As String
Dim strPublicBlob As String
Dim strSessionBlob As String
Dim lngType As Long
Dim strSig As String
Public Property Get ValuePublicPrivateKey() As String
ValuePublicPrivateKey = strPublicPrivateBlob
End Property
Public Property Let ValuePublicPrivateKey(strValue As String)
strPublicPrivateBlob = String(Len(strValue), vbNullChar)
strPublicPrivateBlob = strValue
strPublicBlob = vbNullString
End Property
Public Property Get ValuePublicKey() As String
ValuePublicKey = strPublicBlob
End Property
Public Property Let ValuePublicKey(strValue As String)
strPublicBlob = String(Len(strValue), vbNullChar)
strPublicBlob = strValue
strPublicPrivateBlob = vbNullString
End Property
Public Property Get ValueSessionKey() As String
ValueSessionKey = strSessionBlob
End Property
Public Property Let ValueSessionKey(strValue As String)
strSessionBlob = String(Len(strValue), vbNullChar)
strSessionBlob = strValue
End Property
Public Sub Generate_KeyPair()
Dim lngParams As Long
Dim lngReturnValue As Long
lngParams = &H2000000 Or CRYPT_EXPORTABLE Or CRYPT_NO_SALT
lngType = AT_KEYEXCHANGE
If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
lngReturnValue = CryptGenKey(hCryptProv, lngType, lngParams, hKeyPair)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not generate public/private key pair"
End Sub
Public Sub Export_KeyPair(ByVal PrivateKeyPassword As String)
Dim lngReturnValue As Long
Dim lngLength As Long
lngReturnValue = CryptExportKey(hKeyPair, 0, PUBLICKEYBLOB, 0, vbNull, lngLength)
strPublicBlob = String(lngLength, vbNullChar)
lngReturnValue = CryptExportKey(hKeyPair, 0, PUBLICKEYBLOB, 0, strPublicBlob, lngLength)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not export the public key blob"
SessionKeyFromPassword (PrivateKeyPassword)
lngReturnValue = CryptExportKey(hKeyPair, hSessionKey, PRIVATEKEYBLOB, 0, vbNull, lngLength)
strPublicPrivateBlob = String(lngLength, vbNullChar)
lngReturnValue = CryptExportKey(hKeyPair, hSessionKey, PRIVATEKEYBLOB, 0, strPublicPrivateBlob, lngLength)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not export the private key blob"
End Sub
Public Sub Import_KeyPair(Optional ByVal PrivateKeyPassword As String = vbNullString)
Dim lngReturnValue As Long
Dim lngLength As Long
Dim lngParams As Long
If strPublicPrivateBlob = vbNullString And strPublicBlob = vbNullString Then
Err.Raise vbObjectError + 1, , "One of the ValueXXXKey properties must hold a valid key"
End If
If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
lngType = AT_KEYEXCHANGE
If strPublicPrivateBlob = vbNullString Then
lngLength = Len(strPublicBlob)
lngReturnValue = CryptImportKey(hCryptProv, strPublicBlob, lngLength, 0, 0, hKeyPair)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not import the Public key"
Else
SessionKeyFromPassword PrivateKeyPassword
lngLength = Len(strPublicPrivateBlob)
lngParams = CRYPT_EXPORTABLE Or CRYPT_NO_SALT
lngReturnValue = CryptImportKey(hCryptProv, strPublicPrivateBlob, lngLength, hSessionKey, lngParams, hKeyPair)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not import the PublicPrivate key.  Most likely an incorrect private key password was entered."
End If
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
End Sub

Public Sub EncryptFile_KeyPair(ByVal SourceFile As String, ByVal DestFile As String)
Dim intNextFreeFile As Integer
Dim arrByteArray() As Byte
Dim X As Long
Dim y As Long
Dim z As Long
strSessionBlob = ""
intNextFreeFile = FreeFile
Open SourceFile For Binary As #intNextFreeFile
ReDim arrByteArray(0 To LOF(intNextFreeFile) - 1)
Get #intNextFreeFile, , arrByteArray()
Close #intNextFreeFile
EncryptByteArray_KeyPair arrByteArray()
On Error Resume Next
Kill DestFile
On Error GoTo 0
z = UBound(arrByteArray) + 1
y = 1
ReDim Preserve arrByteArray(0 To (UBound(arrByteArray) + Len(strSessionBlob) + 1))
For X = z To z + Len(strSessionBlob) - 1
arrByteArray(X) = Asc(Mid(strSessionBlob, y, 1))
y = y + 1
Next X
arrByteArray(UBound(arrByteArray)) = Len(strSessionBlob) 'length of blob is last character
intNextFreeFile = FreeFile
Open DestFile For Binary As #intNextFreeFile
Put #intNextFreeFile, , arrByteArray()
Close #intNextFreeFile
End Sub
Public Function DecryptFile_KeyPair(ByVal SourceFile As String, ByVal DestFile As String)
Dim intNextFreeFile As Integer
Dim arrByteArray() As Byte
Dim X As Long
Dim y As Long
strSessionBlob = ""
intNextFreeFile = FreeFile
Open SourceFile For Binary As #intNextFreeFile
ReDim arrByteArray(0 To LOF(intNextFreeFile) - 1)
Get #intNextFreeFile, , arrByteArray()
Close #intNextFreeFile
y = arrByteArray(UBound(arrByteArray))
For X = UBound(arrByteArray) - y To UBound(arrByteArray) - 1
strSessionBlob = strSessionBlob & Chr(arrByteArray(X))
Next X
ReDim Preserve arrByteArray(0 To (UBound(arrByteArray) - Len(strSessionBlob) - 1))
DecryptByteArray_KeyPair arrByteArray()
On Error Resume Next
If FileLen(DestFile) > 0 Then Kill DestFile
On Error GoTo 0
intNextFreeFile = FreeFile
Open DestFile For Binary As #intNextFreeFile
Put #intNextFreeFile, , arrByteArray()
Close #intNextFreeFile
End Function
Public Function EncryptString_KeyPair(ByVal text As String) As String
EncryptString_KeyPair = EncryptDecrypt_KeyPair(text, True)
End Function
Public Function DecryptString_KeyPair(ByVal text As String) As String
DecryptString_KeyPair = EncryptDecrypt_KeyPair(text, False)
End Function
Public Function EncryptByteArray_KeyPair(ByRef arrByteArray() As Byte)
Dim strInput As String
Dim strOutput As String
strInput = StrConv(arrByteArray(), vbUnicode) 'convert to string
strOutput = EncryptDecrypt_KeyPair(strInput, True) 'return the encrypted data
arrByteArray() = StrConv(strOutput, vbFromUnicode) 'convert to byte array
End Function
Public Function DecryptByteArray_KeyPair(ByRef arrByteArray() As Byte)
Dim strInput As String
Dim strOutput As String
strInput = StrConv(arrByteArray(), vbUnicode) 'convert to string
strOutput = EncryptDecrypt_KeyPair(strInput, False) 'return the encrypted data
arrByteArray() = StrConv(strOutput, vbFromUnicode) 'convert to byte array
End Function
Public Sub SessionStart()
Dim lngReturnValue As Long
lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) 'try to make a new key container
If lngReturnValue = 0 Then
lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0) 'try to get a handle to a key container that already exists, and if it fails...
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , " Error during CryptAcquireContext for a new key container." & vbCrLf & "A container with this name probably already exists."
End If
End Sub
Public Sub SessionEnd()
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0
End Sub
Private Function EncryptDecrypt_KeyPair(ByVal text As String, Encrypt As Boolean) As String
Dim lngLength As Long
Dim lngReturnValue As Long
If Encrypt Then
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
lngReturnValue = CryptGenKey(hCryptProv, CALG_RC4, CRYPT_EXPORTABLE, hSessionKey)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create session key for use in key pair encryption"
If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
lngReturnValue = CryptGetUserKey(hCryptProv, lngType, hKeyPair)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not obtain public key for use in key pair encryption"
lngReturnValue = CryptExportKey(hSessionKey, hKeyPair, SIMPLEBLOB, 0, vbNull, lngLength) 'get the size of the buffer needed for the BLOB
strSessionBlob = String(lngLength, vbNullChar)
lngReturnValue = CryptExportKey(hSessionKey, hKeyPair, SIMPLEBLOB, 0, strSessionBlob, lngLength) 'get the BLOB
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not export the session key during key pair encryption"
lngLength = Len(text)
lngReturnValue = CryptEncrypt(hSessionKey, 0, 1, 0, text, lngLength, lngLength)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Error during key pair CryptEncrypt encryption."
Else
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
lngLength = Len(strSessionBlob)
lngReturnValue = CryptImportKey(hCryptProv, strSessionBlob, lngLength, hKeyPair, 0, hSessionKey)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not import the session key during key pair decryption.  Most likely an incorrect private key was used, thus it could not decrypt the session key."
If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
lngReturnValue = CryptGetUserKey(hCryptProv, lngType, hKeyPair)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not obtain private key for use in key pair decryption"
lngLength = Len(text)
lngReturnValue = CryptDecrypt(hSessionKey, 0, 1, 0, text, lngLength)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Error during key pair CryptDecrypt decryption."
End If
EncryptDecrypt_KeyPair = Left$(text, lngLength)
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
End Function
Private Sub SessionKeyFromPassword(ByVal Key As String)
Dim lngParams As Long
Dim lngReturnValue As Long
Dim strHash As String
Dim lngHashLen As Long
Dim hHash As Long
lngReturnValue = CryptCreateHash(hCryptProv, CALG_SHA, 0, 0, hHash)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a Hash Object (CryptCreateHash API)"
lngReturnValue = CryptHashData(hHash, Key, Len(Key), 0)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not calculate a Hash Value (CryptHashData API)"
lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, vbNull, lngHashLen, 0) 'get the hash length
strHash = String(lngHashLen + 1, vbNullChar)
lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, strHash, lngHashLen, 0) 'get the hash value
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not lngReturnValuerieve the hash value"
lngParams = CRYPT_EXPORTABLE
If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
lngReturnValue = CryptDeriveKey(hCryptProv, CALG_RC4, hHash, lngParams, hSessionKey)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a session key (CryptDeriveKey API)"
If hHash <> 0 Then CryptDestroyHash hHash
End Sub
Private Function Sign_KeyPair(ByVal text As String) As String
Dim hHash As Long
Dim lngReturnValue As Long
Dim lngSigLen As Long
Sign_KeyPair = vbNullString
lngReturnValue = CryptCreateHash(hCryptProv, CALG_SHA, 0, 0, hHash)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a Hash Object (CryptCreateHash API)"
lngReturnValue = CryptHashData(hHash, text, Len(text), 0)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not calculate a Hash Value (CryptHashData API)"
lngReturnValue = CryptSignHash(hHash, AT_SIGNATURE, 0, 0, vbNull, lngSigLen)
strSig = String(lngSigLen, vbNullChar)
lngReturnValue = CryptSignHash(hHash, AT_SIGNATURE, 0, 0, strSig, lngSigLen)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not sign the hash"
strSig = Left$(strSig, lngSigLen)
Sign_KeyPair = text
If hHash <> 0 Then CryptDestroyHash hHash
End Function
Private Function Validate_Keypair(ByVal text As String) As String
Dim hHash As Long
Dim lngReturnValue As Long
Dim lngSigLen As Long
Validate_Keypair = vbNullString
lngReturnValue = CryptCreateHash(hCryptProv, CALG_SHA, 0, 0, hHash)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a Hash Object (CryptCreateHash API)"
lngReturnValue = CryptHashData(hHash, text, Len(text), 0)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not calculate a Hash Value (CryptHashData API)"
lngSigLen = Len(strSig)
lngReturnValue = CryptVerifySignature(hHash, strSig, lngSigLen, hKeyPair, 0, 0)
If lngReturnValue = 0 Then
If Err.LastDllError = NTE_BAD_SIGNATURE Then
Err.Raise Err.LastDllError, , "Bad signature.  This might be because the data has changed, or the wrong public key was used to check the signature."
Else
Err.Raise Err.LastDllError, , "Could not verify the signature"
End If
End If
Validate_Keypair = text
If hHash <> 0 Then CryptDestroyHash hHash
End Function
Public Sub SignFile_KeyPair(ByVal SourceFile As String, ByVal DestFile As String)
Dim intNextFreeFile As Integer
Dim arrByteArray() As Byte
Dim X As Long
Dim y As Long
Dim z As Long
strSig = ""
intNextFreeFile = FreeFile
Open SourceFile For Binary As #intNextFreeFile
ReDim arrByteArray(0 To LOF(intNextFreeFile) - 1)
Get #intNextFreeFile, , arrByteArray()
Close #intNextFreeFile
SignByteArray_KeyPair arrByteArray()
On Error Resume Next
Kill DestFile
On Error GoTo 0
z = UBound(arrByteArray) + 1
y = 1
ReDim Preserve arrByteArray(0 To (UBound(arrByteArray) + Len(strSig) + 1))
For X = z To z + Len(strSig) - 1
arrByteArray(X) = Asc(Mid(strSig, y, 1))
y = y + 1
Next X
arrByteArray(UBound(arrByteArray)) = Len(strSig) 'length of blob is last character
intNextFreeFile = FreeFile
Open DestFile For Binary As #intNextFreeFile
Put #intNextFreeFile, , arrByteArray()
Close #intNextFreeFile
End Sub

Public Function ValidateFile_KeyPair(ByVal SourceFile As String, ByVal DestFile As String)
Dim intNextFreeFile As Integer
Dim arrByteArray() As Byte
Dim X As Long
Dim y As Long
strSig = ""
intNextFreeFile = FreeFile
Open SourceFile For Binary As #intNextFreeFile
ReDim arrByteArray(0 To LOF(intNextFreeFile) - 1)
Get #intNextFreeFile, , arrByteArray()
Close #intNextFreeFile
y = arrByteArray(UBound(arrByteArray))
For X = UBound(arrByteArray) - y To UBound(arrByteArray) - 1
strSig = strSig & Chr(arrByteArray(X))
Next X
ReDim Preserve arrByteArray(0 To (UBound(arrByteArray) - Len(strSig) - 1))
ValidateByteArray_KeyPair arrByteArray()
On Error Resume Next
If FileLen(DestFile) > 0 Then Kill DestFile
On Error GoTo 0
intNextFreeFile = FreeFile
Open DestFile For Binary As #intNextFreeFile
Put #intNextFreeFile, , arrByteArray()
Close #intNextFreeFile
End Function

Public Function SignByteArray_KeyPair(ByRef arrByteArray() As Byte)
Dim strInput As String
Dim strOutput As String
strInput = StrConv(arrByteArray(), vbUnicode)
strOutput = Sign_KeyPair(strInput)
arrByteArray() = StrConv(strOutput, vbFromUnicode) '
End Function

Public Function ValidateByteArray_KeyPair(ByRef arrByteArray() As Byte)
Dim strInput As String
Dim strOutput As String
strInput = StrConv(arrByteArray(), vbUnicode)
strOutput = Validate_Keypair(strInput)
arrByteArray() = StrConv(strOutput, vbFromUnicode)
End Function

Public Property Get ValueSignature() As String
ValueSignature = strSig
End Property

Public Property Let ValueSignature(strValue As String)
strSig = String(Len(strValue), vbNullChar)
strSig = strValue
End Property

