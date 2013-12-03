Attribute VB_Name = "mMD5"
'This module is used to gather the contents of a file quickly and to grab the MD5 of a file quickly by using API functions. Use this
'code in any projects you wish, no need to give credit. Please vote though.
'marcin@malwarebytes.org if you have any questions.
'Special thanks to Hossein Moradi for the optimizations with CryptHashData()
Option Explicit

Private Const OPEN_EXISTING             As Long = 3
Private Const GENERIC_READ              As Long = &H80000000
Private Const FILE_SHARE_READ           As Long = &H1
Private Const PROV_RSA_FULL             As Long = 1
Private Const CRYPT_VERIFYCONTEXT       As Long = &HF0000000
Private Const HP_HASHVAL                As Long = 2
Private Const CALG_MD5                  As Long = 32771
Private Const lMD5Length                As Long = 16

Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptReleaseContext _
                          Lib "advapi32.dll" (ByVal hProv As Long, _
                                              ByVal dwFlags As Long) As Long

Private Declare Function CryptAcquireContext _
                          Lib "advapi32.dll" _
                              Alias "CryptAcquireContextA" (ByRef phProv As Long, _
                                                            ByVal pszContainer As String, _
                                                            ByVal pszProvider As String, _
                                                            ByVal dwProvType As Long, _
                                                            ByVal dwFlags As Long) As Long

Private Declare Function CryptCreateHash _
                          Lib "advapi32.dll" (ByVal hProv As Long, _
                                              ByVal Algid As Long, _
                                              ByVal hkey As Long, _
                                              ByVal dwFlags As Long, _
                                              ByRef phHash As Long) As Long

Private Declare Function CryptHashData _
                          Lib "advapi32.dll" (ByVal hHash As Long, _
                                              pbData As Any, _
                                              ByVal dwDataLen As Long, _
                                              ByVal dwFlags As Long) As Long

Private Declare Function CryptGetHashParam _
                          Lib "advapi32.dll" (ByVal pCryptHash As Long, _
                                              ByVal dwParam As Long, _
                                              ByRef pbData As Any, _
                                              ByRef pcbData As Long, _
                                              ByVal dwFlags As Long) As Long

'׀אסקוע ץ‎ר-סףללû MD5 פאיכא
Public Function GetMD5(sFile As String) As String

Dim hFile                               As Long
Dim uBuffer()                           As Byte
Dim lFileSize                           As Long
Dim lBytesRead                          As Long
Dim uMD5(lMD5Length)                    As Byte
Dim i                                   As Long
Dim hCrypt                              As Long
Dim hHash                               As Long
Dim sMD5                                As String

    'Get a handle to the file
    hFile = CreateFile(StrPtr(sFile & vbNullChar), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)

    'Check if file opened successfully
    If hFile > 0 Then
        'Get the file size
        lFileSize = GetFileSize(hFile, ByVal 0&)

        'File size must be greater than 0
        If lFileSize > 0 Then
            'Prepare the buffer
            ReDim uBuffer(lFileSize - 1) As Byte

            'Read the file
            If ReadFile(hFile, uBuffer(0), lFileSize, lBytesRead, ByVal 0&) <> 0 Then
                If lBytesRead <> lFileSize Then
                    ReDim Preserve uBuffer(lBytesRead - 1)

                End If

                'Acquire the context, create the hash, and hash the data
                If CryptAcquireContext(hCrypt, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
                    If CryptCreateHash(hCrypt, CALG_MD5, 0&, 0&, hHash) <> 0 Then
                        If CryptHashData(hHash, uBuffer(0), lBytesRead, ByVal 0&) <> 0 Then
                            If CryptGetHashParam(hHash, HP_HASHVAL, uMD5(0), lMD5Length, 0) <> 0 Then

                                'Build the MD5 string
                                For i = 0 To lMD5Length - 1
                                    sMD5 = sMD5 & (Right$("0" & Hex$(uMD5(i)), 2))
                                Next

                            End If

                        End If

                        'Destroy the hash
                        CryptDestroyHash hHash

                    End If

                    'Release the context
                    CryptReleaseContext hCrypt, 0

                End If

            End If

        End If

        'Close the handle to the file
        CloseHandle hFile

    End If

    'Convert to lower case
    GetMD5 = LCase$(sMD5)

End Function
