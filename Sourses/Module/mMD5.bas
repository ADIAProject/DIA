Attribute VB_Name = "mMD5"
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=69092&lngWId=1
'This module is used to gather the contents of a file quickly and to grab the MD5 of a file quickly by using API functions. Use this
'code in any projects you wish, no need to give credit. Please vote though.
'marcin@malwarebytes.org if you have any questions.
'Special thanks to Hossein Moradi for the optimizations with CryptHashData()
Option Explicit

Private Const OPEN_EXISTING       As Long = 3
Private Const GENERIC_READ        As Long = &H80000000
Private Const FILE_SHARE_READ     As Long = &H1
Private Const PROV_RSA_FULL       As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const HP_HASHVAL          As Long = 2
Private Const CALG_MD5            As Long = 32771
Private Const lMD5Length          As Long = 16

Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hkey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal pCryptHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pcbData As Long, ByVal dwFlags As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetMD5
'! Description (Описание)  :   [Расчет хэш-суммы MD5 файла]
'! Parameters  (Переменные):   sFile (String)
'!--------------------------------------------------------------------------------
Public Function GetMD5(sFile As String) As String

    Dim hFile            As Long
    Dim uBuffer()        As Byte
    Dim lFileSize        As Long
    Dim lBytesRead       As Long
    Dim uMD5(lMD5Length) As Byte
    Dim ii               As Long
    Dim hCrypt           As Long
    Dim hHash            As Long
    Dim sMD5             As String
    Dim lngFilePathPtr   As Long
    
    'Get a pointer to a string with file name.
    If PathIsValidUNC(sFile) = False Then
        lngFilePathPtr = StrPtr("\\?\" & sFile)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(sFile, Len(sFile) - 2))
    End If
    'Get a handle to the file
    hFile = CreateFile(lngFilePathPtr, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, ByVal 0&)

    'Check if file opened successfully
    If hFile Then
        'Get the file size
        lFileSize = GetFileSize(hFile, ByVal 0&)

        'File size must be greater than 0
        If lFileSize Then

            'Prepare the buffer
            ReDim uBuffer(lFileSize - 1)

            'Read the file
            If ReadFile(hFile, VarPtr(uBuffer(0)), lFileSize, lBytesRead, 0) <> 0 Then
                If lBytesRead <> lFileSize Then

                    ReDim Preserve uBuffer(lBytesRead - 1)

                End If

                'Acquire the context, create the hash, and hash the data
                If CryptAcquireContext(hCrypt, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
                    If CryptCreateHash(hCrypt, CALG_MD5, 0&, 0&, hHash) <> 0 Then
                        If CryptHashData(hHash, uBuffer(0), lBytesRead, ByVal 0&) <> 0 Then
                            If CryptGetHashParam(hHash, HP_HASHVAL, uMD5(0), lMD5Length, 0) <> 0 Then

                                'Build the MD5 string
                                For ii = 0 To lMD5Length - 1
                                    sMD5 = sMD5 & (Right$("0" & Hex$(uMD5(ii)), 2))
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
