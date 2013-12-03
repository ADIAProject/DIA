Attribute VB_Name = "mHashCalculate"
Option Explicit

Public Enum lngHASH_TYPE
    CAPICOM_HASH_ALGORITHM_SHA1 = 0
    CAPICOM_HASH_ALGORITHM_MD2 = 1
    CAPICOM_HASH_ALGORITHM_MD4 = 2
    CAPICOM_HASH_ALGORITHM_MD5 = 3
    CAPICOM_HASH_ALGORITHM_SHA_256 = 4
    CAPICOM_HASH_ALGORITHM_SHA_384 = 5
    CAPICOM_HASH_ALGORITHM_SHA_512 = 6

End Enum

' CAPICOM 2.1.0.2 (http://support.microsoft.com/kb/931906/)
Public Function CalcHashFile(ByVal StrPathFile As String, _
                             ByVal lngHashAlgoritm As lngHASH_TYPE) As String

Dim objHashedData                       As New CAPICOM.HashedData
Dim objStream                           As New ADODB.Stream

    objHashedData.Algorithm = lngHashAlgoritm

    'Для строки
    '.Hash UStr2BStr(strText)
    If PathFileExists(StrPathFile) = 1 Then

        With objStream
            .Type = adTypeBinary
            .Open
            .LoadFromFile (StrPathFile)

            Do Until .EOS
                objHashedData.Hash .Read
            Loop
            .Close

        End With

        CalcHashFile = objHashedData.Value

    End If

    Set objStream = Nothing
    Set objHashedData = Nothing

End Function
