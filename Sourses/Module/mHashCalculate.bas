Attribute VB_Name = "mHashCalculate"
Option Explicit

#Const mbIDE_DBSProject = False

' Not add to project (if not DBS) - option for compile
#If mbIDE_DBSProject Then
    Public Enum lngHASH_TYPE
        CAPICOM_HASH_ALGORITHM_SHA1 = 0
        CAPICOM_HASH_ALGORITHM_MD2 = 1
        CAPICOM_HASH_ALGORITHM_MD4 = 2
        CAPICOM_HASH_ALGORITHM_MD5 = 3
        CAPICOM_HASH_ALGORITHM_SHA_256 = 4
        CAPICOM_HASH_ALGORITHM_SHA_384 = 5
        CAPICOM_HASH_ALGORITHM_SHA_512 = 6
    End Enum
#End If

' CAPICOM 2.1.0.2 (http://support.microsoft.com/kb/931906/)
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CalcHashFile
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'                              lngHashAlgoritm (lngHASH_TYPE)
'!--------------------------------------------------------------------------------
#If mbIDE_DBSProject Then
    Public Function CalcHashFile(ByVal strPathFile As String, ByVal lngHashAlgoritm As lngHASH_TYPE) As String
    
        Dim objHashedData As CAPICOM.HashedData
        Dim objStream     As ADODB.Stream
    
        Set objHashedData = New CAPICOM.HashedData
        Set objStream = New ADODB.Stream
        objHashedData.Algorithm = lngHashAlgoritm
    
        'Для строки
        '.Hash UStr2BStr(strText)
        If PathExists(strPathFile) Then
    
            With objStream
                .Type = adTypeBinary
                .Open
                .LoadFromFile (strPathFile)
    
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
#End If
