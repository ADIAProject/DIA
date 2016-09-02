Attribute VB_Name = "mLoadImage"
Option Explicit

Public strPathImageStatusButton     As String
Public strPathImageMain             As String
'Public strPathImageMenu             As String

Public strPathImageStatusButtonWork As String
Public strPathImageMainWork         As String
'Public strPathImageMenuWork         As String

Private Const lngIMG_SIZE           As Long = &H20

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadIconImage2Object
'! Description (Описание)  :   [Set Picture to Object]
'! Parameters  (Переменные):   ObjectName (Object)
'                              strPictureName (String)
'                              strPathImageDir (String)
'!--------------------------------------------------------------------------------
Public Sub LoadIconImage2Object(objName As Object, strPictureName As String, strPathImageDir As String)
Dim strFile()       As FindListStruct
Dim strFilePicture  As String
    
    strFile = SearchFilesInRoot(strPathImageDir, strPictureName & ".ico", False, True)
    strFilePicture = strFile(0).FullPath
    
    If LenB(strFilePicture) Then
        If TypeOf objName Is ctlJCbutton Then
            LoadImageFromFile2JCbutton objName, strFilePicture
        ElseIf TypeOf objName Is ctlJCFrames Then
            LoadImageFromFile2JCFrames objName, strFilePicture
        Else
            LoadImageFromFile2PictureBox objName, strFilePicture
        End If
    End If
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadImageFromFile2PictureBox
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   objName (PictureBox)
'                              PicturePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadImageFromFile2PictureBox(objName As PictureBox, PicturePath As String)
    If mbDebugDetail Then DebugMode vbTab & "LoadImageFromFile2PictureBox: PicturePath=" & PicturePath

    If FileExists(PicturePath) Then
        Set objName.Picture = Nothing
        objName.Picture = StdPictureEx.LoadPicture(PicturePath) ', lpsCustom, , lngIMG_SIZE, lngIMG_SIZE)
    Else

        If Not mbSilentRun Then
            If mbDebugDetail Then DebugMode vbTab & "LoadImageFromFile2PictureBox: Path to picture: " & PicturePath & " not Exist. Standard picture Will is used"
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadImageFromFile2JCbutton
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   btnName (ctlJCbutton)
'                              strPicturePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadImageFromFile2JCbutton(ByVal btnName As ctlJCbutton, ByVal strPicturePath As String)
    Dim objPictTmp As StdPicture
    
    If mbDebugDetail Then DebugMode str2VbTab & "LoadImageFromFile2JCbutton: PicturePath=" & strPicturePath

    If FileExists(strPicturePath) Then

        On Error GoTo ExitFromSub

        With btnName
            Set .PictureNormal = Nothing
            Set objPictTmp = StdPictureEx.LoadPicture(strPicturePath)
            
            If Not (.PictureNormal Is Nothing) Then
                
                If .PictureNormal <> objPictTmp Then
                    Set .PictureNormal = objPictTmp
                    If mbDebugDetail Then DebugMode str2VbTab & "LoadImageFromFile2JCbutton: Picture is Installed"
                Else
                    If mbDebugDetail Then DebugMode str2VbTab & "LoadImageFromFile2JCbutton: Picture is already set"
                End If

            Else
                
                Set .PictureNormal = objPictTmp
                If mbDebugDetail Then DebugMode str2VbTab & "LoadImageFromFile2JCbutton: Picture is Installed"
            End If

        End With

    Else
        If mbDebugDetail Then DebugMode str2VbTab & "Path to picture: " & strPicturePath & " not Exist. Standard picture Will is used"
    End If

ExitFromSub:

    Set objPictTmp = Nothing
    Exit Sub

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadImageFromFile2JCFrames
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   objName (ctlJCFrames)
'                              PicturePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadImageFromFile2JCFrames(ByVal objName As ctlJCFrames, ByVal PicturePath As String)
    If mbDebugDetail Then DebugMode str2VbTab & "LoadImageFromFile2JCFrames: PicturePath=" & PicturePath

    If FileExists(PicturePath) Then
        Set objName.Picture = Nothing
        Set objName.Picture = StdPictureEx.LoadPicture(PicturePath)
    Else
        If mbDebugStandart Then DebugMode str2VbTab & "LoadImageFromFile2JCFrames-Path to picture: " & PicturePath & " not Exist. Standard picture Will is used"
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetImageFromFile
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPictureName (String)
'                              strPathImageDir (String)
'!--------------------------------------------------------------------------------
Public Function GetImageFromFile(strPictureName As String, strPathImageDir As String) As IPictureDisp

    Dim strFile()       As FindListStruct
    Dim strFilePicture  As String
    
    strFile = SearchFilesInRoot(strPathImageDir, strPictureName & ".*", False, True)
    strFilePicture = strFile(0).FullPath
    
    If mbDebugDetail Then DebugMode vbTab & "GetImageFromFile: PicturePath=" & strFilePicture
    
    If FileExists(strFilePicture) Then
        Set GetImageFromFile = StdPictureEx.LoadPicture(strFilePicture)
    Else

        If Not mbSilentRun Then
            If mbDebugDetail Then DebugMode vbTab & "GetImageFromFile: Path to picture: " & strFilePicture & " not Exist. Standard picture Will is used"
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub GetImageSkinPath
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub GetImageSkinPath()
    strPathImageMainWork = strPathImageMain & strImageMainName
    strPathImageStatusButtonWork = strPathImageStatusButton & strImageStatusButtonName
    'strPathImageMenuWork = strPathImageMenu & strImageMenuName
    
    If PathExists(strPathImageMainWork) = False Then
        If Not mbSilentRun Then
            MsgBox strMessages(15), vbCritical, strProductName
        End If

        strPathImageMainWork = strPathImageMain & "Standart"
    End If

    If PathExists(strPathImageStatusButtonWork) = False Then
        If Not mbSilentRun Then
            MsgBox strMessages(15), vbCritical, strProductName
        End If

        strPathImageStatusButtonWork = strPathImageStatusButton & "Standart"
    End If

    'If PathExists(strPathImageMenuWork) = False Then
    'If Not mbSilentRun Then
    'MsgBox strMessages(15), vbCritical, strProductName
    'End If
    'strPathImageMenuWork = strPathImageMenu & "Standart"
    'End If
End Sub
