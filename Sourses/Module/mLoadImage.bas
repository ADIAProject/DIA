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
'! Procedure   (�������)   :   Sub LoadIconImage2Object
'! Description (��������)  :   [Set Picture to Object]
'! Parameters  (����������):   ObjectName (Object)
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
            LoadIconImageFromFileBtnJC objName, strFilePicture
        ElseIf TypeOf objName Is ctlJCFrames Then
            LoadIconImageFromFileJC objName, strFilePicture
        Else
            LoadIconImageFromFile objName, strFilePicture
        End If
    End If
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadIconImageFromFile
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   imgName (PictureBox)
'                              PicturePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadIconImageFromFile(imgName As PictureBox, PicturePath As String)
    If mbDebugDetail Then DebugMode vbTab & "LoadIconImageFromFile: PicturePath=" & PicturePath

    If PathExists(PicturePath) Then
        Set imgName.Picture = Nothing
        imgName.Picture = stdole.LoadPicture(PicturePath, lngIMG_SIZE, lngIMG_SIZE, Color)
    Else

        If Not mbSilentRun Then
            If mbDebugDetail Then DebugMode vbTab & "LoadIconImageFromFile: Path to picture: " & PicturePath & " not Exist. Standard picture Will is used"
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadIconImageFromFileBtnJC
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   btnName (ctlJCbutton)
'                              PicturePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadIconImageFromFileBtnJC(ByVal btnName As ctlJCbutton, ByVal PicturePath As String)
    If mbDebugDetail Then DebugMode str2VbTab & "LoadIconImageFromFileBtnJC: PicturePath=" & PicturePath

    If PathExists(PicturePath) Then

        On Error GoTo ExitFromSub

        With btnName

            If Not (.PictureNormal Is Nothing) Then
                If .PictureNormal <> stdole.LoadPicture(PicturePath) Then
                    Set .PictureNormal = Nothing
                    Set .PictureNormal = stdole.LoadPicture(PicturePath)
                    If mbDebugDetail Then DebugMode str2VbTab & "LoadIconImageFromFileBtnJC: Picture is Installed"
                Else
                    If mbDebugDetail Then DebugMode str2VbTab & "LoadIconImageFromFileBtnJC: Picture is already set"
                End If

            Else
                Set .PictureNormal = Nothing
                Set .PictureNormal = stdole.LoadPicture(PicturePath)
                If mbDebugDetail Then DebugMode str2VbTab & "LoadIconImageFromFileBtnJC: Picture is Installed"
            End If

        End With

    Else
        If mbDebugDetail Then DebugMode str2VbTab & "Path to picture: " & PicturePath & " not Exist. Standard picture Will is used"
    End If


ExitFromSub:

    Exit Sub

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadIconImageFromFileJC
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   imgName (ctlJCFrames)
'                              PicturePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadIconImageFromFileJC(ByVal imgName As ctlJCFrames, ByVal PicturePath As String)
    If mbDebugDetail Then DebugMode str2VbTab & "LoadIconImageFromFileJC: PicturePath=" & PicturePath

    If PathExists(PicturePath) Then
        'imgName.Picture = stdole.LoadPicture(PicturePath, lngIMG_SIZE, lngIMG_SIZE, Color)
        Set imgName.Picture = Nothing
        Set imgName.Picture = stdole.LoadPicture(PicturePath)
    Else
        If mbDebugStandart Then DebugMode str2VbTab & "LoadIconImageFromFileJC-Path to picture: " & PicturePath & " not Exist. Standard picture Will is used"
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function LoadIconImageFromPath
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strPictureName (String)
'                              strPathImageDir (String)
'!--------------------------------------------------------------------------------
Public Function LoadIconImageFromPath(strPictureName As String, strPathImageDir As String) As IPictureDisp

    Dim strFile()       As FindListStruct
    Dim strFilePicture  As String
    
    strFile = SearchFilesInRoot(strPathImageDir, strPictureName & ".*", False, True)
    strFilePicture = strFile(0).FullPath
    
    If mbDebugDetail Then DebugMode vbTab & "LoadIconImageFromPath: PicturePath=" & strFilePicture
    
    If PathExists(strFilePicture) Then
        Set LoadIconImageFromPath = stdole.LoadPicture(strFilePicture)
    Else

        If Not mbSilentRun Then
            If mbDebugDetail Then DebugMode vbTab & "LoadIconImageFromPath: Path to picture: " & strFilePicture & " not Exist. Standard picture Will is used"
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadIconImagePath
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub LoadIconImagePath()
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
