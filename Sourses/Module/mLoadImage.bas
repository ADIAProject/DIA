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
Dim strFile() As FindFileListStruct
Dim strFilePicture As String
    
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
'! Procedure   (Функция)   :   Sub LoadIconImageFromFile
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   imgName (PictureBox)
'                              PicturePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadIconImageFromFile(imgName As PictureBox, PicturePath As String)
    DebugMode "LoadIconImageFromFile-Start" & vbNewLine & _
              vbTab & "LoadIconImageFromFile: PicturePath=" & PicturePath, 2

    If PathExists(PicturePath) Then
        Set imgName.Picture = Nothing
        imgName.Picture = stdole.LoadPicture(PicturePath, lngIMG_SIZE, lngIMG_SIZE, Color)
    Else

        If Not mbSilentRun Then
            DebugMode vbTab & "LoadIconImageFromFile: Path to picture: " & PicturePath & " not Exist. Standard picture Will is used", 2
        End If
    End If

    DebugMode "LoadIconImageFromFile-End", 2
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadIconImageFromFileBtnJC
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   btnName (ctlJCbutton)
'                              PicturePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadIconImageFromFileBtnJC(ByVal btnName As ctlJCbutton, ByVal PicturePath As String)
    DebugMode vbTab & "LoadIconImageFromFileBtnJC-Start" & vbNewLine & _
              str2VbTab & "LoadIconImageFromFileBtnJC: PicturePath=" & PicturePath, 2

    If PathExists(PicturePath) Then

        On Error GoTo PictureNotAllowFormat

        With btnName

            If Not (.PictureNormal Is Nothing) Then
                If .PictureNormal <> stdole.LoadPicture(PicturePath) Then
                    Set .PictureNormal = Nothing
                    Set .PictureNormal = stdole.LoadPicture(PicturePath)
                    DebugMode str2VbTab & "LoadIconImageFromFileBtnJC: Picture is Installed", 2
                Else
                    DebugMode str2VbTab & "LoadIconImageFromFileBtnJC: Picture is already set", 2
                End If

            Else
                Set .PictureNormal = Nothing
                Set .PictureNormal = stdole.LoadPicture(PicturePath)
                DebugMode str2VbTab & "LoadIconImageFromFileBtnJC: Picture is Installed", 2
            End If

        End With

    Else
        DebugMode str2VbTab & "Path to picture: " & PicturePath & " not Exist. Standard picture Will is used", 2
    End If

    DebugMode vbTab & "LoadIconImageFromFileBtnJC-End", 2
ExitFromSub:

    Exit Sub

PictureNotAllowFormat:

    If Err.Number = 481 Then
        MsgBox "Error №: " & Err.Number & vbNewLine & "Description: " & Err.Description & str2vbNewLine & "This Error in Function 'CreateRestorePoint'. Probably trouble with WMI.", vbCritical, strProductName
    ElseIf Err.Number <> 0 Then
        GoTo ExitFromSub
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadIconImageFromFileJC
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   imgName (ctlJCFrames)
'                              PicturePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadIconImageFromFileJC(ByVal imgName As ctlJCFrames, ByVal PicturePath As String)
    DebugMode vbTab & "LoadIconImageFromFileJC-Start" & vbNewLine & _
              str2VbTab & "LoadIconImageFromFileJC: PicturePath=" & PicturePath, 2

    If PathExists(PicturePath) Then
        'imgName.Picture = stdole.LoadPicture(PicturePath, lngIMG_SIZE, lngIMG_SIZE, Color)
        Set imgName.Picture = Nothing
        Set imgName.Picture = stdole.LoadPicture(PicturePath)
    Else
        DebugMode str2VbTab & "Path to picture: " & PicturePath & " not Exist. Standard picture Will is used"
    End If

    DebugMode vbTab & "LoadIconImageFromFileJC-End", 2
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function LoadIconImageFromPath
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPictureName (String)
'                              strPathImageDir (String)
'!--------------------------------------------------------------------------------
Public Function LoadIconImageFromPath(strPictureName As String, strPathImageDir As String) As IPictureDisp

    Dim strFile() As FindFileListStruct
    Dim strFilePicture As String
    
    DebugMode "LoadIconImageFromPath-Start", 2
    
    strFile = SearchFilesInRoot(strPathImageDir, strPictureName & ".*", False, True)
    strFilePicture = strFile(0).FullPath
    
    DebugMode vbTab & "LoadIconImageFromPath: PicturePath=" & strFilePicture, 2
    
    If PathExists(strFilePicture) Then
        Set LoadIconImageFromPath = stdole.LoadPicture(strFilePicture)
    Else

        If Not mbSilentRun Then
            DebugMode vbTab & "LoadIconImageFromPath: Path to picture: " & strFilePicture & " not Exist. Standard picture Will is used", 2
        End If
    End If

    DebugMode "LoadIconImageFromPath-End", 2
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadIconImagePath
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub LoadIconImagePath()
    DebugMode vbTab & "LoadIconImagePath-Start", 2
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
    DebugMode vbTab & "LoadIconImagePath-End", 2
End Sub
