Attribute VB_Name = "mLoadImage"
Option Explicit

Public strPathImageStatusButton         As String
Public strPathImageMain                 As String

'Public strPathImageMenu                 As String
Public strPathImageStatusButtonWork     As String
Public strPathImageMainWork             As String

'Public strPathImageMenuWork             As String
Private Const lngIMG_SIZE               As Long = &H20

Public Sub LoadIconImage2Btn(ByVal ObjectName As ctlXpButton, _
                             ByVal strPictureName As String, _
                             ByVal strPathImageDir As String)
    LoadIconImageFromFileBtn ObjectName, SearchFilesInRoot(strPathImageDir, strPictureName & ".*", False, True)

End Sub

Public Sub LoadIconImage2BtnJC(ByVal ObjectName As ctlJCbutton, _
                               ByVal strPictureName As String, _
                               ByVal strPathImageDir As String)
    LoadIconImageFromFileBtnJC ObjectName, SearchFilesInRoot(strPathImageDir, strPictureName & ".ico", False, True)

End Sub

Public Sub LoadIconImage2FrameJC(ByVal ObjectName As ctlJCFrames, _
                                 ByVal strPictureName As String, _
                                 ByVal strPathImageDir As String)
    LoadIconImageFromFileJC ObjectName, SearchFilesInRoot(strPathImageDir, strPictureName & ".*", False, True)

End Sub

Public Sub LoadIconImage2Object(ObjectName As Object, _
                                strPictureName As String, _
                                strPathImageDir As String)
    LoadIconImageFromFile ObjectName, SearchFilesInRoot(strPathImageDir, strPictureName & ".*", False, True)

End Sub

Private Sub LoadIconImageFromFile(imgName As PictureBox, PicturePath As String)
    DebugMode "LoadIconImageFromFile-Start", 2
    DebugMode vbTab & "LoadIconImageFromFile: PicturePath=" & PicturePath, 2

    If PathFileExists(PicturePath) = 1 Then
        Set imgName.Picture = Nothing
        imgName.Picture = stdole.LoadPicture(PicturePath, lngIMG_SIZE, lngIMG_SIZE, Color)
        'imgName.Picture = LoadPicture(PicturePath, lngIMG_SIZE, Color)
    Else

        If Not mbSilentRun Then
            DebugMode vbTab & "LoadIconImageFromFile: Path to picture: " & PicturePath & " not Exist. Standard picture Will is used", 2

        End If

    End If

    DebugMode "LoadIconImageFromFile-End", 2

End Sub

Private Sub LoadIconImageFromFileBtn(ByVal imgName As ctlXpButton, _
                                     ByVal PicturePath As String)
    DebugMode vbTab & "LoadIconImageFromFileBtn-Start", 2
    DebugMode str2VbTab & "LoadIconImageFromFileBtn: PicturePath=" & PicturePath, 2

    If PathFileExists(PicturePath) = 1 Then

        With imgName

            If Not (.Picture Is Nothing) Then
                If .Picture <> stdole.LoadPicture(PicturePath) Then
                    Set .Picture = Nothing
                    Set .Picture = stdole.LoadPicture(PicturePath)
                    DebugMode str2VbTab & "LoadIconImageFromFileBtn: Picture is Installed", 2
                Else
                    DebugMode str2VbTab & "LoadIconImageFromFileBtn: Picture is already set", 2

                End If

            Else
                Set .Picture = Nothing
                Set .Picture = stdole.LoadPicture(PicturePath)
                DebugMode str2VbTab & "LoadIconImageFromFileBtn: Picture is Installed", 2

            End If

        End With

        'imgName
    Else

        If Not mbSilentRun Then
            DebugMode str2VbTab & "LoadIconImageFromFileBtn: Path to picture: " & PicturePath & " not Exist. Standard picture Will is used", 2

        End If

    End If

    DebugMode vbTab & "LoadIconImageFromFileBtn-End", 2

End Sub

Private Sub LoadIconImageFromFileBtnJC(ByVal btnName As ctlJCbutton, _
                                       ByVal PicturePath As String)
    DebugMode vbTab & "LoadIconImageFromFileBtnJC-Start", 2
    DebugMode str2VbTab & "LoadIconImageFromFileBtnJC: PicturePath=" & PicturePath, 2

    If PathFileExists(PicturePath) = 1 Then

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

        'imgName
    Else
        DebugMode str2VbTab & "Path to picture: " & PicturePath & " not Exist. Standard picture Will is used", 2

    End If

    DebugMode vbTab & "LoadIconImageFromFileBtnJC-End", 2
ExitFromSub:
    Exit Sub
PictureNotAllowFormat:

    If err.Number = 481 Then
        MsgBox "Error ¹: " & err.Number & vbNewLine & "Description: " & err.Description & str2vbNewLine & "This Error in Function 'CreateRestorePoint'. Probably trouble with WMI.", vbCritical, strProductName
    ElseIf err.Number <> 0 Then
        GoTo ExitFromSub

    End If

End Sub

Private Sub LoadIconImageFromFileJC(ByVal imgName As ctlJCFrames, _
                                    ByVal PicturePath As String)
    DebugMode vbTab & "LoadIconImageFromFileJC-Start", 2
    DebugMode str2VbTab & "LoadIconImageFromFileJC: PicturePath=" & PicturePath, 2

    If PathFileExists(PicturePath) = 1 Then
        'imgName.Picture = stdole.LoadPicture(PicturePath, lngIMG_SIZE, lngIMG_SIZE, Color)
        Set imgName.Picture = Nothing
        Set imgName.Picture = stdole.LoadPicture(PicturePath)
    Else
        DebugMode str2VbTab & "Path to picture: " & PicturePath & " not Exist. Standard picture Will is used"

    End If

    DebugMode vbTab & "LoadIconImageFromFileJC-End", 2

End Sub

Public Function LoadIconImageFromPath(strPictureName As String, _
                                      strPathImageDir As String) As IPictureDisp

Dim strPicturePath                      As String

    DebugMode "LoadIconImageFromPath-Start", 2
    strPicturePath = SearchFilesInRoot(strPathImageDir, strPictureName & ".*", False, True)
    DebugMode vbTab & "LoadIconImageFromPath: PicturePath=" & strPicturePath, 2

    If PathFileExists(strPicturePath) = 1 Then
        'Set LoadIconImageFromPath = LoadPicture(strPicturePath)
        Set LoadIconImageFromPath = stdole.LoadPicture(strPicturePath)
    Else

        If Not mbSilentRun Then
            DebugMode vbTab & "LoadIconImageFromPath: Path to picture: " & strPicturePath & " not Exist. Standard picture Will is used", 2

        End If

    End If

    DebugMode "LoadIconImageFromPath-End", 2

End Function

Public Sub LoadIconImagePath()
    DebugMode vbTab & "LoadIconImagePath-Start", 2
    strPathImageMainWork = strPathImageMain & strImageMainName
    strPathImageStatusButtonWork = strPathImageStatusButton & strImageStatusButtonName

    'strPathImageMenuWork = strPathImageMenu & strImageMenuName
    If PathFileExists(strPathImageMainWork) = 0 Then
        If Not mbSilentRun Then
            MsgBox strMessages(15), vbCritical, strProductName

        End If

        strPathImageMainWork = strPathImageMain & "Standart"

    End If

    If PathFileExists(strPathImageStatusButtonWork) = 0 Then
        If Not mbSilentRun Then
            MsgBox strMessages(15), vbCritical, strProductName

        End If

        strPathImageStatusButtonWork = strPathImageStatusButton & "Standart"

    End If

    'If PathFileExists(strPathImageMenuWork) = 0 Then
    'If Not mbSilentRun Then
    'MsgBox strMessages(15), vbCritical, strProductName
    'End If
    'strPathImageMenuWork = strPathImageMenu & "Standart"
    'End If
    DebugMode vbTab & "LoadIconImagePath-End", 2

End Sub
