Attribute VB_Name = "mResource"
Option Explicit

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function GetBinaryFileFromResource
'! Description (��������)  :   [���������� ��������� ������� ��������� � ����� �� ����� ������� � ��� ID]
'! Parameters  (����������):   File_Path (String)
'                              ID (String)
'                              Resource (String)
'!--------------------------------------------------------------------------------
Public Function GetBinaryFileFromResource(ByVal File_Path As String, ByVal ID As String, ByVal Resource As String) As Boolean

    Dim iFile        As Long
    Dim BinaryData() As Byte

    iFile = FreeFile
    GetBinaryFileFromResource = False

    '�������� �� ��������
    On Error GoTo HandErr

    BinaryData = LoadResData(ID, Resource)

    If LenB(BinaryData(1)) > 0 Then
        '���� ��� - �� ����, �� ��� ���
        Open File_Path For Binary Access Write As #iFile
        '������ � ����
        Put #iFile, 1, BinaryData
        Close #iFile
        '�������� �������
        GetBinaryFileFromResource = True
    End If

ExitFromSub:

    Exit Function

HandErr:

    If Err.Number = 326 Then
        If MsgBox("Error �: " & Err.Number & vbNewLine & "Description: " & Err.Description & str2vbNewLine & "This Error in Function 'GetBinaryFileFromResource'." & vbNewLine & _
                                    "Executable file is corrupted, or required library removed from the resources of program." & str2vbNewLine & "Download the latest re-distribution program!!!" & vbNewLine & _
                                    "If the error persists, please report it to the developer." & str2vbNewLine & "Normal work of program is not guaranteed, you want to continue?", vbCritical + vbYesNo, strProductName) = vbNo Then

            End

        End If

    ElseIf Err.Number <> 0 Then
        GoTo ExitFromSub
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function ExtractResource
'! Description (��������)  :   [���������� ��������� ������� ��������� � ����� �� ����� ������� � ��� ID]
'! Parameters  (����������):   strOCXFileName (String)
'                              strPathOcx (String)
'!--------------------------------------------------------------------------------
Public Function ExtractResource(ByVal strOCXFileName As String, ByVal strPathOcx As String) As Boolean

    Dim strCopyOcxFileTo As String

    strCopyOcxFileTo = BackslashAdd2Path(strPathOcx) & strOCXFileName

    ' ��������� ������ � ����
    If GetBinaryFileFromResource(strCopyOcxFileTo, "OCX_" & FileName_woExt(strOCXFileName), "CUSTOM") Then
        DebugMode str2VbTab & strOCXFileName & ": BinaryFileFromResourse: True"
        ExtractResource = True
    Else
        DebugMode str2VbTab & strOCXFileName & ": BinaryFileFromResourse: False"
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function ExtractResourceAll
'! Description (��������)  :   [�������� ���� �������� ��������� (OCX-DLL) � �������� �������]
'! Parameters  (����������):   strPathOcxTo (String)
'!--------------------------------------------------------------------------------
Public Function ExtractResourceAll(ByVal strPathOcxTo As String) As Boolean
    DebugMode "ExtractResourceAll - Start"
    ExtractResourceAll = True

    If ExtractResource("MSFLXGRD.OCX", strPathOcxTo) = False Then
        If MsgBox("Extract OCX or DLL: 'MSFLXGRD.OCX' - False" & str2vbNewLine & strMessages(134), vbYesNo + vbQuestion, strProductName) = vbNo Then

            End

        End If

        ExtractResourceAll = False
    End If

'    DebugMode vbTab & "ExtractResourceAll - *****************Check Next File********************"
'
'    If ExtractResource("RICHTX32.OCX", strPathOcxTo) = False Then
'        If MsgBox("Extract OCX or DLL: 'RICHTX32.OCX' - False" & str2vbNewLine & strMessages(134), vbYesNo + vbQuestion, strProductName) = vbNo Then
'
'            End
'
'        End If
'
'        ExtractResourceAll = False
'    End If

    DebugMode vbTab & "ExtractResourceAll - *****************Check Next File********************"

    If ExtractResource("TABCTL32.OCX", strPathOcxTo) = False Then
        If MsgBox("Extract OCX or DLL: 'TABCTL32.OCX' - False" & str2vbNewLine & strMessages(134), vbYesNo + vbQuestion, strProductName) = vbNo Then

            End

        End If

        ExtractResourceAll = False
    End If

    DebugMode vbTab & "ExtractResourceAll - *****************Check Next File********************"

    If ExtractResource("vbscript.dll", strPathOcxTo) = False Then
        If MsgBox("Extract OCX or DLL: 'vbscript.dll' - False" & str2vbNewLine & strMessages(134), vbYesNo + vbQuestion, strProductName) = vbNo Then

            End

        End If

        ExtractResourceAll = False
    End If

'    DebugMode vbTab & "ExtractResourceAll - *****************Check Next File********************"
'
'    If ExtractResource("capicom.dll", strPathOcxTo) = False Then
'        If MsgBox("Extract OCX or DLL: capicom.dll' - False" & str2vbNewLine & strMessages(20), vbYesNo + vbQuestion, strProductName) = vbNo Then
'            End
'
'        End If
'
'        ExtractResourceAll = False
'
'    End If
    DebugMode "ExtractResourceAll - End"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ExtractrResToFolder
'! Description (��������)  :   [���������� �������� ��������� � �������]
'! Parameters  (����������):   strArg (String)
'!--------------------------------------------------------------------------------
Public Sub ExtractrResToFolder(strArg As String)

    Dim strArg_x()    As String
    Dim strPathToTemp As String
    Dim strPathTo     As String

    ' ��������� ���� �� ���������
    strPathToTemp = strArg

    ' ��������� ������������ ��������
    If LenB(strPathToTemp) > 0 Then
        If PathExists(strPathToTemp) = False Then
            CreateNewDirectory strPathToTemp
        End If

        strPathTo = BackslashAdd2Path(strPathToTemp)
    Else
        strPathTo = strWorkTemp
    End If

    ' ������ ���������� ���� (dll-ocx) �������� ��������� � �������� �������� � �������
    If ExtractResourceAll(strPathTo) Then
        If MsgBox(strMessages(135), vbYesNo + vbInformation, strProductName) = vbYes Then
            ShellEx strPathTo, essSW_SHOWNORMAL
        End If

    Else

        If MsgBox(strMessages(136), vbYesNo + vbInformation, strProductName) = vbYes Then
            ShellEx strPathTo, essSW_SHOWNORMAL
        End If
    End If

End Sub
