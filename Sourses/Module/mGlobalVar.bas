Attribute VB_Name = "mGlobalVar"
Option Explicit

Public strAppPath                 As String
Public strAppPathBackSL           As String

' ���������� ��� ����� ������ ������������� ���������
Public lngShowMessageResult       As Long

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub GetCurAppPath
'! Description (��������)  :   [��������� ��������� ���������� ���������� ����� ���������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub GetCurAppPath()
    strAppPath = App.Path
    strAppPathBackSL = BackslashAdd2Path(strAppPath)
End Sub
