Attribute VB_Name = "mGlobalVar"
Option Explicit

Public strAppPath                 As String
Public strAppPathBackSL           As String
Public mbInitXPStyle              As Boolean

' ���������� ��� ����� ������ ������������� ���������
Public lngShowMessageResult       As Long
Public hc_Handle_Hand             As Long       'The hand cursor handle is used by the jcButton control as well, so it is declared publicly.

'Maximum width (in pixels) for custom-built tooltips
Public Const PD_MAX_TOOLTIP_WIDTH As Long = 400

' ��������� ��������� ���������� ���������� ����� ���������
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub GetCurAppPath
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub GetCurAppPath()
    strAppPath = App.Path
    strAppPathBackSL = BackslashAdd2Path(strAppPath)
End Sub
