Attribute VB_Name = "mGlobalVar"
Option Explicit

Public strAppPath                 As String
Public strAppPathBackSL           As String

' ���������� ��� ����� ������ ������������� ���������
Public lngShowMessageResult       As Long

'Per the excellent advice of Kroc (camendesign.com), a custom UserMode variable is less prone to errors than the usual
' Ambient.UserMode value supplied to ActiveX controls.  This fixes a problem where ActiveX controls sometimes think they
' are being run in a compiled EXE, when actually their properties are just being written as part of .exe compiling.
Public g_UserModeFix As Boolean

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub GetCurAppPath
'! Description (��������)  :   [��������� ��������� ���������� ���������� ����� ���������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub GetCurAppPath()
    strAppPath = App.Path
    strAppPathBackSL = BackslashAdd2Path(strAppPath)
End Sub
