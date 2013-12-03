Attribute VB_Name = "mPrevInstance"
'��������� ����� ���������� ���� �� ���������
Option Explicit

Private Declare Function OpenIcon Lib "user32.dll" (ByVal hWnd As Long) As Long

'! -----------------------------------------------------------
'!  �������     :  ShowPrevInstance
'!  ����������  :
'!  ��������    :  ���������� ���������� ����� ���������, ���� ��������� �������� ������
'! -----------------------------------------------------------
Public Sub ShowPrevInstance()

Dim OldTitle                            As String
Dim hWnd                                As Long
Const SW_RESTORE                        As Long = 9

    OldTitle = App.Title
    App.Title = "This App Will Be Closed"
    hWnd = FindWindow(StrPtr("ThunderRT6FormDC"), StrPtr(OldTitle))

    If hWnd <> 0 Then
        ShowWindow hWnd, SW_RESTORE
        OpenIcon hWnd
        SetForegroundWindow hWnd
        AppActivate OldTitle
        End
    End If

End Sub

