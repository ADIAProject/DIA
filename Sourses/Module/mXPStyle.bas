Attribute VB_Name = "mXPStyle"
Option Explicit

' ������ ��� ������������� ����� XP+ � ����������, ��������� ���� ��������� � �������� ���������
Public Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Public mbAeroEnabled As Boolean
Public mbAppThemed   As Boolean

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function IsAeroEnabled
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Function IsAeroEnabled() As Boolean

    Dim GlassState As Long

    If APIFunctionPresent("DwmIsCompositionEnabled", "dwmapi.dll") Then
        Call DwmIsCompositionEnabled(GlassState)
        IsAeroEnabled = CBool(GlassState)
    End If

End Function
