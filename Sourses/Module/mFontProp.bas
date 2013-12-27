Attribute VB_Name = "mFontProp"
Option Explicit

'����������, �������� �������� ��������� ������, ����������� �� ��������� �����
Public lngFont_Charset       As Long

' ����� �������� ����� � ������ ���������
Public strFontMainForm_Name  As String
Public lngFontMainForm_Size  As Long
' ����� ������ ����
Public strFontOtherForm_Name As String
Public lngFontOtherForm_Size As Long

'����������, �������� ���������� �������� ������ ��� ������ ������� ���������
Public lngFontBtn_Color      As Long
Public strFontBtn_Name       As String
Public miFontBtn_Size        As Integer
Public mbFontBtn_Italic      As Boolean
Public mbFontBtn_Underline   As Boolean
Public mbFontBtn_Strikethru  As Boolean
Public mbFontBtn_Bold        As Boolean

'����������, �������� ���������� �������� ������ ssTAB1
Public lngFontTab_Color      As Long
Public strFontTab_Name       As String
Public miFontTab_Size        As Integer
Public mbFontTab_Italic      As Boolean
Public mbFontTab_Underline   As Boolean
Public mbFontTab_Strikethru  As Boolean
Public mbFontTab_Bold        As Boolean

'����������, �������� ���������� �������� ������ ssTAB2
Public lngFontTab2_Color     As Long
Public strFontTab2_Name      As String
Public miFontTab2_Size       As Integer
Public mbFontTab2_Italic     As Boolean
Public mbFontTab2_Underline  As Boolean
Public mbFontTab2_Strikethru As Boolean
Public mbFontTab2_Bold       As Boolean

'����������, �������� ���������� �������� ������ ��� ����������� ��������� (ToolTip)
Public lngFontTT_Color       As Long
Public strFontTT_Name        As String
Public miFontTT_Size         As Integer
Public mbFontTT_Italic       As Boolean
Public mbFontTT_Underline    As Boolean
Public mbFontTT_Strikethru   As Boolean
Public mbFontTT_Bold         As Boolean

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub SetBtnFontProperties
'! Description (��������)  :   [��������� ������� ������ ��� ������� (������)]
'! Parameters  (����������):   ctlObject (Object)
'!--------------------------------------------------------------------------------
Public Sub SetBtnFontProperties(ctlObject As Object)

    With ctlObject
        .Font.Name = strFontBtn_Name
        .Font.Size = miFontBtn_Size
        .Font.Underline = mbFontBtn_Underline
        .Font.Strikethrough = mbFontBtn_Strikethru
        .Font.Bold = mbFontBtn_Bold
        .Font.Italic = mbFontBtn_Italic
        .Font.Charset = lngFont_Charset
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub SetTTFontProperties
'! Description (��������)  :   [��������� ������� ������ ��� ������� (ToolTip)]
'! Parameters  (����������):   ctlObject (Object)
'!--------------------------------------------------------------------------------
Public Sub SetTTFontProperties(ctlObject As Object)

    With ctlObject
        .Font.Name = strFontTT_Name
        .Font.Size = miFontTT_Size
        .Font.Underline = mbFontTT_Underline
        .Font.Strikethrough = mbFontTT_Strikethru
        .Font.Bold = mbFontTT_Bold
        .Font.Italic = mbFontTT_Italic
        .Font.Charset = lngFont_Charset
    End With

End Sub


