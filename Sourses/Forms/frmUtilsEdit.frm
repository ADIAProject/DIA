VERSION 5.00
Begin VB.Form frmUtilsEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������������� ������"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   7665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUtilsEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.TextBoxW txtParamUtil 
      Height          =   330
      Left            =   2640
      TabIndex        =   2
      Top             =   1620
      Width           =   4935
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "frmUtilsEdit.frx":000C
      CueBanner       =   "frmUtilsEdit.frx":002C
   End
   Begin prjDIADBS.TextBoxW txtUtilName 
      Height          =   330
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "frmUtilsEdit.frx":004C
      CueBanner       =   "frmUtilsEdit.frx":006C
   End
   Begin prjDIADBS.ctlUcPickBox ucPathUtil 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   4935
      _ExtentX        =   10398
      _ExtentY        =   556
      UseAutoForeColor=   0   'False
      Color           =   0
      DefaultExt      =   ""
      DialogType      =   3
      Enabled         =   0   'False
      FileFlags       =   2621446
      Filters         =   "Supported files|*.*|All Files (*.*)"
      ToolTipText3    =   "Click Here to Locate File"
   End
   Begin prjDIADBS.ctlUcPickBox ucPathUtil64 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   4935
      _ExtentX        =   10398
      _ExtentY        =   556
      UseAutoForeColor=   0   'False
      Color           =   0
      DefaultExt      =   ""
      DialogType      =   3
      Enabled         =   0   'False
      FileFlags       =   2621446
      Filters         =   "Supported files|*.*|All Files (*.*)"
      ToolTipText3    =   "Click Here to Locate File"
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   750
      Left            =   5760
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1323
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12244692
      Caption         =   "��������� ��������� � �����"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Height          =   750
      Left            =   3840
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12244692
      Caption         =   "����� ��� ����������"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.LabelW lblPathUtil64 
      Height          =   450
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "���� �� ������������ ����� x64"
   End
   Begin prjDIADBS.LabelW lblParamUtil 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1580
      Width           =   2415
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "�������������� ��������� �������"
   End
   Begin prjDIADBS.LabelW lblUtilName 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2415
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "������������ �������:"
   End
   Begin prjDIADBS.LabelW lblPathUtil 
      Height          =   400
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   2415
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "���� �� ������������ �����"
   End
End
Attribute VB_Name = "frmUtilsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFormName                     As String

Private Sub FontCharsetChange()
' ���������� �����
    With Me.Font
        .Name = strOtherForm_FontName
        .Size = lngOtherForm_FontSize
        .Charset = lngDialog_Charset
    End With

End Sub

'! -----------------------------------------------------------
'!  �������     :  cmdExit_Click
'!  ����������  :
'!  ��������    :  ������ �����
'! -----------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'! -----------------------------------------------------------
'!  �������     :  cmdOK_Click
'!  ����������  :
'!  ��������    :  ������ ��
'! -----------------------------------------------------------
Private Sub cmdOK_Click()
    SaveOptions
    Unload Me

End Sub

Private Sub Form_Activate()
    txtUtilName_Change
    'txtUtilName.SetFocus
End Sub

'! -----------------------------------------------------------
'!  �������     :  Form_KeyDown
'!  ����������  :  KeyCode As Integer, Shift As Integer
'!  ��������    :  ��������� ������� ������ ����������
'! -----------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me

    End If

End Sub

'! -----------------------------------------------------------
'!  �������     :  Form_Load
'!  ����������  :
'!  ��������    :  ��������� ��� �������� �����
'! -----------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me


    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmUtilsEdit", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork

    ' ����������z ����������
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' ���������� �����
        FontCharsetChange

    End If

End Sub

'Private Sub Form_Terminate()
'
'    On Error Resume Next
'
'    If Forms.Count = 0 Then
'        UnloadApp
'
'    End If
'
'End Sub

Private Sub Localise(ByVal StrPathFile As String)

' ���������� ����� ��������� (��������� ������ �� �� ��� ������� �� �������������� ������)
    FontCharsetChange
    ' �������� �����
    Me.Caption = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
    ' ������
    lblUtilName.Caption = LocaliseString(StrPathFile, strFormName, "lblUtilName", lblUtilName.Caption)
    lblPathUtil.Caption = LocaliseString(StrPathFile, strFormName, "lblPathUtil", lblPathUtil.Caption)
    lblPathUtil64.Caption = LocaliseString(StrPathFile, strFormName, "lblPathUtil64", lblPathUtil64.Caption)
    lblParamUtil.Caption = LocaliseString(StrPathFile, strFormName, "lblParamUtil", lblParamUtil.Caption)
    '������
    cmdOK.Caption = LocaliseString(StrPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(StrPathFile, strFormName, "cmdExit", cmdExit.Caption)

End Sub

'! -----------------------------------------------------------
'!  �������     :  SaveOptions
'!  ����������  :
'!  ��������    :  ���������� ��������
'! -----------------------------------------------------------
Private Sub SaveOptions()

Dim i                                   As Long

    If mbAddInList Then
        i = LastIdUtil + 1

        With frmOptions.lvUtils.ListItems.Add(, , txtUtilName)
            .SubItems(1) = ucPathUtil.Path
            .SubItems(2) = ucPathUtil64.Path
            .SubItems(3) = txtParamUtil
        End With

        'frmOptions
    Else

        With frmOptions.lvUtils
            i = .SelectedItem.Index
            .ListItems.Item(i).Text = txtUtilName
            .ListItems.Item(i).SubItems(1) = ucPathUtil.Path
            .ListItems.Item(i).SubItems(2) = ucPathUtil64.Path

            'frmOptions
            If txtParamUtil.Text <> "�������������� ��������� �������" Then
                .ListItems.Item(i).SubItems(3) = txtParamUtil
            Else
                .ListItems.Item(i).SubItems(3) = vbNullString

            End If
        End With

    End If

    LastIdUtil = frmOptions.lvUtils.ListItems.Count
    frmOptions.lvUtils.Refresh
    mbAddInList = False

End Sub

Private Sub txtParamUtil_GotFocus()
    HighlightActiveControl Me, txtParamUtil, True
End Sub

Private Sub txtParamUtil_LostFocus()
    HighlightActiveControl Me, txtParamUtil, False
End Sub

Private Sub txtUtilName_Change()
    cmdOK.Enabled = LenB(Trim$(txtUtilName)) > 0 And LenB(Trim$(ucPathUtil.Path)) > 0

End Sub

Private Sub txtUtilName_GotFocus()
    HighlightActiveControl Me, txtUtilName, True
End Sub

Private Sub txtUtilName_LostFocus()
    HighlightActiveControl Me, txtUtilName, False
End Sub

Private Sub ucPathUtil_GotFocus()
    HighlightActiveControl Me, ucPathUtil, True
End Sub

Private Sub ucPathUtil_LostFocus()
    HighlightActiveControl Me, ucPathUtil, False
End Sub

'! -----------------------------------------------------------
'!  �������     :  ucPathUtil64_Click
'!  ����������  :
'!  ��������    :
'! -----------------------------------------------------------
Private Sub ucPathUtil64_Click()

Dim strTempPath                         As String

    If ucPathUtil64.FileCount > 0 Then
        strTempPath = ucPathUtil64.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)

        End If

    End If

    If LenB(strTempPath) > 0 Then
        ucPathUtil64.Path = strTempPath

    End If

End Sub

Private Sub ucPathUtil64_GotFocus()
    HighlightActiveControl Me, ucPathUtil64, True
End Sub

Private Sub ucPathUtil64_LostFocus()
    HighlightActiveControl Me, ucPathUtil64, False
End Sub

Private Sub ucPathUtil64_PathChanged()
    cmdOK.Enabled = LenB(Trim$(txtUtilName)) > 0 And LenB(Trim$(ucPathUtil.Path)) > 0

End Sub

'! -----------------------------------------------------------
'!  �������     :  ucPathUtil_Click
'!  ����������  :
'!  ��������    :
'! -----------------------------------------------------------
Private Sub ucPathUtil_Click()

Dim strTempPath                         As String

    If ucPathUtil.FileCount > 0 Then
        strTempPath = ucPathUtil.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)

        End If

    End If

    If LenB(strTempPath) > 0 Then
        ucPathUtil.Path = strTempPath

    End If

End Sub

Private Sub ucPathUtil_PathChanged()
    cmdOK.Enabled = LenB(Trim$(txtUtilName)) > 0 And LenB(Trim$(ucPathUtil.Path)) > 0

End Sub
