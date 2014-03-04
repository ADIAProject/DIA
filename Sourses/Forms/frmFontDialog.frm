VERSION 5.00
Begin VB.Form frmFontDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Locate Font and Color ..."
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjDIADBS.TextBoxW txtFont 
      Height          =   495
      Left            =   60
      TabIndex        =   11
      Top             =   1260
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "frmFontDialog.frx":0000
      Alignment       =   2
      CueBanner       =   "frmFontDialog.frx":0052
   End
   Begin prjDIADBS.OptionButtonW optControl 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   0   'False
      Caption         =   "frmFontDialog.frx":0072
   End
   Begin prjDIADBS.OptionButtonW optControl 
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   0   'False
      Caption         =   "frmFontDialog.frx":009E
   End
   Begin prjDIADBS.OptionButtonW optControl 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmFontDialog.frx":00D2
   End
   Begin prjDIADBS.SpinBox txtFontSize 
      Height          =   315
      Left            =   1860
      TabIndex        =   4
      Top             =   420
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   6
      Max             =   20
      Value           =   6
      AllowOnlyNumbers=   -1  'True
   End
   Begin prjDIADBS.ctlColorButton ctlFontColor 
      Height          =   330
      Left            =   1980
      TabIndex        =   3
      Top             =   780
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   582
      Icon            =   "frmFontDialog.frx":00FC
   End
   Begin prjDIADBS.CheckBoxW chkItalic 
      Height          =   315
      Left            =   2700
      TabIndex        =   2
      Top             =   675
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmFontDialog.frx":0256
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.CheckBoxW chkBold 
      Height          =   255
      Left            =   2700
      TabIndex        =   1
      Top             =   420
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmFontDialog.frx":0282
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.ctlFontCombo ctlFontCombo 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      PreviewText     =   "ctlFontCombo1"
      ComboFontSize   =   10
      ButtonOverColor =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjDIADBS.CheckBoxW chkUnderline 
      Height          =   315
      Left            =   2700
      TabIndex        =   5
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmFontDialog.frx":02AA
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   650
      Left            =   2280
      TabIndex        =   9
      Top             =   1860
      Width           =   2100
      _extentx        =   3704
      _extenty        =   1138
      font            =   "frmFontDialog.frx":02DC
      buttonstyle     =   8
      backcolor       =   12244692
      caption         =   "��������� ��������� � �����"
      pictureshadow   =   -1  'True
      picturepushonhover=   -1  'True
      captioneffects  =   0
      picturealign    =   0
      colorscheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Default         =   -1  'True
      Height          =   650
      Left            =   60
      TabIndex        =   10
      Top             =   1860
      Width           =   2100
      _extentx        =   3704
      _extenty        =   1138
      font            =   "frmFontDialog.frx":0304
      buttonstyle     =   8
      backcolor       =   12244692
      caption         =   "����� ��� ����������"
      pictureshadow   =   -1  'True
      picturepushonhover=   -1  'True
      captioneffects  =   0
      picturealign    =   0
      colorscheme     =   3
   End
   Begin prjDIADBS.OptionButtonW optControl 
      Height          =   255
      Index           =   2
      Left            =   2100
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   0   'False
      Caption         =   "frmFontDialog.frx":032C
   End
   Begin prjDIADBS.LabelW lblFontSize 
      Height          =   375
      Left            =   60
      TabIndex        =   12
      Top             =   420
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Caption         =   "������ ������"
   End
   Begin prjDIADBS.LabelW lblFontColor 
      Height          =   375
      Left            =   60
      TabIndex        =   13
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Caption         =   "���� ������"
   End
End
Attribute VB_Name = "frmFontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFormName As String
Private m_Caption   As String

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkBold_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkBold_Click()
    ctlFontCombo.ComboFontBold = chkBold.Value = 1
    txtFont.Font.Bold = chkBold.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkItalic_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkItalic_Click()
    ctlFontCombo.ComboFontItalic = chkItalic.Value = 1
    txtFont.Font.Italic = chkItalic.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ctlFontColor_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ctlFontColor_Click()
    txtFont.ForeColor = ctlFontColor.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkUnderline_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkUnderline_Click()
    txtFont.Font.Underline = chkUnderline.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ctlFontCombo_FontNotFound
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   FontName (String)
'!--------------------------------------------------------------------------------
Private Sub ctlFontCombo_FontNotFound(FontName As String)
    MsgBox "Cant find this font: " & FontName
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ctlFontCombo_SelectedFontChanged
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   NewFontName (String)
'!--------------------------------------------------------------------------------
Private Sub ctlFontCombo_SelectedFontChanged(NewFontName As String)
    txtFont.Font.Name = NewFontName
    ctlFontCombo.ClearUsedList
    ctlFontCombo.AddToUsedList NewFontName
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_Activate
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()
    ctlFontCombo.SelectedFont = txtFont.Font.Name
    txtFontSize.Value = txtFont.Font.Size
    ctlFontCombo.PreviewText = txtFont.Text
    ctlFontCombo.AddToUsedList txtFont.Font.Name
    chkBold.Value = txtFont.Font.Bold
    chkItalic.Value = txtFont.Font.Italic
    chkUnderline.Value = txtFont.Font.Underline
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_KeyDown
'! Description (��������)  :   [��������� ������� ������ ����������]
'! Parameters  (����������):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_Load
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmFontDialog", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    ' ������������� �������� ������ � ������� �������� ������
    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork

    ' ����������z ����������
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' ���������� �����
        FontCharsetChange
    End If

    txtFontSize.Min = 6
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Localise
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal StrPathFile As String)
    ' ���������� ����� ��������� (��������� ������ �� �� ��� ������� �� �������������� ������)
    FontCharsetChange
    ' �������� �����
    Me.CaptionW = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
    ' ������
    lblFontSize.Caption = LocaliseString(StrPathFile, strFormName, "lblFontSize", lblFontSize.Caption)
    lblFontColor.Caption = LocaliseString(StrPathFile, strFormName, "lblFontColor", lblFontColor.Caption)
    chkBold.Caption = LocaliseString(StrPathFile, strFormName, "chkBold", chkBold.Caption)
    chkItalic.Caption = LocaliseString(StrPathFile, strFormName, "chkItalic", chkItalic.Caption)
    chkUnderline.Caption = LocaliseString(StrPathFile, strFormName, "chkUnderline", chkUnderline.Caption)
    txtFont.Text = LocaliseString(StrPathFile, strFormName, "txtFont", txtFont.Text)
    ctlFontColor.DropDownCaption = LocaliseString(StrPathFile, strFormName, "ctlFontColor", ctlFontColor.DropDownCaption)
    
    '������
    cmdOK.Caption = LocaliseString(StrPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(StrPathFile, strFormName, "cmdExit", cmdExit.Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub FontCharsetChange
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub FontCharsetChange()

    ' ���������� �����
    With Me.Font
        .Name = strFontOtherForm_Name
        .Size = lngFontOtherForm_Size
        .Charset = lngFont_Charset
    End With

    SetBtnFontProperties cmdExit
    SetBtnFontProperties cmdOK
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtFont_Change
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtFont_Change()
    ctlFontCombo.PreviewText = txtFont.Text
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdExit_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdOK_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
    SaveOptions
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub SaveOptions
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub SaveOptions()

    With txtFont

        If optControl.Item(0).Value Then
            strFontTab_Name = .Font.Name
            miFontTab_Size = .Font.Size
            mbFontTab_Underline = .Font.Underline
            mbFontTab_Strikethru = .Font.Strikethrough
            mbFontTab_Bold = .Font.Bold
            mbFontTab_Italic = .Font.Italic
            lngFontTab_Color = .ForeColor
            
        ElseIf optControl.Item(1).Value Then
            strFontTab2_Name = .Font.Name
            miFontTab2_Size = .Font.Size
            mbFontTab2_Underline = .Font.Underline
            mbFontTab2_Strikethru = .Font.Strikethrough
            mbFontTab2_Bold = .Font.Bold
            mbFontTab2_Italic = .Font.Italic
            lngFontTab2_Color = .ForeColor
            
        ElseIf optControl.Item(2).Value Then
            strFontTT_Name = .Font.Name
            miFontTT_Size = .Font.Size
            mbFontTT_Underline = .Font.Underline
            mbFontTT_Strikethru = .Font.Strikethrough
            mbFontTT_Bold = .Font.Bold
            mbFontTT_Italic = .Font.Italic
            lngFontTT_Color = .ForeColor
            SetTTFontProperties frmOptions.TT
            
        ElseIf optControl.Item(3).Value Then
            strFontBtn_Name = .Font.Name
            miFontBtn_Size = .Font.Size
            mbFontBtn_Underline = .Font.Underline
            mbFontBtn_Strikethru = .Font.Strikethrough
            mbFontBtn_Bold = .Font.Bold
            mbFontBtn_Italic = .Font.Italic
            lngFontBtn_Color = .ForeColor
            SetBtnStatusFontProperties frmOptions.cmdFutureButton
            frmOptions.cmdFutureButton.ForeColor = .ForeColor
            
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtFontSize_Change
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtFontSize_Change()
    txtFont.Font.Size = txtFontSize.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtFontSize_TextChange
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtFontSize_TextChange()
    txtFont.Font.Size = txtFontSize.Value
End Sub

Public Property Let CaptionW(ByVal NewValue As String)
    DefWindowProc Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue & vbNullChar)
End Property

Public Property Get CaptionW() As String
    Dim strLen As Long
    strLen = DefWindowProc(Me.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    CaptionW = Space$(strLen)
    DefWindowProc Me.hWnd, WM_GETTEXT, Len(CaptionW) + 1, ByVal StrPtr(CaptionW)
End Property

