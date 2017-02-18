VERSION 5.00
Begin VB.Form frmFontDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locate Font and Color ..."
   ClientHeight    =   2625
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
   Icon            =   "frmFontDialog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin prjDIADBS.TextBoxW txtFont 
      Height          =   495
      Left            =   60
      TabIndex        =   5
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
      Text            =   "frmFontDialog.frx":000C
      Alignment       =   2
   End
   Begin prjDIADBS.OptionButtonW optControl 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   14
      TabStop         =   0   'False
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
      Caption         =   "frmFontDialog.frx":005E
   End
   Begin prjDIADBS.OptionButtonW optControl 
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   12
      TabStop         =   0   'False
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
      Caption         =   "frmFontDialog.frx":008A
   End
   Begin prjDIADBS.OptionButtonW optControl 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
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
      Caption         =   "frmFontDialog.frx":00BE
   End
   Begin prjDIADBS.SpinBox txtFontSize 
      Height          =   315
      Left            =   1860
      TabIndex        =   1
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
      TabIndex        =   10
      Top             =   840
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   582
      Icon            =   "frmFontDialog.frx":00E8
   End
   Begin prjDIADBS.CheckBoxW chkItalic 
      Height          =   315
      Left            =   2700
      TabIndex        =   3
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
      Caption         =   "frmFontDialog.frx":066E
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.CheckBoxW chkBold 
      Height          =   255
      Left            =   2700
      TabIndex        =   2
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
      Caption         =   "frmFontDialog.frx":069A
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
      TabIndex        =   4
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
      Caption         =   "frmFontDialog.frx":06C2
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   650
      Left            =   2280
      TabIndex        =   7
      Top             =   1860
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   8
      BackColor       =   12244692
      Caption         =   "Сохранить изменения и выйти"
      CaptionEffects  =   0
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Default         =   -1  'True
      Height          =   650
      Left            =   60
      TabIndex        =   6
      Top             =   1860
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   8
      BackColor       =   12244692
      Caption         =   "Выход без сохранения"
      CaptionEffects  =   0
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.OptionButtonW optControl 
      Height          =   255
      Index           =   2
      Left            =   2100
      TabIndex        =   13
      TabStop         =   0   'False
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
      Caption         =   "frmFontDialog.frx":06F4
   End
   Begin prjDIADBS.LabelW lblFontSize 
      Height          =   375
      Left            =   60
      TabIndex        =   8
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
      Caption         =   "Размер шрифта"
   End
   Begin prjDIADBS.LabelW lblFontColor 
      Height          =   375
      Left            =   60
      TabIndex        =   9
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
      Caption         =   "Цвет шрифта"
   End
End
Attribute VB_Name = "frmFontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFormName As String

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Get CaptionW
'! Description (Описание)  :   [Получение Caption-формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get CaptionW() As String
    Dim lngLenStr As Long
    
    lngLenStr = DefWindowProc(Me.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    CaptionW = Space$(lngLenStr)
    DefWindowProc Me.hWnd, WM_GETTEXT, Len(CaptionW) + 1, ByVal StrPtr(CaptionW)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Let CaptionW
'! Description (Описание)  :   [Изменение Caption-формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Let CaptionW(ByVal NewValue As String)
    DefWindowProc Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue & vbNullChar)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkBold_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkBold_Click()
    ctlFontCombo.ComboFontBold = chkBold.Value = 1
    txtFont.Font.Bold = chkBold.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkItalic_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkItalic_Click()
    ctlFontCombo.ComboFontItalic = chkItalic.Value = 1
    txtFont.Font.Italic = chkItalic.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkUnderline_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkUnderline_Click()
    txtFont.Font.Underline = chkUnderline.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdExit_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdOK_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
    SaveOptions
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ctlFontColor_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ctlFontColor_Click()
    txtFont.ForeColor = ctlFontColor.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ctlFontCombo_FontNotFound
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   FontName (String)
'!--------------------------------------------------------------------------------
Private Sub ctlFontCombo_FontNotFound(FontName As String)
    MsgBox "Cant find this font: " & FontName
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ctlFontCombo_SelectedFontChanged
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewFontName (String)
'!--------------------------------------------------------------------------------
Private Sub ctlFontCombo_SelectedFontChanged(NewFontName As String)
    txtFont.Font.Name = NewFontName
    ctlFontCombo.ClearUsedList
    ctlFontCombo.AddToUsedList NewFontName
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FontCharsetChange
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub FontCharsetChange()

    ' Выставляем шрифт
    With Me.Font
        .Name = strFontOtherForm_Name
        .Size = lngFontOtherForm_Size
        .Charset = lngFont_Charset
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Activate
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()
    
    ctlFontCombo.PreviewText = txtFont.Text
    With txtFont.Font
        ctlFontCombo.SelectedFont = .Name
        ctlFontCombo.AddToUsedList .Name
        txtFontSize.Value = .Size
        chkBold.Value = .Bold
        chkItalic.Value = .Italic
        chkUnderline.Value = .Underline
    End With
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_KeyDown
'! Description (Описание)  :   [обработка нажатий клавиш клавиатуры]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Load
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, strFormName, False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    ' Устанавливаем картинки кнопок и убираем описание кнопок
    LoadIconImage2Object cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2Object cmdExit, "BTN_EXIT", strPathImageMainWork

    ' Локализация приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    txtFontSize.Min = 6
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal strPathFile As String)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.CaptionW = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    ' Лэйблы
    lblFontSize.Caption = LocaliseString(strPathFile, strFormName, "lblFontSize", lblFontSize.Caption)
    lblFontColor.Caption = LocaliseString(strPathFile, strFormName, "lblFontColor", lblFontColor.Caption)
    chkBold.Caption = LocaliseString(strPathFile, strFormName, "chkBold", chkBold.Caption)
    chkItalic.Caption = LocaliseString(strPathFile, strFormName, "chkItalic", chkItalic.Caption)
    chkUnderline.Caption = LocaliseString(strPathFile, strFormName, "chkUnderline", chkUnderline.Caption)
    txtFont.Text = LocaliseString(strPathFile, strFormName, "txtFont", txtFont.Text)
    ctlFontColor.DropDownCaption = LocaliseString(strPathFile, strFormName, "ctlFontColor", ctlFontColor.DropDownCaption)
    
    'Кнопки
    cmdOK.Caption = LocaliseString(strPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveOptions
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SaveOptions()

    With txtFont.Font

        If optControl.item(0).Value Then
            strFontTab_Name = .Name
            miFontTab_Size = .Size
            mbFontTab_Underline = .Underline
            mbFontTab_Strikethru = .Strikethrough
            mbFontTab_Bold = .Bold
            mbFontTab_Italic = .Italic
            lngFontTab_Color = txtFont.ForeColor
            
        ElseIf optControl.item(1).Value Then
            strFontTab2_Name = .Name
            miFontTab2_Size = .Size
            mbFontTab2_Underline = .Underline
            mbFontTab2_Strikethru = .Strikethrough
            mbFontTab2_Bold = .Bold
            mbFontTab2_Italic = .Italic
            lngFontTab2_Color = txtFont.ForeColor
            
        ElseIf optControl.item(2).Value Then
            strFontTT_Name = .Name
            miFontTT_Size = .Size
            mbFontTT_Underline = .Underline
            mbFontTT_Strikethru = .Strikethrough
            mbFontTT_Bold = .Bold
            mbFontTT_Italic = .Italic
            lngFontTT_Color = txtFont.ForeColor
            'SetTTFontProperties frmOptions.TT
            
        ElseIf optControl.item(3).Value Then
            strFontBtn_Name = .Name
            miFontBtn_Size = .Size
            mbFontBtn_Underline = .Underline
            mbFontBtn_Strikethru = .Strikethrough
            mbFontBtn_Bold = .Bold
            mbFontBtn_Italic = .Italic
            lngFontBtn_Color = txtFont.ForeColor
            frmOptions.cmdFutureButton.ForeColor = txtFont.ForeColor
            SetBtnStatusFontProperties frmOptions.cmdFutureButton
            
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtFontSize_Change
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtFontSize_Change()
    txtFont.Font.Size = txtFontSize.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtFontSize_TextChange
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtFontSize_TextChange()
    txtFont.Font.Size = txtFontSize.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtFont_Change
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtFont_Change()
    ctlFontCombo.PreviewText = txtFont.Text
End Sub
