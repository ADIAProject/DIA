VERSION 5.00
Begin VB.Form frmUtilsEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Редактирование записи"
   ClientHeight    =   2790
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
   ScaleHeight     =   2790
   ScaleWidth      =   7665
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.TextBoxW txtParamUtil 
      Height          =   330
      Left            =   2640
      TabIndex        =   7
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
      TabIndex        =   1
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
      TabIndex        =   3
      Top             =   600
      Width           =   4935
      _ExtentX        =   10398
      _ExtentY        =   556
      UseAutoForeColor=   0   'False
      DefaultExt      =   ""
      DialogType      =   1
      Enabled         =   0   'False
      FileFlags       =   2621446
      Filters         =   "Supported files|*.*|All Files (*.*)"
   End
   Begin prjDIADBS.ctlUcPickBox ucPathUtil64 
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   4935
      _ExtentX        =   10398
      _ExtentY        =   556
      UseAutoForeColor=   0   'False
      DefaultExt      =   ""
      DialogType      =   1
      Enabled         =   0   'False
      FileFlags       =   2621446
      Filters         =   "Supported files|*.*|All Files (*.*)"
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   650
      Left            =   5760
      TabIndex        =   9
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
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
      Left            =   3840
      TabIndex        =   8
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
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
   Begin prjDIADBS.LabelW lblPathUtil64 
      Height          =   450
      Left            =   120
      TabIndex        =   4
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
      Caption         =   "Путь до исполняемого файла x64"
   End
   Begin prjDIADBS.LabelW lblParamUtil 
      Height          =   495
      Left            =   120
      TabIndex        =   6
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
      Caption         =   "Дополнительные параметры запуска"
   End
   Begin prjDIADBS.LabelW lblUtilName 
      Height          =   375
      Left            =   120
      TabIndex        =   0
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
      Caption         =   "Наименование утилиты:"
   End
   Begin prjDIADBS.LabelW lblPathUtil 
      Height          =   400
      Left            =   120
      TabIndex        =   2
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
      Caption         =   "Путь до исполняемого файла"
   End
End
Attribute VB_Name = "frmUtilsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFormName As String

Public Property Get CaptionW() As String
    Dim lngLenStr As Long
    
    lngLenStr = DefWindowProc(Me.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    CaptionW = Space$(lngLenStr)
    DefWindowProc Me.hWnd, WM_GETTEXT, Len(CaptionW) + 1, ByVal StrPtr(CaptionW)
End Property

Public Property Let CaptionW(ByVal NewValue As String)
    DefWindowProc Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue & vbNullChar)
End Property

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
    lblUtilName.Caption = LocaliseString(strPathFile, strFormName, "lblUtilName", lblUtilName.Caption)
    lblPathUtil.Caption = LocaliseString(strPathFile, strFormName, "lblPathUtil", lblPathUtil.Caption)
    lblPathUtil64.Caption = LocaliseString(strPathFile, strFormName, "lblPathUtil64", lblPathUtil64.Caption)
    lblParamUtil.Caption = LocaliseString(strPathFile, strFormName, "lblParamUtil", lblParamUtil.Caption)
    'Кнопки
    cmdOK.Caption = LocaliseString(strPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveOptions
'! Description (Описание)  :   [Сохранение настроек]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SaveOptions()

    Dim i As Long

    If mbAddInList Then
        i = lngLastIdUtil + 1

        With frmOptions.lvUtils.ListItems.Add(, , txtUtilName)
            .SubItems(1) = ucPathUtil.Path
            .SubItems(2) = ucPathUtil64.Path
            .SubItems(3) = txtParamUtil
        End With

        'frmOptions
    Else

        With frmOptions.lvUtils
            i = .SelectedItem.Index
            .ListItems.item(i).Text = txtUtilName
            .ListItems.item(i).SubItems(1) = ucPathUtil.Path
            .ListItems.item(i).SubItems(2) = ucPathUtil64.Path

            'frmOptions
            If txtParamUtil.Text <> "Дополнительные параметры запуска" Then
                .ListItems.item(i).SubItems(3) = txtParamUtil
            Else
                .ListItems.item(i).SubItems(3) = vbNullString
            End If

        End With

    End If

    lngLastIdUtil = frmOptions.lvUtils.ListItems.Count
    frmOptions.lvUtils.Refresh
    mbAddInList = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdExit_Click
'! Description (Описание)  :   [нажали выход]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdOK_Click
'! Description (Описание)  :   [нажали ок]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
    SaveOptions
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Activate
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()
    txtUtilName_Change
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
'! Description (Описание)  :   [обработка при загрузке формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmUtilsEdit", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    LoadIconImage2Object cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2Object cmdExit, "BTN_EXIT", strPathImageMainWork

    ' Локализация приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtParamUtil_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtParamUtil_GotFocus()
    HighlightActiveControl Me, txtParamUtil, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtParamUtil_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtParamUtil_LostFocus()
    HighlightActiveControl Me, txtParamUtil, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtUtilName_Change
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtUtilName_Change()
    cmdOK.Enabled = LenB(Trim$(txtUtilName)) And LenB(Trim$(ucPathUtil.Path))
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtUtilName_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtUtilName_GotFocus()
    HighlightActiveControl Me, txtUtilName, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtUtilName_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtUtilName_LostFocus()
    HighlightActiveControl Me, txtUtilName, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathUtil64_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathUtil64_Click()

    Dim strTempPath As String

    If ucPathUtil64.FileCount Then
        strTempPath = ucPathUtil64.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucPathUtil64.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathUtil64_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathUtil64_GotFocus()
    HighlightActiveControl Me, ucPathUtil64, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathUtil64_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathUtil64_LostFocus()
    HighlightActiveControl Me, ucPathUtil64, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathUtil64_PathChanged
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathUtil64_PathChanged()
    cmdOK.Enabled = LenB(Trim$(txtUtilName)) And LenB(Trim$(ucPathUtil.Path))
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathUtil_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathUtil_Click()

    Dim strTempPath As String

    If ucPathUtil.FileCount Then
        strTempPath = ucPathUtil.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucPathUtil.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathUtil_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathUtil_GotFocus()
    HighlightActiveControl Me, ucPathUtil, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathUtil_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathUtil_LostFocus()
    HighlightActiveControl Me, ucPathUtil, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathUtil_PathChanged
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathUtil_PathChanged()
    cmdOK.Enabled = LenB(Trim$(txtUtilName)) And LenB(Trim$(ucPathUtil.Path))
End Sub
