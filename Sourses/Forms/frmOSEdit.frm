VERSION 5.00
Begin VB.Form frmOSEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Редактирование записи"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   8355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOSEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8355
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.TextBoxW txtOSVer 
      Height          =   375
      Left            =   2000
      TabIndex        =   1
      Top             =   240
      Width           =   6225
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
   End
   Begin prjDIADBS.FrameW frExcludeFileName 
      Height          =   1250
      Left            =   120
      Top             =   3850
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2196
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483635
      Caption         =   "frmOSEdit.frx":000C
      Begin prjDIADBS.TextBoxW txtExcludeFileName 
         Height          =   850
         Left            =   120
         TabIndex        =   15
         Top             =   280
         Width           =   7935
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
   End
   Begin prjDIADBS.FrameW frDopFile 
      Height          =   1650
      Left            =   120
      Top             =   2150
      Width           =   8175
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
      ForeColor       =   -2147483635
      Caption         =   "frmOSEdit.frx":0094
      Begin prjDIADBS.ctlUcPickBox ucPhysXPath 
         Height          =   315
         Left            =   2760
         TabIndex        =   10
         Top             =   270
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.7z|7z Files (*.7z)"
         Locked          =   -1  'True
      End
      Begin prjDIADBS.ctlUcPickBox ucLangPath 
         Height          =   315
         Left            =   2760
         TabIndex        =   12
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.7z|7z Files (*.7z)"
         Locked          =   -1  'True
      End
      Begin prjDIADBS.ctlUcPickBox ucRuntimesPath 
         Height          =   315
         Left            =   2760
         TabIndex        =   14
         Top             =   1185
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.7z|7z Files (*.7z)"
         Locked          =   -1  'True
      End
      Begin prjDIADBS.LabelW lblRuntimes 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1185
         Width           =   2535
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
         Caption         =   "DP_Runtimes_*.7z"
      End
      Begin prjDIADBS.LabelW lblPhysX 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   2535
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
         Caption         =   "DP_Graphics_A_PhysX_*.7z"
      End
      Begin prjDIADBS.LabelW lblLang 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2535
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
         Caption         =   "DP_Graphics_Languages_*.7z"
      End
      Begin VB.Image imgNo 
         Height          =   480
         Left            =   8535
         Top             =   45
         Width           =   480
      End
   End
   Begin prjDIADBS.CheckBoxW chk64bit 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5150
      Width           =   4215
      _ExtentX        =   7435
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
      Caption         =   "frmOSEdit.frx":0154
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.TextBoxW txtOSName 
      Height          =   375
      Left            =   2000
      TabIndex        =   3
      Top             =   720
      Width           =   6225
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
   End
   Begin prjDIADBS.CheckBoxW chkNotCheckBitOS 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5500
      Width           =   4215
      _ExtentX        =   7435
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
      Caption         =   "frmOSEdit.frx":01A6
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.ctlUcPickBox ucPathDRP 
      Height          =   315
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Width           =   5275
      _ExtentX        =   9313
      _ExtentY        =   556
      UseAutoForeColor=   0   'False
      DefaultExt      =   ""
      Enabled         =   0   'False
      Filters         =   "Supported files|*.*|All Files (*.*)"
      Locked          =   -1  'True
   End
   Begin prjDIADBS.ctlUcPickBox ucPathDB 
      Height          =   315
      Left            =   2880
      TabIndex        =   7
      Top             =   1710
      Width           =   5275
      _ExtentX        =   9313
      _ExtentY        =   556
      DefaultExt      =   ""
      Enabled         =   0   'False
      Filters         =   "Supported files|*.*|All Files (*.*)"
      Locked          =   -1  'True
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   650
      Left            =   6420
      TabIndex        =   9
      Top             =   5150
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
      Left            =   4500
      TabIndex        =   18
      Top             =   5150
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
   Begin prjDIADBS.LabelW lblPathDB 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1710
      Width           =   2655
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
      Caption         =   "Путь до каталога хранения БД"
   End
   Begin prjDIADBS.LabelW lblPathDRP 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
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
      Caption         =   "Путь до каталога с пакетами драйверов"
   End
   Begin prjDIADBS.LabelW lblNameOS 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
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
      Caption         =   "Наименование ОС"
   End
   Begin prjDIADBS.LabelW lblOSVer 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
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
      Caption         =   "Версия ОС"
   End
End
Attribute VB_Name = "frmOSEdit"
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
    lblOSVer.Caption = LocaliseString(strPathFile, strFormName, "lblOSVer", lblOSVer.Caption)
    lblNameOS.Caption = LocaliseString(strPathFile, strFormName, "lblNameOS", lblNameOS.Caption)
    lblPathDRP.Caption = LocaliseString(strPathFile, strFormName, "lblPathDRP", lblPathDRP.Caption)
    lblPathDB.Caption = LocaliseString(strPathFile, strFormName, "lblPathDB", lblPathDB.Caption)
    frDopFile.Caption = LocaliseString(strPathFile, strFormName, "frDopFile", frDopFile.Caption)
    lblPhysX.Caption = LocaliseString(strPathFile, strFormName, "lblPhysX", lblPhysX.Caption)
    lblLang.Caption = LocaliseString(strPathFile, strFormName, "lblLang", lblLang.Caption)
    lblRuntimes.Caption = LocaliseString(strPathFile, strFormName, "lblRuntimes", lblRuntimes.Caption)
    frExcludeFileName.Caption = LocaliseString(strPathFile, strFormName, "frExcludeFileName", frExcludeFileName.Caption)
    chk64bit.Caption = LocaliseString(strPathFile, strFormName, "chk64bit", chk64bit.Caption)
    chkNotCheckBitOS.Caption = LocaliseString(strPathFile, strFormName, "chkNotCheckBitOS", chkNotCheckBitOS.Caption)
    'Кнопки
    cmdOK.Caption = LocaliseString(strPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
    ' Сообщения диалогов выбора файлов и каталогов
    ucPathDRP.ToolTipTexts(ucFolder) = strMessages(152)
    ucPathDRP.DialogMsg(ucFolder) = strMessages(152)
    ucPathDB.ToolTipTexts(ucFolder) = strMessages(152)
    ucPathDB.DialogMsg(ucFolder) = strMessages(152)
    ucPhysXPath.ToolTipTexts(ucOpen) = strMessages(151)
    ucPhysXPath.DialogMsg(ucOpen) = strMessages(151)
    ucRuntimesPath.ToolTipTexts(ucOpen) = strMessages(151)
    ucRuntimesPath.DialogMsg(ucOpen) = strMessages(151)
    ucLangPath.ToolTipTexts(ucOpen) = strMessages(151)
    ucLangPath.DialogMsg(ucOpen) = strMessages(151)
     
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveOptions
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SaveOptions()

    Dim ii As Long

    If mbAddInList Then
        ii = lngLastIdOS + 1

        With frmOptions.lvOS.ListItems.Add(, , txtOSVer)
            .SubItems(1) = txtOSName
            .SubItems(2) = ucPathDRP.Path
            .SubItems(3) = ucPathDB.Path

            If chk64bit.Value Then
                If chkNotCheckBitOS.Value Then
                    .SubItems(4) = "3"
                Else
                    .SubItems(4) = "1"
                End If

            Else

                If chkNotCheckBitOS.Value Then
                    .SubItems(4) = "2"
                Else
                    .SubItems(4) = "0"
                End If
            End If

            .SubItems(5) = ucPhysXPath.Path
            .SubItems(6) = ucLangPath.Path
            .SubItems(7) = ucRuntimesPath.Path
            .SubItems(8) = txtExcludeFileName
        End With

    Else

        With frmOptions.lvOS
            ii = .SelectedItem.Index
            .ListItems.item(ii).Text = txtOSVer
            .ListItems.item(ii).SubItems(1) = txtOSName
            .ListItems.item(ii).SubItems(2) = ucPathDRP.Path
            .ListItems.item(ii).SubItems(3) = ucPathDB.Path

            If chk64bit.Value Then
                If chkNotCheckBitOS.Value Then
                    .ListItems.item(ii).SubItems(4) = "3"
                Else
                    .ListItems.item(ii).SubItems(4) = "1"
                End If

            Else

                If chkNotCheckBitOS.Value Then
                    .ListItems.item(ii).SubItems(4) = "2"
                Else
                    .ListItems.item(ii).SubItems(4) = "0"
                End If
            End If

            .ListItems.item(ii).SubItems(5) = ucPhysXPath.Path
            .ListItems.item(ii).SubItems(6) = ucLangPath.Path
            .ListItems.item(ii).SubItems(7) = ucRuntimesPath.Path
            .ListItems.item(ii).SubItems(8) = txtExcludeFileName
        End With


    End If

    lngLastIdOS = frmOptions.lvOS.ListItems.count
    frmOptions.lvOS.Refresh
    mbAddInList = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkNotCheckBitOS_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkNotCheckBitOS_Click()
    chk64bit.Enabled = Not chkNotCheckBitOS.Value
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

    If ucPathDB.Path = "Путь до каталога хранения БД" Then
        ucPathDB.Path = BackslashAdd2Path(ucPathDRP.Path) & "dev_db"
    End If

    SaveOptions
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Activate
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()
    txtOSVer_Change
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

    ' Устанавливаем картинки кнопок
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
'! Procedure   (Функция)   :   Sub txtExcludeFileName_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtExcludeFileName_GotFocus()
    HighlightActiveControl Me, txtExcludeFileName, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtExcludeFileName_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtExcludeFileName_LostFocus()
    HighlightActiveControl Me, txtExcludeFileName, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtOSName_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtOSName_GotFocus()
    HighlightActiveControl Me, txtOSName, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtOSName_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtOSName_LostFocus()
    HighlightActiveControl Me, txtOSName, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtOSVer_Change
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtOSVer_Change()
    cmdOK.Enabled = LenB(Trim$(txtOSVer)) And LenB(Trim$(ucPathDRP.Path))
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtOSVer_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtOSVer_GotFocus()
    HighlightActiveControl Me, txtOSVer, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtOSVer_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtOSVer_LostFocus()
    HighlightActiveControl Me, txtOSVer, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucLangPath_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucLangPath_Click()

    Dim strTempPath As String

    If ucLangPath.FileCount Then
        strTempPath = ucLangPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucLangPath.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucLangPath_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucLangPath_GotFocus()
    HighlightActiveControl Me, ucLangPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucLangPath_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucLangPath_LostFocus()
    HighlightActiveControl Me, ucLangPath, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathDB_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathDB_Click()

    Dim strTempPath As String

    If ucPathDB.FileCount Then
        strTempPath = ucPathDB.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucPathDB.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathDB_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathDB_GotFocus()
    HighlightActiveControl Me, ucPathDB, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathDB_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathDB_LostFocus()
    HighlightActiveControl Me, ucPathDB, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathDRP_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathDRP_Click()

    Dim strTempPath As String

    If ucPathDRP.FileCount Then
        strTempPath = ucPathDRP.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucPathDRP.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathDRP_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathDRP_GotFocus()
    HighlightActiveControl Me, ucPathDRP, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPathDRP_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPathDRP_LostFocus()
    HighlightActiveControl Me, ucPathDRP, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPhysXPath_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPhysXPath_Click()

    Dim strTempPath As String

    If ucPhysXPath.FileCount Then
        strTempPath = ucPhysXPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucPhysXPath.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPhysXPath_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPhysXPath_GotFocus()
    HighlightActiveControl Me, ucPhysXPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucPhysXPath_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucPhysXPath_LostFocus()
    HighlightActiveControl Me, ucPhysXPath, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucRuntimesPath_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucRuntimesPath_Click()

    Dim strTempPath As String

    If ucRuntimesPath.FileCount Then
        strTempPath = ucRuntimesPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucRuntimesPath.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucRuntimesPath_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucRuntimesPath_GotFocus()
    HighlightActiveControl Me, ucRuntimesPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucRuntimesPath_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucRuntimesPath_LostFocus()
    HighlightActiveControl Me, ucRuntimesPath, False
End Sub

