VERSION 5.00
Begin VB.Form frmOSEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Редактирование записи"
   ClientHeight    =   5925
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
   ScaleHeight     =   5925
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.TextBoxW txtOSVer 
      Height          =   375
      Left            =   2000
      TabIndex        =   6
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
      Text            =   "frmOSEdit.frx":000C
      CueBanner       =   "frmOSEdit.frx":002C
   End
   Begin prjDIADBS.FrameW frExcludeFileName 
      Height          =   1175
      Left            =   120
      Top             =   3850
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
      Caption         =   "frmOSEdit.frx":004C
      Begin prjDIADBS.TextBoxW txtExcludeFileName 
         Height          =   850
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
         Text            =   "frmOSEdit.frx":00D4
         MultiLine       =   -1  'True
         ScrollBars      =   2
         CueBanner       =   "frmOSEdit.frx":00F4
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
      Caption         =   "frmOSEdit.frx":0114
      Begin prjDIADBS.ctlUcPickBox ucPhysXPath 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   270
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.7z|7z Files (*.7z)"
      End
      Begin prjDIADBS.ctlUcPickBox ucLangPath 
         Height          =   315
         Left            =   2760
         TabIndex        =   3
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.7z|7z Files (*.7z)"
      End
      Begin prjDIADBS.ctlUcPickBox ucRuntimesPath 
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   1185
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.7z|7z Files (*.7z)"
      End
      Begin prjDIADBS.LabelW lblRuntimes 
         Height          =   255
         Left            =   120
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
      Begin prjDIADBS.LabelW lblNo 
         Height          =   500
         Left            =   9135
         TabIndex        =   15
         Top             =   45
         Width           =   2175
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
         Caption         =   "В БД не найдено драйверов для ваших устройств"
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
      TabIndex        =   5
      Top             =   5080
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
      Caption         =   "frmOSEdit.frx":01D4
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.TextBoxW txtOSName 
      Height          =   375
      Left            =   2000
      TabIndex        =   0
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
      Text            =   "frmOSEdit.frx":0226
      CueBanner       =   "frmOSEdit.frx":0246
   End
   Begin prjDIADBS.CheckBoxW chkNotCheckBitOS 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5475
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
      Caption         =   "frmOSEdit.frx":0266
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.ctlUcPickBox ucPathDRP 
      Height          =   315
      Left            =   2880
      TabIndex        =   9
      Top             =   1200
      Width           =   5275
      _ExtentX        =   9313
      _ExtentY        =   556
      UseAutoForeColor=   0   'False
      DefaultExt      =   ""
      Enabled         =   0   'False
      Filters         =   "Supported files|*.*|All Files (*.*)"
   End
   Begin prjDIADBS.ctlUcPickBox ucPathDB 
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Top             =   1710
      Width           =   5275
      _ExtentX        =   9313
      _ExtentY        =   556
      DefaultExt      =   ""
      Enabled         =   0   'False
      Filters         =   "Supported files|*.*|All Files (*.*)"
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   750
      Left            =   6420
      TabIndex        =   11
      Top             =   5080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   13
      BackColor       =   12244692
      Caption         =   "Сохранить изменения и выйти"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Height          =   750
      Left            =   4500
      TabIndex        =   7
      Top             =   5080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   13
      BackColor       =   12244692
      Caption         =   "Выход без сохранения"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.LabelW lblPathDB 
      Height          =   495
      Left            =   120
      TabIndex        =   16
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
      TabIndex        =   17
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
      TabIndex        =   18
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
      TabIndex        =   19
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

Private strFormName                     As String

Private Sub chkNotCheckBitOS_Click()
    chk64bit.Enabled = Not chkNotCheckBitOS.Value

End Sub

Private Sub FontCharsetChange()
' Выставляем шрифт
    With Me.Font
        .Name = strOtherForm_FontName
        .Size = lngOtherForm_FontSize
        .Charset = lngDialog_Charset
    End With

End Sub

'! -----------------------------------------------------------
'!  Функция     :  cmdExit_Click
'!  Переменные  :
'!  Описание    :
'! -----------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'! -----------------------------------------------------------
'!  Функция     :  cmdOK_Click
'!  Переменные  :
'!  Описание    :
'! -----------------------------------------------------------
Private Sub cmdOK_Click()

    If ucPathDB.Path = "Путь до каталога хранения БД" Then
        ucPathDB.Path = BackslashAdd2Path(ucPathDRP.Path) & "dev_db"

    End If

    SaveOptions
    Unload Me

End Sub

Private Sub Form_Activate()
    txtOSVer_Change
    'txtOSVer.SetFocus
End Sub

'! -----------------------------------------------------------
'!  Функция     :  Form_KeyDown
'!  Переменные  :  KeyCode As Integer, Shift As Integer
'!  Описание    :  обработка нажатий клавиш клавиатуры
'! -----------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me

    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  Form_Load
'!  Переменные  :
'!  Описание    :
'! -----------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me


    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmOSEdit", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    ' Устанавливаем картинки кнопок
    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork

    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
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

' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.Caption = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
    ' Лэйблы
    lblOSVer.Caption = LocaliseString(StrPathFile, strFormName, "lblOSVer", lblOSVer.Caption)
    lblNameOS.Caption = LocaliseString(StrPathFile, strFormName, "lblNameOS", lblNameOS.Caption)
    lblPathDRP.Caption = LocaliseString(StrPathFile, strFormName, "lblPathDRP", lblPathDRP.Caption)
    lblPathDB.Caption = LocaliseString(StrPathFile, strFormName, "lblPathDB", lblPathDB.Caption)
    frDopFile.Caption = LocaliseString(StrPathFile, strFormName, "frDopFile", frDopFile.Caption)
    lblPhysX.Caption = LocaliseString(StrPathFile, strFormName, "lblPhysX", lblPhysX.Caption)
    lblLang.Caption = LocaliseString(StrPathFile, strFormName, "lblLang", lblLang.Caption)
    lblRuntimes.Caption = LocaliseString(StrPathFile, strFormName, "lblRuntimes", lblRuntimes.Caption)
    frExcludeFileName.Caption = LocaliseString(StrPathFile, strFormName, "frExcludeFileName", frExcludeFileName.Caption)
    chk64bit.Caption = LocaliseString(StrPathFile, strFormName, "chk64bit", chk64bit.Caption)
    chkNotCheckBitOS.Caption = LocaliseString(StrPathFile, strFormName, "chkNotCheckBitOS", chkNotCheckBitOS.Caption)
    'Кнопки
    cmdOK.Caption = LocaliseString(StrPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(StrPathFile, strFormName, "cmdExit", cmdExit.Caption)

End Sub

'! -----------------------------------------------------------
'!  Функция     :  SaveOptions
'!  Переменные  :
'!  Описание    :
'! -----------------------------------------------------------
Private Sub SaveOptions()

Dim i                                   As Long

    If mbAddInList Then
        i = LastIdOS + 1

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

        'FRMOPTIONS
    Else

        With frmOptions.lvOS
            i = .SelectedItem.Index
            .ListItems.Item(i).Text = txtOSVer
            .ListItems.Item(i).SubItems(1) = txtOSName
            .ListItems.Item(i).SubItems(2) = ucPathDRP.Path
            .ListItems.Item(i).SubItems(3) = ucPathDB.Path

            If chk64bit.Value Then
                If chkNotCheckBitOS.Value Then
                    .ListItems.Item(i).SubItems(4) = "3"
                Else
                    .ListItems.Item(i).SubItems(4) = "1"

                End If

            Else

                If chkNotCheckBitOS.Value Then
                    .ListItems.Item(i).SubItems(4) = "2"
                Else
                    .ListItems.Item(i).SubItems(4) = "0"
                End If

            End If

            .ListItems.Item(i).SubItems(5) = ucPhysXPath.Path
            .ListItems.Item(i).SubItems(6) = ucLangPath.Path
            .ListItems.Item(i).SubItems(7) = ucRuntimesPath.Path
            .ListItems.Item(i).SubItems(8) = txtExcludeFileName

        End With

        'FRMOPTIONS
    End If

    LastIdOS = frmOptions.lvOS.ListItems.Count
    frmOptions.lvOS.Refresh
    mbAddInList = False

End Sub

Private Sub txtExcludeFileName_GotFocus()
    HighlightActiveControl Me, txtExcludeFileName, True
End Sub

Private Sub txtExcludeFileName_LostFocus()
    HighlightActiveControl Me, txtExcludeFileName, False
End Sub

Private Sub txtOSName_GotFocus()
    HighlightActiveControl Me, txtOSName, True
End Sub

Private Sub txtOSName_LostFocus()
    HighlightActiveControl Me, txtOSName, False
End Sub

Private Sub txtOSVer_GotFocus()
    HighlightActiveControl Me, txtOSVer, True
End Sub

Private Sub txtOSVer_LostFocus()
    HighlightActiveControl Me, txtOSVer, False
End Sub

Private Sub txtOSVer_Change()
    cmdOK.Enabled = LenB(Trim$(txtOSVer)) > 0 And LenB(Trim$(ucPathDRP.Path)) > 0
End Sub

Private Sub ucLangPath_Click()

Dim strTempPath                         As String

    If ucLangPath.FileCount > 0 Then
        strTempPath = ucLangPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)

        End If

    End If

    If LenB(strTempPath) > 0 Then
        ucLangPath.Path = strTempPath

    End If

End Sub

Private Sub ucLangPath_GotFocus()
    HighlightActiveControl Me, ucLangPath, True
End Sub

Private Sub ucLangPath_LostFocus()
    HighlightActiveControl Me, ucLangPath, False
End Sub

Private Sub ucPathDB_Click()

Dim strTempPath                         As String

    If ucPathDB.FileCount > 0 Then
        strTempPath = ucPathDB.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)

        End If

    End If

    If LenB(strTempPath) > 0 Then
        ucPathDB.Path = strTempPath

    End If

End Sub

Private Sub ucPathDB_GotFocus()
    HighlightActiveControl Me, ucPathDB, True
End Sub

Private Sub ucPathDB_LostFocus()
    HighlightActiveControl Me, ucPathDB, False
End Sub

Private Sub ucPathDRP_Click()

Dim strTempPath                         As String

    If ucPathDRP.FileCount > 0 Then
        strTempPath = ucPathDRP.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)

        End If

    End If

    If LenB(strTempPath) > 0 Then
        ucPathDRP.Path = strTempPath

    End If

End Sub

Private Sub ucPathDRP_GotFocus()
    HighlightActiveControl Me, ucPathDRP, True
End Sub

Private Sub ucPathDRP_LostFocus()
    HighlightActiveControl Me, ucPathDRP, False
End Sub

Private Sub ucPhysXPath_Click()

Dim strTempPath                         As String

    If ucPhysXPath.FileCount > 0 Then
        strTempPath = ucPhysXPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
        ucPhysXPath.Path = strTempPath

    End If

End Sub

Private Sub ucPhysXPath_GotFocus()
    HighlightActiveControl Me, ucPhysXPath, True
End Sub

Private Sub ucPhysXPath_LostFocus()
    HighlightActiveControl Me, ucPhysXPath, False
End Sub

Private Sub ucRuntimesPath_Click()

Dim strTempPath                         As String

    If ucRuntimesPath.FileCount > 0 Then
        strTempPath = ucRuntimesPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)

        End If

    End If

    If LenB(strTempPath) > 0 Then
        ucRuntimesPath.Path = strTempPath

    End If

End Sub

Private Sub ucRuntimesPath_GotFocus()
    HighlightActiveControl Me, ucRuntimesPath, True
End Sub

Private Sub ucRuntimesPath_LostFocus()
    HighlightActiveControl Me, ucRuntimesPath, False
End Sub
