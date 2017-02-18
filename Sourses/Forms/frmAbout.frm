VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "О программе..."
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9645
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.ctlJCbutton cmdHomePage 
      Height          =   650
      Left            =   7320
      TabIndex        =   4
      Top             =   5505
      Width           =   2205
      _ExtentX        =   3889
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
      Caption         =   "HomePage"
      CaptionEffects  =   0
      PictureAlign    =   0
      DropDownSymbol  =   6
      DropDownSeparator=   -1  'True
      DropDownEnable  =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdOsZoneNet 
      Height          =   650
      Left            =   4980
      TabIndex        =   3
      Top             =   5505
      Width           =   2205
      _ExtentX        =   3889
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
      Caption         =   "Обсуждение на OsZone.Net"
      CaptionEffects  =   0
      PictureAlign    =   0
   End
   Begin prjDIADBS.ctlJCbutton cmdCheckUpd 
      Height          =   650
      Left            =   1320
      TabIndex        =   5
      Top             =   6300
      Width           =   2205
      _ExtentX        =   3889
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
      Caption         =   "Проверить обновление..."
      CaptionEffects  =   0
      PictureAlign    =   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdLicence 
      Height          =   650
      Left            =   2460
      TabIndex        =   2
      Top             =   5505
      Width           =   2205
      _ExtentX        =   3889
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
      Caption         =   "Лицензионное соглашение"
      CaptionEffects  =   0
      PictureAlign    =   0
   End
   Begin prjDIADBS.ctlJCbutton cmdDonate 
      Height          =   650
      Left            =   120
      TabIndex        =   1
      Top             =   5505
      Width           =   2200
      _ExtentX        =   3889
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
      Caption         =   "Поддержать проект"
      CaptionEffects  =   0
      PictureAlign    =   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Default         =   -1  'True
      Height          =   650
      Left            =   6120
      TabIndex        =   0
      Top             =   6300
      Width           =   2205
      _ExtentX        =   3889
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
      Caption         =   "Закрыть"
      CaptionEffects  =   0
      PictureAlign    =   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton ctlAquaButton 
      Height          =   1995
      Left            =   75
      TabIndex        =   6
      Top             =   120
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   3519
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   10
      BackColor       =   16765357
      Caption         =   ""
      CaptionEffects  =   0
      PictureNormal   =   "frmAbout.frx":000C
      PictureShadow   =   -1  'True
   End
   Begin prjDIADBS.LabelW lblTranslator 
      Height          =   315
      Left            =   105
      TabIndex        =   10
      Top             =   2820
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   4
      BackStyle       =   0
      Caption         =   "Перевод программы: Головеев Роман"
   End
   Begin prjDIADBS.LabelW lblThanks 
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   3180
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   4
      BackStyle       =   0
      Caption         =   "Благодарности:"
      WordWrap        =   0   'False
   End
   Begin prjDIADBS.LabelW lblAuthor 
      Height          =   375
      Left            =   105
      TabIndex        =   9
      Top             =   2520
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "Автор программы: Головеев Роман"
   End
   Begin prjDIADBS.LabelW lblInfo 
      Height          =   1095
      Left            =   2280
      TabIndex        =   8
      Top             =   1440
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   1931
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "Описание программы"
   End
   Begin prjDIADBS.LabelW lblNameProg 
      Height          =   1305
      Left            =   2280
      TabIndex        =   7
      Top             =   45
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   2302
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Label1"
   End
   Begin prjDIADBS.LabelW lblMailTo 
      Height          =   240
      Left            =   105
      TabIndex        =   12
      Top             =   5160
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12582912
      MousePointer    =   4
      BackStyle       =   0
      Caption         =   "Написать E-mail автору программы"
   End
   Begin VB.Menu mnuContextMenu1 
      Caption         =   "Контекстное меню 1"
      Begin VB.Menu mnuContextLink 
         Caption         =   "Посетить сайт 1"
         Index           =   0
      End
      Begin VB.Menu mnuContextLink 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContextLink 
         Caption         =   "Посетить сайт 2"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strTranslatorName As String
Private strTranslatorUrl  As String
Private strFormName       As String
Private strCreditList_x() As String
Private lngCurCredit      As Long

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
'! Procedure   (Функция)   :   Sub cmdCheckUpd_Click
'! Description (Описание)  :   [Запуск формы проверки обновления]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdCheckUpd_Click()
    CheckUpd False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDonate_Click
'! Description (Описание)  :   [Запуск формы Donate]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdDonate_Click()
    frmDonate.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdExit_Click
'! Description (Описание)  :   [Выход из формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdHomePage_Click
'! Description (Описание)  :   [Переход на домашнюю страницу]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdHomePage_Click()
    RunUtilsShell strUrl_MainWWWSite, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdLicence_Click
'! Description (Описание)  :   [Показ лицензионного соглашения]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdLicence_Click()
    frmLicence.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdOsZoneNet_Click
'! Description (Описание)  :   [Переход на форум OsZone.net]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdOsZoneNet_Click()
    RunUtilsShell strUrlOsZoneNetThread, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ctlAquaButton1_Click
'! Description (Описание)  :   [Переход на сайт программы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ctlAquaButton_Click()
    RunUtilsShell strUrl_MainWWWSite, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FontCharsetChange
'! Description (Описание)  :   [Изменение шрифта формы]
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
'! Procedure   (Функция)   :   Sub Form_KeyDown
'! Description (Описание)  :   [Обработка нажатий клавиш клавиатуры]
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
'! Description (Описание)  :   [События при  загрузке формы]
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

    lblNameProg.Caption = strFrmMainCaptionTemp & vbNewLine & " v." & strProductVersion & vbNewLine & strFrmMainCaptionTempDate & strDateProgram & ")"

    LoadIconImage2Object cmdExit, "BTN_EXIT", strPathImageMainWork
    LoadIconImage2Object cmdDonate, "BTN_DONATE", strPathImageMainWork
    LoadIconImage2Object cmdCheckUpd, "BTN_UPDATE", strPathImageMainWork
    LoadIconImage2Object cmdHomePage, "BTN_HOME", strPathImageMainWork
    LoadIconImage2Object cmdOsZoneNet, "BTN_HOME", strPathImageMainWork
    LoadIconImage2Object cmdLicence, "BTN_LICENCE", strPathImageMainWork

    Select Case strPCLangCurrentID

        Case "0419"
            lblAuthor.Caption = "Автор программы: Головеев Роман aka Romeo91"
            lblThanks(0).Caption = "Мои благодарности:"
        Case Else
            lblAuthor.Caption = "Author of the program: Goloveev Roman (Romeo91)"
            lblThanks(0).Caption = "My thanks:"
    End Select

    mnuContextMenu1.Enabled = False
    
    ' Локализация приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If
    
    ' Создание списка благодарностей и показ на форме
    LoadThankYou
    
    ' Присваиваем меню для кнопки
    cmdHomePage.SetPopupMenu mnuContextMenu1
    mnuContextMenu1.Enabled = True
    
    ' Подсказки и свойства label
    With lblAuthor
        .MouseIcon = lblMailTo.MouseIcon
        .MousePointer = lblMailTo.MousePointer
        .ForeColor = lblMailTo.ForeColor
        .ToolTipText = strUrl_MainWWWSite
    End With
    lblMailTo.ToolTipText = "roman-novosib@ngs.ru"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Unload
'! Description (Описание)  :   [Выгрузка формы]
'! Parameters  (Переменные):   Cancel (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    cmdHomePage.UnsetPopupMenu
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub GenerateThankyou
'! Description (Описание)  :   [Генерация текста благодарности со ссылкой на страницу]
'!                              Idea from
'!                              Copyright ©2001-2013 by Tanner Helland
'!                              http://www.tannerhelland.com/photodemon
'! Parameters  (Переменные):   thxText (String)
'                              creditURL (String = vbNullString)
'!--------------------------------------------------------------------------------
Private Sub GenerateThankyou(ByVal thxText As String, Optional ByVal creditURL As String = vbNullString)
    'Because I now have too many people to thank, it's necessary to split the list into multiple columns
    Dim columnLimit As Long
    Dim thxOffset   As Long

    'Generate a new label
    Load lblThanks(lngCurCredit)
    
    columnLimit = 7
    thxOffset = 750

    With lblThanks(lngCurCredit)

        If lngCurCredit = 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 30 + thxOffset
        ElseIf lngCurCredit < columnLimit Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 30 + thxOffset
        ElseIf lngCurCredit = columnLimit Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60 - (lblThanks(columnLimit - 1).Top - lblThanks(0).Top)
            .Left = lblThanks(0).Left + 2700 + thxOffset
        ElseIf lngCurCredit < columnLimit * 2 - 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 2700 + thxOffset
        ElseIf lngCurCredit = columnLimit * 2 - 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60 - (lblThanks(columnLimit * 2 - 2).Top - lblThanks(0).Top)
            .Left = lblThanks(0).Left + 5400 + thxOffset
        ElseIf lngCurCredit < columnLimit * 3 - 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 5400 + thxOffset
        ElseIf lngCurCredit = columnLimit * 3 - 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60 - (lblThanks(columnLimit * 3 - 3).Top - lblThanks(0).Top)
            .Left = lblThanks(0).Left + 8100 + thxOffset
        Else
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 8100 + thxOffset
        End If

        .Caption = thxText

        If LenB(creditURL) = 0 Then
            .MousePointer = vbDefault
        Else
            .Font.Underline = True
            .MouseIcon = lblMailTo.MouseIcon
            .MousePointer = lblMailTo.MousePointer
            .ForeColor = lblMailTo.ForeColor
            .ToolTipText = creditURL
        End If

        .Visible = True
    End With

    ReDim Preserve strCreditList_x(0 To lngCurCredit)

    strCreditList_x(lngCurCredit) = creditURL
    lngCurCredit = lngCurCredit + 1
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lblAuthor_Click
'! Description (Описание)  :   [Переход на сайт разработчика]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub lblAuthor_Click()
    RunUtilsShell strUrl_MainWWWSite, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lblMailTo_MouseDown
'! Description (Описание)  :   [Нажатие мышкой на "Связаться с разработчиком"]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub lblMailTo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim strSubject As String
    
    If Button = vbLeftButton Then
        strSubject = "My wishes for the program (" & App.ProductName & ")"
        ShellExecute Me.hWnd, vbNullString, "mailto:Romeo91<roman-novosib@ngs.ru>?Subject=" & Replace$(strSubject, strSpace, "%20"), vbNullString, "c:\", 1
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lblThanks_Click
'! Description (Описание)  :   [When a thank-you credit is clicked, launch the corresponding website]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub lblThanks_Click(Index As Integer)

    If LenB(strCreditList_x(Index)) Then
        RunUtilsShell strCreditList_x(Index), False
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lblTranslator_MouseDown
'! Description (Описание)  :   [Переход на сайт переводчика, или отправка почты]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub lblTranslator_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If LenB(strTranslatorUrl) Then
        If Button = vbLeftButton Then
            RunUtilsShell strTranslatorUrl, False
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadThankYou
'! Description (Описание)  :   [Загрузка списка благодарностей]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadThankYou()
    lngCurCredit = 1
    
    GenerateThankyou "SamLab", "http://driveroff.net/"
    GenerateThankyou "OSzone.net forum's users", "http://forum.oszone.net/forum-62.html"
    ' Replacement CommonControls (TextBoxW, ListView, ComboBoxW, ListBoxW, ProgressBar, ToolTip, ImageList, OptionButtonW,RichTextBox, CheckBoxW, LabelW, SpinBox)
    GenerateThankyou "Krool", "http://www.vbforums.com/showthread.php?698563-CommonControls-(Replacement-of-the-MS-common-controls)"
    'JCbutton
    GenerateThankyou "Juned Chhipa", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71482&lngWId=1"
    'jcFrames
    GenerateThankyou "Juan Carlos San Roman", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64261&lngWId=1"
    'clsmenuimage, ScrollControl
    GenerateThankyou "Leandro Ascierto", "http://leandroascierto.com/blog/clsmenuimage/"
    GenerateThankyou "VBnet and Randy Birch", "http://vbnet.mvps.org/"
    'cmdparsing
    GenerateThankyou "EliteXP Software Solutions", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=72018&lngWId=1"
    'ucPickBox' ucStatusBar
    GenerateThankyou "Paul R.Territos", "http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=63905&lngWId=1"
    '[VB6] Function Wait (non-freezing & non-CPU-intensive)
    GenerateThankyou "Bonnie West", "http://www.vbforums.com/showthread.php?700373-VB6-Shell-amp-Wait"
    'Team HomeWork
    ' Timed MessageBox
    GenerateThankyou "Anirudha Vengurlekar", "anirudhav@yahoo.com"
    'AnimateForm
    GenerateThankyou "Jim Jose", "jimjosev33@yahoo.com"
    'MD5
    GenerateThankyou "Marcin Kleczynski", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69092&lngWId=1"
    'HighlightActiveControl
    GenerateThankyou "Giorgio Brausi", "http://nuke.vbcorner.net/"
    'Unicode String Array Sorting Class - cBlizzard.cls
    GenerateThankyou "Rohan Edwards aka Rd", "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=72576&lngWId=1"
    'LoadThankYou and other idea
    GenerateThankyou "Tanner Helland", "http://photodemon.org/"
    ' SortDMArray
    GenerateThankyou "Ellis Dee"
    GenerateThankyou "Zhu JinYong"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadTranslator
'! Description (Описание)  :   [Загрузка сведений о переводчике]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadTranslator()

    Select Case strPCLangCurrentID

        Case "0419"
            lblTranslator.Caption = "Перевод программы: " & strTranslatorName

        Case Else
            lblTranslator.Caption = "Translation of the program: " & strTranslatorName
    End Select

    If LenB(strTranslatorUrl) Then

        With lblTranslator
            .MouseIcon = lblMailTo.MouseIcon
            .MousePointer = lblMailTo.MousePointer
            .ForeColor = lblMailTo.ForeColor
            .ToolTipText = strTranslatorUrl
        End With

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [Загрузка данных локализации для компонентов]
'! Parameters  (Переменные):   strPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal strPathFile As String)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.CaptionW = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    '  Вызов основной функции для вывода Caption меню с поддержкой Unicode
    Call LocaliseMenu(strPathFile)
    'Кнопки
    cmdDonate.Caption = LocaliseString(strPathFile, strFormName, "cmdDonate", cmdDonate.Caption)
    cmdCheckUpd.Caption = LocaliseString(strPathFile, strFormName, "cmdCheckUpd", cmdCheckUpd.Caption)
    cmdLicence.Caption = LocaliseString(strPathFile, strFormName, "cmdLicence", cmdLicence.Caption)
    cmdHomePage.Caption = LocaliseString(strPathFile, strFormName, "cmdHomePage", cmdHomePage.Caption)
    cmdOsZoneNet.Caption = LocaliseString(strPathFile, strFormName, "cmdOsZoneNet", cmdOsZoneNet.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
    ' Лейблы
    lblMailTo.Caption = LocaliseString(strPathFile, strFormName, "lblMailTo", lblMailTo.Caption)
    lblInfo.Caption = LocaliseString(strPathFile, strFormName, "lblInfo", lblInfo.Caption)
    ' Перевод программы
    strTranslatorName = LocaliseString(strPathFile, "Lang", "TranslatorName", lblTranslator.Caption)
    strTranslatorUrl = LocaliseString(strPathFile, "Lang", "TranslatorUrl", vbNullString)
    LoadTranslator
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LocaliseMenu
'! Description (Описание)  :   [Загрузка данных локализации для меню]
'! Parameters  (Переменные):   strPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub LocaliseMenu(ByVal strPathFile As String)
    SetUniMenu -1, 0, -1, mnuContextMenu1, LocaliseString(strPathFile, strFormName, "cmdHomePage", cmdHomePage.Caption)
    SetUniMenu 0, 0, -1, mnuContextMenu1, LocaliseString(strPathFile, strFormName, "mnuContextLink1", mnuContextLink(0).Caption)
    SetUniMenu 0, 2, -1, mnuContextMenu1, LocaliseString(strPathFile, strFormName, "mnuContextLink2", mnuContextLink(2).Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextLink_Click
'! Description (Описание)  :   [Список действий для выпадающего меню кнопки]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuContextLink_Click(Index As Integer)

    Select Case Index

        Case 0
            RunUtilsShell strUrl_MainWWWSite, False

        Case 2
            RunUtilsShell strUrl_MainWWWForum, False
    End Select
    
End Sub
