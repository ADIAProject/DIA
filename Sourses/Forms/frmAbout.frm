VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "О программе..."
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   9405
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
   ScaleHeight     =   6330
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.ctlXpButton cmdSoftGetNet 
      Height          =   735
      Left            =   6225
      TabIndex        =   1
      Top             =   5500
      Width           =   1450
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
      Caption         =   "HomePage"
      PicturePosition =   0
      ButtonStyle     =   3
      PictureWidth    =   0
      PictureHeight   =   0
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      TextColor       =   0
      MenuCaption0    =   "#"
      MenuExist       =   -1  'True
   End
   Begin prjDIADBS.ctlXpButton cmdOsZoneNet 
      Height          =   735
      Left            =   4550
      TabIndex        =   2
      Top             =   5500
      Width           =   1575
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
      Caption         =   "Обсуждение на OsZone.Net"
      ButtonStyle     =   3
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin prjDIADBS.ctlXpButton cmdLicence 
      Height          =   735
      Left            =   2375
      TabIndex        =   3
      Top             =   5500
      Width           =   2050
      _ExtentX        =   3625
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
      Caption         =   "Лицензионное соглашение"
      ButtonStyle     =   3
      PictureWidth    =   48
      PictureHeight   =   48
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      TextColor       =   0
      MenuCaption0    =   "#"
   End
   Begin prjDIADBS.ctlXpButton cmdDonate 
      Height          =   735
      Left            =   100
      TabIndex        =   5
      Top             =   5500
      Width           =   2150
      _ExtentX        =   3784
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
      Caption         =   "Поддержать проект"
      ButtonStyle     =   3
      PictureWidth    =   51
      PictureHeight   =   28
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      TextColor       =   0
      MenuCaption0    =   "#"
   End
   Begin prjDIADBS.ctlXpButton cmdExit 
      Height          =   735
      Left            =   7800
      TabIndex        =   0
      Top             =   5500
      Width           =   1550
      _ExtentX        =   2725
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
      Caption         =   "Закрыть"
      ButtonStyle     =   3
      PictureWidth    =   0
      PictureHeight   =   0
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
   End
   Begin prjDIADBS.ctlJCbutton ctlAquaButton1 
      Height          =   1995
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   3519
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   ""
      PictureNormal   =   "frmAbout.frx":000C
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin prjDIADBS.LabelW lblTranslator 
      Height          =   255
      Left            =   105
      TabIndex        =   6
      Top             =   3175
      Width           =   9255
      _ExtentX        =   0
      _ExtentY        =   0
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
      Left            =   105
      TabIndex        =   7
      Top             =   3480
      Width           =   2500
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
      BackStyle       =   0
      Caption         =   "Благодарности:"
      AutoSize        =   -1  'True
      WordWrap        =   0   'False
   End
   Begin prjDIADBS.LabelW lblAuthor 
      Height          =   255
      Left            =   105
      TabIndex        =   8
      Top             =   2880
      Width           =   9255
      _ExtentX        =   0
      _ExtentY        =   0
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
      Height          =   1335
      Left            =   2220
      TabIndex        =   9
      Top             =   1560
      Width           =   7155
      _ExtentX        =   0
      _ExtentY        =   0
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
      Height          =   1575
      Left            =   2220
      TabIndex        =   10
      Top             =   45
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   2778
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Label1"
   End
   Begin prjDIADBS.LabelW lblMailTo 
      Height          =   255
      Left            =   105
      TabIndex        =   11
      Top             =   5160
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   450
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
      Caption         =   "Написать E-mail автору программу"
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strTranslatorName               As String
Private strTranslatorUrl                As String
Private strFormName                     As String

Private strCreditList() As String
Private lngCurCredit As Long

Private Sub FontCharsetChange()
' Выставляем шрифт
    With Me.Font
        .Name = strOtherForm_FontName
        .Size = lngOtherForm_FontSize
        .Charset = lngDialog_Charset
    End With

    SetButtonProperties cmdDonate
    SetButtonProperties cmdLicence
    SetButtonProperties cmdOsZoneNet
    SetButtonProperties cmdSoftGetNet
    SetButtonProperties cmdExit

End Sub

Private Sub cmdDonate_Click()
    frmDonate.Show vbModal, Me
End Sub

'! -----------------------------------------------------------
'!  Функция     :  cmdExit_Click
'!  Переменные  :
'!  Описание    :  Выход из формы
'! -----------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLicence_Click()
    frmLicence.Show vbModal, Me
End Sub

Private Sub cmdOsZoneNet_Click()

Dim cmdString                           As String
Dim nRetShellEx                         As Boolean

    cmdString = Kavichki & "http://forum.oszone.net/thread-139908.html" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx

End Sub

Private Sub cmdSoftGetNet_Click()

Dim cmdString                           As String
Dim nRetShellEx                         As Boolean

    cmdString = Kavichki & "http://www.adia-project.net" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx

End Sub

Private Sub cmdSoftGetNet_ClickMenu(mnuIndex As Integer)

Dim nRetShellEx                         As Boolean
Dim cmdString                           As String

    Select Case mnuIndex

        Case 0
            cmdString = Kavichki & "http://www.adia-project.net" & Kavichki

        Case 2
            cmdString = Kavichki & "http://www.adia-project.net/forum/index.php" & Kavichki

    End Select

    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx

End Sub

Private Sub ctlAquaButton1_Click()

Dim cmdString                           As String
Dim nRetShellEx                         As Boolean

    cmdString = Kavichki & "http://www.adia-project.net" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx

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
'!  Описание    :  События при  загрузке формы
'! -----------------------------------------------------------
Private Sub Form_Load()

    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, "FRMABOUT", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    LoadIconImage2Btn cmdExit, "BTN_EXIT", strPathImageMainWork
    LoadIconImage2Btn cmdDonate, "BTN_DONATE", strPathImageMainWork
    LoadIconImage2Btn cmdLicence, "BTN_LICENCE", strPathImageMainWork
    lblNameProg.Caption = strFrmMainCaptionTemp & vbNewLine & " v." & strProductVersion & vbNewLine & strFrmMainCaptionTempDate & strDateProgram & ")"

    Select Case strPCLangCurrentID
        Case "0419"
            lblAuthor.Caption = "Автор программы: Головеев Роман aka Romeo91"
            lblThanks(0).Caption = "Мои благодарности:" '& vbNewLine & _
            "* Участникам форума сайта OSZONE.NET за помощь в тестировании и за помощь в развитии проекта" & vbNewLine & _
            "* Всем остальным пользователям, которые помогли сделать эту программу лучше (за поиск ошибок, за идеи развития проекта, за критику)" & vbNewLine & _
            "* Всем, кто бескорыстно поддерживает проект - морально и финансово" & vbNewLine & _
            lblThanks(0).Caption = "Мои благодарности:"
            '& vbNewLine & "* Участникам форума сайта OSZONE.NET за помощь в тестировании и за помощь в развитии проекта" & vbNewLine & "* Всем остальным пользователям, которые помогли сделать эту программу лучше (за поиск ошибок, за идеи развития проекта, за критику)" & vbNewLine & "* Всем, кто бескорыстно поддерживает проект - морально и финансово" & vbNewLine & lblThanks(0).Caption = "Мои благодарности:"
            '& vbNewLine & "* Участникам форума сайта OSZONE.NET за помощь в тестировании и за помощь в развитии проекта" & vbNewLine & "* Всем остальным пользователям, которые помогли сделать эту программу лучше (за поиск ошибок, за идеи развития проекта, за критику)" & vbNewLine & "* Всем, кто бескорыстно поддерживает проект - морально и финансово" & vbNewLine & "* Также огромное спасибо Александру Дровосекову (apexsun.narod.ru) - в программе использованы, написанных когда-то им, элементы управления (User Control)"

        Case Else
            lblAuthor.Caption = "Author of the program: Goloveev Roman (Romeo91)"
            lblThanks(0).Caption = "My thanks:" '& vbNewLine & _
            "* The Users of the forum of the site OSZONE.NET for help in testing and for help in development of the project" & vbNewLine & _
            "* All rest user, which helped to do this program better (for searching for error, for ideas of the development of the project, for critic)" & vbNewLine & _
            "* All, who unselfish supports project - morally and financial" & vbNewLine & _
            lblThanks(0).Caption = "My thanks:"
            '& vbNewLine & "* The Users of the forum of the site OSZONE.NET for help in testing and for help in development of the project" & vbNewLine & "* All rest user, which helped to do this program better (for searching for error, for ideas of the development of the project, for critic)" & vbNewLine & "* All, who unselfish supports project - morally and financial" & vbNewLine & lblThanks(0).Caption = "My thanks:"
            '& vbNewLine & "* The Users of the forum of the site OSZONE.NET for help in testing and for help in development of the project" & vbNewLine & "* All rest user, which helped to do this program better (for searching for error, for ideas of the development of the project, for critic)" & vbNewLine & "* All, who unselfish supports project - morally and financial" & vbNewLine & "* Also big thank to Alexander Drovosekov (apexsun.narod.ru) - in program are used, written at one time him, elements of control (User Control)"
    End Select

    With cmdSoftGetNet
        If .MenuExist Then
            If .MenuCount = 0 Then
                .AddMenu "Site"
                .AddMenu "-"
                .AddMenu "Forum"
            End If
        End If
    End With

    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If
    
    LoadThankYou
    
End Sub

'! -----------------------------------------------------------
'!  Функция     :  lblMailTo_MouseDown
'!  Переменные  :  Button As Integer, Shift As Integer,X As Single, Y As Single
'!  Описание    :  Нажатие мышкой на "Связаться с разработчиком"
'! -----------------------------------------------------------
Private Sub lblMailTo_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    If Button = vbLeftButton Then
        ShellExecute Me.hWnd, vbNullString, "mailto:Romeo91<roman-novosib@ngs.ru>?Subject=My%20wish%20for%20update%20program%20(Drivers%20Installer%20Assistant)", vbNullString, "c:\", 1
    End If

End Sub


Private Sub lblTranslator_MouseDown(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

Dim cmdString                           As String
Dim nRetShellEx                         As Boolean

    If LenB(strTranslatorUrl) > 0 Then
        If Button = vbLeftButton Then
            RunUtilsShell Kavichki & strTranslatorUrl, False
        End If

    End If

End Sub

Private Sub LoadTranslator()

    Select Case strPCLangCurrentID

        Case "0419"
            lblTranslator.Caption = "Перевод программы: " & strTranslatorName

        Case Else
            lblTranslator.Caption = "Translation of the program: " & strTranslatorName

    End Select

    If LenB(strTranslatorUrl) > 0 Then
        With lblTranslator
            .MouseIcon = lblMailTo.MouseIcon
            .MousePointer = lblMailTo.MousePointer
            .ForeColor = lblMailTo.ForeColor
        End With
    End If

End Sub

Private Sub Localise(ByVal strPathFile As String)
' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.Caption = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    'Кнопки
    cmdDonate.Caption = LocaliseString(strPathFile, strFormName, "cmdDonate", cmdDonate.Caption)
    cmdLicence.Caption = LocaliseString(strPathFile, strFormName, "cmdLicence", cmdLicence.Caption)
    cmdSoftGetNet.Caption = LocaliseString(strPathFile, strFormName, "cmdSoftGetNet", cmdSoftGetNet.Caption)
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

Private Sub LoadThankYou()
    lngCurCredit = 1
    GenerateThankyou "SamLab", "http://driveroff.net/"
    GenerateThankyou "OSzone.net forum's users", "http://forum.oszone.net/forum-62.html"
    GenerateThankyou "Krool", "http://www.vbforums.com/showthread.php?698563-CommonControls-(Replacement-of-the-MS-common-controls)"
    GenerateThankyou "Juned Chhipa", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71482&lngWId=1"
    GenerateThankyou "Leandro Ascierto", "http://leandroascierto.com/blog/clsmenuimage/"
    GenerateThankyou "VBnet and Randy Birch", "http://vbnet.mvps.org/"
    ' win7Toolbar
    GenerateThankyou "AndRAY (Makarov Andrey)", "http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=72856&lngWId=1"
    'cmdparsing
    GenerateThankyou "EliteXP Software Solutions", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=72018&lngWId=1"
    'ucPickBox' ucStatusBar
    GenerateThankyou "Paul R.Territos", "http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=63905&lngWId=1"
    '[VB6] Function Wait (non-freezing & non-CPU-intensive)
    GenerateThankyou "Bonnie West", "http://www.vbforums.com/showthread.php?700373-VB6-Shell-amp-Wait"
    'Team HomeWork
     ' Timed MessageBox
    GenerateThankyou "Anirudha Vengurlekar"
     ' SortDMArray
    GenerateThankyou "Ellis Dee"
    GenerateThankyou "Zhu JinYong"
    
    'AnimateForm - Jim Jose
End Sub

'Генерация текста благодарности со ссылкой на страницу
'Idea from
'Copyright ©2001-2013 by Tanner Helland
'http://www.tannerhelland.com/photodemon
Private Sub GenerateThankyou(ByVal thxText As String, Optional ByVal creditURL As String = vbNullString)
    
    'Generate a new label
    Load lblThanks(lngCurCredit)
    
    'Because I now have too many people to thank, it's necessary to split the list into multiple columns
    Dim columnLimit As Long
    columnLimit = 5
    
    Dim thxOffset As Long
    thxOffset = 750
    
    With lblThanks(lngCurCredit)
        If lngCurCredit = 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 300
            .Left = lblThanks(0).Left + 30 + thxOffset
        ElseIf lngCurCredit < columnLimit Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 30 + thxOffset
        ElseIf lngCurCredit = columnLimit Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 300 - (lblThanks(columnLimit - 1).Top - lblThanks(0).Top)
            .Left = lblThanks(0).Left + 2700 + thxOffset
        ElseIf lngCurCredit < columnLimit * 2 - 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 2700 + thxOffset
        ElseIf lngCurCredit = columnLimit * 2 - 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 300 - (lblThanks(columnLimit * 2 - 2).Top - lblThanks(0).Top)
            .Left = lblThanks(0).Left + 5400 + thxOffset
        Else
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 5400 + thxOffset
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
    
    ReDim Preserve strCreditList(0 To lngCurCredit) As String
    strCreditList(lngCurCredit) = creditURL
    
    lngCurCredit = lngCurCredit + 1

End Sub

'When a thank-you credit is clicked, launch the corresponding website
Private Sub lblThanks_Click(Index As Integer)
    If LenB(strCreditList(Index)) Then
        RunUtilsShell Kavichki & strCreditList(Index) & Kavichki, False
    End If
End Sub

