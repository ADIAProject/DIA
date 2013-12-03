VERSION 5.00
Begin VB.Form frmListHwidAll 
   Caption         =   "Список всех устройств вашего компьютера"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   720
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListHwidAll.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   11760
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   750
      Left            =   9120
      TabIndex        =   0
      Top             =   5580
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1323
      ButtonStyle     =   8
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
      Caption         =   "Выход"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdReNewHW 
      Height          =   750
      Left            =   6480
      TabIndex        =   1
      Top             =   5595
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1323
      ButtonStyle     =   8
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
      Caption         =   "Обновить конфигурацию оборудования"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdBackUpDrivers 
      Height          =   750
      Left            =   3840
      TabIndex        =   2
      Top             =   5595
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   1323
      ButtonStyle     =   8
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
      Caption         =   "Создать резервную копию драйверов"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdCheckAll 
      Height          =   300
      Left            =   60
      TabIndex        =   3
      Top             =   5580
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      ButtonStyle     =   8
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
      Caption         =   "Выделить всё"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdUnCheckAll 
      Height          =   300
      Left            =   60
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      ButtonStyle     =   8
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
      Caption         =   "Снять выделение"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCFrames frGroup 
      Height          =   1250
      Left            =   75
      Top             =   40
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   2196
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15783104
      FillColor       =   15783104
      TextBoxColor    =   11595760
      Style           =   3
      RoundedCorner   =   0   'False
      Caption         =   "Группы драйверов:"
      Alignment       =   0
      HeaderStyle     =   1
      Begin prjDIADBS.OptionButtonW optGrp4 
         Height          =   345
         Left            =   1860
         TabIndex        =   8
         Top             =   780
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
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
         Caption         =   "frmListHwidAll.frx":000C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optGrp3 
         Height          =   405
         Left            =   1860
         TabIndex        =   7
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmListHwidAll.frx":0054
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW optGrp1 
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmListHwidAll.frx":007A
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW optGrp2 
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmListHwidAll.frx":00AC
         Transparent     =   -1  'True
      End
   End
   Begin prjDIADBS.ctlJCFrames frFindDrvInternet 
      Height          =   1250
      Left            =   4200
      Top             =   40
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   2196
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15783104
      FillColor       =   15783104
      TextBoxColor    =   11595760
      Style           =   3
      RoundedCorner   =   0   'False
      Caption         =   "Поиск драйвера в Интернете:"
      Alignment       =   0
      HeaderStyle     =   1
      Begin prjDIADBS.TextBoxW txtFindText 
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   360
         Width           =   5355
         _ExtentX        =   9446
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
         Text            =   "frmListHwidAll.frx":00D2
         CueBanner       =   "frmListHwidAll.frx":00F2
      End
      Begin prjDIADBS.CheckBoxW chkParseHwid 
         Height          =   210
         Left            =   60
         TabIndex        =   9
         Top             =   960
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmListHwidAll.frx":0112
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optDevID 
         Height          =   405
         Left            =   60
         TabIndex        =   10
         Top             =   600
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   714
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
         Caption         =   "frmListHwidAll.frx":0184
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optMS 
         Height          =   405
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   714
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
         Caption         =   "frmListHwidAll.frx":01B8
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optGoogle 
         Height          =   405
         Left            =   3660
         TabIndex        =   12
         Top             =   600
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmListHwidAll.frx":01F2
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdGoSite 
         Height          =   795
         Left            =   5520
         TabIndex        =   14
         Top             =   360
         Width           =   1905
         _ExtentX        =   5318
         _ExtentY        =   688
         ButtonStyle     =   8
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
         Caption         =   "Выделить"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
   End
   Begin prjDIADBS.ListView lvDevices 
      Height          =   1455
      Left            =   60
      TabIndex        =   15
      Top             =   1320
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   2566
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icons           =   "frmListHwidAll.frx":0226
      SmallIcons      =   "frmListHwidAll.frx":0252
      ColumnHeaderIcons=   "frmListHwidAll.frx":027E
      View            =   3
      Arrange         =   1
      AllowColumnReorder=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      LabelEdit       =   2
      HideSelection   =   0   'False
      ShowLabelTips   =   -1  'True
      HoverSelection  =   -1  'True
      HotTracking     =   -1  'True
      HighlightHot    =   -1  'True
      TextBackground  =   1
   End
   Begin prjDIADBS.LabelW lblWait 
      Height          =   375
      Left            =   105
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   11640
      _ExtentX        =   17383
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Идет обновление конфигурации оборудования. Пожалуйста, подождите...."
   End
   Begin prjDIADBS.LabelW lblInformation 
      Height          =   795
      Left            =   2160
      TabIndex        =   17
      Top             =   5520
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1402
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
      Caption         =   "Выберите драйвера необходимые для бекапирования и нажмите кнопку 'BackUp''. "
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Контекстное меню"
      Begin VB.Menu mnuContextProperties 
         Caption         =   "Показать свойства драйвера"
      End
      Begin VB.Menu mnuContextDelete 
         Caption         =   "Удалить драйвер"
      End
   End
End
Attribute VB_Name = "frmListHwidAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Минимальные размеры формы
Private lngFormWidthMin                 As Long
Private lngFormHeightMin                As Long

Private lngDeviceCount                  As Long

Private strFormName                     As String

Private Sub FontCharsetChange()
' Выставляем шрифт
    With Me.Font
        .Name = strOtherForm_FontName
        .Size = lngOtherForm_FontSize
        .Charset = lngDialog_Charset
    End With

    frGroup.Font.Charset = lngDialog_Charset
    frFindDrvInternet.Font.Charset = lngDialog_Charset

    SetButtonProperties , cmdReNewHW, True
    SetButtonProperties , cmdBackUpDrivers, True
    SetButtonProperties , cmdOK, True
    SetButtonProperties , cmdCheckAll, True
    SetButtonProperties , cmdUnCheckAll, True
    SetButtonProperties , cmdGoSite, True
End Sub

Private Sub cmdBackUpDrivers_Click()

Dim lngMsgRet                           As Long

    lngMsgRet = MsgBox(strMessages(123), vbYesNo + vbQuestion, strProductName)

    Select Case lngMsgRet

        Case vbYes
            RunUtilsShell Kavichki & "http://www.adia-project.net" & Kavichki, False

    End Select

End Sub

Private Sub cmdCheckAll_Click()

Dim i                                   As Integer

    With lvDevices.ListItems

        For i = 1 To .Count

            If Not .Item(i).Checked Then
                .Item(i).Checked = True

            End If

        Next
    End With

    'LVDEVICES
    FindCheckCountList

End Sub

'найти драйвер для выделенного устройства
Private Sub cmdGoSite_Click()

Dim strDevID                            As String
Dim cmdString                           As String
Dim nRetShellEx                         As Boolean

    strDevID = txtFindText.Text
    strDevID = Replace$(strDevID, "\", "%5C", , , vbTextCompare)
    strDevID = Replace$(strDevID, "&", "%26", , , vbTextCompare)

    If optDevID.Value Then
        cmdString = Kavichki & "http://www.devid.info/search.php?text=" & strDevID & "&=" & Kavichki
    ElseIf optGoogle.Value Then
        cmdString = Kavichki & "http://www.google.com/search?q=driver+" & strDevID & "&=" & Kavichki
    Else
        cmdString = Kavichki & "http://catalog.update.microsoft.com/v7/site/Search.aspx?q=" & strDevID & "&=" & Kavichki
    End If

    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx

End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub cmdReNewHW_Click()
    BlockControl False

    lvDevices.ListItems.Clear
    lvDevices.Visible = False
    lblWait.Visible = True

    ReCollectHWID
    ' Провести повторно поиск драйверов в пакетах, т.е обновить кнопки
    frmMain.UpdateStatusButtonAll
    ' Обновить список неизвестных дров и описание для кнопки
    frmMain.LoadCmdViewAllDeviceCaption
    SaveHWIDs2File

    'загружаем список по новой
    LoadListbyMode
    lblWait.Visible = False
    lvDevices.Visible = True

    BlockControl True

End Sub

'! -----------------------------------------------------------
'!  Функция     :  BlockControl
'!  Переменные  :
'!  Описание    :  Блокировка(Разблокировка) некоторых элементов формы при работе сложных функций
'! -----------------------------------------------------------
Public Sub BlockControl(ByVal mbBlock As Boolean)
'frGroup
    optGrp1.Enabled = mbBlock
    optGrp2.Enabled = mbBlock
    optGrp3.Enabled = mbBlock
    optGrp4.Enabled = mbBlock
    'frFindDrvInternet.Enabled = mbBlock
    cmdBackUpDrivers.Enabled = mbBlock
    cmdOK.Enabled = mbBlock
    cmdReNewHW.Enabled = mbBlock
End Sub

Private Sub cmdUnCheckAll_Click()

Dim i                                   As Integer

    With lvDevices.ListItems
        For i = 1 To .Count
            If .Item(i).Checked Then
                .Item(i).Checked = False
            End If
        Next
    End With

    FindCheckCountList

End Sub

Private Sub FindCheckCountList()

Dim i                                   As Integer
Dim miCount                             As Integer

    For i = 1 To lvDevices.ListItems.Count

        If lvDevices.ListItems.Item(i).Checked Then
            miCount = miCount + 1
        End If
    Next

    If miCount > 0 Then
        With cmdOK
            If Not .Enabled Then
                .Enabled = True
            End If

        End With
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me

    End If

End Sub

Private Sub Form_Load()
    SetupVisualStyles Me


    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmListHwidAll", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
        lngFormWidthMin = .Width
        lngFormHeightMin = .Height
    End With

    mnuContext.Visible = False

    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdCheckAll, "BTN_CHECKMARK", strPathImageMainWork
    LoadIconImage2BtnJC cmdUnCheckAll, "BTN_UNCHECKMARK", strPathImageMainWork
    LoadIconImage2BtnJC cmdGoSite, "BTN_VIEW_SEARCH", strPathImageMainWork
    LoadIconImage2BtnJC cmdReNewHW, "BTN_RENEWHW", strPathImageMainWork
    LoadIconImage2BtnJC cmdBackUpDrivers, "BTN_BACKUP", strPathImageMainWork

    ' все остальные процедуры
    FormLoadDefaultParam
    FormLoadAction

End Sub

Public Sub FormLoadDefaultParam()
    If Not (lvDevices Is Nothing) Then
        lvDevices.ColumnHeaders.Clear
        lvDevices.ListItems.Clear
    End If
    optGrp1.Value = 0
    'uncheck
    optGrp2.Value = 1
    'check
    optGrp3.Value = False
    optGrp4.Value = True
End Sub

Public Sub FormLoadAction()

' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    LoadList_Device False
    Me.Caption = Me.Caption & " (Find: " & lvDevices.ListItems.Count & ")"
    lngDeviceCount = lvDevices.ListItems.Count

    LoadListbyMode
    LoadFormCaption
    cmdGoSite.Enabled = LenB(txtFindText.Text) > 0
End Sub

Private Sub LoadFormCaption()
Dim MeCaptionView                       As String

    MeCaptionView = LocaliseString(strPCLangCurrentPath, strFormName, strFormName, Me.Caption)
    Me.Caption = MeCaptionView & " (" & lvDevices.ListItems.Count & " " & strMessages(124) & " " & lngDeviceCount & ")"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Выгружаем из памяти форму и другие компоненты
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
    Else
        Set frmListHwidAll = Nothing
    End If
End Sub

Private Sub Form_Resize()

Dim lngLVHeight                         As Long
Dim lngLVWidht                          As Long
Dim lngLVTop                            As Long

    On Error Resume Next

    With Me

        If .WindowState <> vbMinimized Then

            Dim miDeltaFrm              As Long

            If OsCurrVersionStruct.VerFull >= "6.0" Then
                miDeltaFrm = 125
            Else

                If mbAppThemed Then
                    miDeltaFrm = 0
                Else
                    miDeltaFrm = 0
                End If
            End If

            If .Width < lngFormWidthMin Then
                .Width = lngFormWidthMin
                .Enabled = False
                .Enabled = True
                Exit Sub
            End If

            If .Height < lngFormHeightMin Then
                .Height = lngFormHeightMin
                .Enabled = False
                .Enabled = True
                Exit Sub
            End If

            frFindDrvInternet.Width = .Width - frGroup.Left - frGroup.Width - 250 - miDeltaFrm

            With cmdGoSite
                .Left = frFindDrvInternet.Width - .Width - 125
                txtFindText.Width = .Left - 200
            End With

            'cmdGoSite
            cmdOK.Left = .Width - cmdOK.Width - 200 - miDeltaFrm
            cmdOK.Top = .Height - cmdOK.Height - 550 - miDeltaFrm
            cmdReNewHW.Top = cmdOK.Top
            cmdReNewHW.Left = .Width - cmdOK.Width - 250 - cmdReNewHW.Width - 150 - miDeltaFrm
            cmdBackUpDrivers.Top = cmdOK.Top
            cmdBackUpDrivers.Left = .Width - cmdOK.Width - 250 - cmdReNewHW.Width - 250 - cmdBackUpDrivers.Width - 100 - miDeltaFrm
            lngLVTop = (frGroup.Top + frGroup.Height) + 5 * Screen.TwipsPerPixelX
            lngLVHeight = cmdOK.Top - lngLVTop - 10 * Screen.TwipsPerPixelX
            lngLVWidht = (.Width - miDeltaFrm) - 18 * Screen.TwipsPerPixelX

            If Not (lvDevices Is Nothing) Then
                lvDevices.Move 5 * Screen.TwipsPerPixelX, lngLVTop, lngLVWidht, lngLVHeight
                lvDevices.Refresh
            End If

            lblWait.Left = 100
            lblWait.Width = .Width - 200

        End If

    End With

    'Me
End Sub

'! -----------------------------------------------------------
'!  Функция     :  LoadList_Device
'!  Переменные  :
'!  Описание    :  Построение полного спиcка устройств
'! -----------------------------------------------------------
Private Sub LoadList_Device(Optional ByVal mbViewed As Boolean = True, _
                            Optional ByVal lngMode As Long = 0)

Dim strDevHwid                          As String
Dim strDevDriverLocal                   As String
Dim strDevStatus                        As String
Dim strDevName                          As String
Dim strProvider                         As String
Dim strCompatID                         As String
Dim strStrDescription                   As String
Dim strOrigHwid                         As String
Dim ii                                  As Integer
Dim strInDPacks                         As String
Dim lngNumRow                           As Long

    With lvDevices
        .ListItems.Clear

        If .ColumnHeaders.Count = 0 Then
            .ColumnHeaders.Add 1, , strTableHwidHeader1, 225 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 2, , strTableHwidHeader7, 300 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 3, , strTableHwidHeader6, 60 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 4, , strTableHwidHeader5, 150 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 5, , strTableHwidHeader10, 150 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 6, , strTableHwidHeader11, 150 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 7, , strTableHwidHeader12, 250 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 8, , strTableHwidHeader14, 250 * Screen.TwipsPerPixelX
        End If

    End With

    For ii = LBound(arrHwidsLocal) To UBound(arrHwidsLocal)
        strDevHwid = arrHwidsLocal(ii).HWID
        strDevName = arrHwidsLocal(ii).DevName
        strDevStatus = arrHwidsLocal(ii).Status
        strDevDriverLocal = arrHwidsLocal(ii).VerLocal
        strProvider = arrHwidsLocal(ii).Provider
        strCompatID = arrHwidsLocal(ii).HWIDCompat
        strStrDescription = arrHwidsLocal(ii).Description

        If StrComp(strDevName, strStrDescription, vbTextCompare) <> 0 Then
            If LenB(strStrDescription) > 0 Then
                If InStr(strStrDescription, "unknown") = 0 Then
                    strDevName = strStrDescription
                End If
            End If
        End If

        strInDPacks = arrHwidsLocal(ii).DPsList
        strOrigHwid = arrHwidsLocal(ii).HWIDOrig

        Select Case lngMode
                ' All - ALL
            Case 0, 3

                With lvDevices.ListItems.Add(, , strDevHwid)
                    .SubItems(1) = strDevName
                    .SubItems(2) = strDevStatus
                    .SubItems(3) = strDevDriverLocal
                    .SubItems(4) = strProvider
                    .SubItems(5) = strCompatID
                    .SubItems(6) = strOrigHwid
                    .SubItems(7) = strInDPacks
                End With

                ' Microsoft - All
            Case 1

                If InStr(1, strProvider, "microsoft", vbTextCompare) Or _
                   InStr(1, strProvider, "майкрософт", vbTextCompare) Or _
                   InStr(1, strProvider, "standard", vbTextCompare) Then

                    With lvDevices.ListItems.Add(, , strDevHwid)
                        .SubItems(1) = strDevName
                        .SubItems(2) = strDevStatus
                        .SubItems(3) = strDevDriverLocal
                        .SubItems(4) = strProvider
                        .SubItems(5) = strCompatID
                        .SubItems(6) = strOrigHwid
                        .SubItems(7) = strInDPacks
                    End With

                    lngNumRow = lngNumRow + 1

                End If

                ' OEM - All
            Case 2

                If InStr(1, strProvider, "microsoft", vbTextCompare) = 0 And _
                   InStr(1, strProvider, "майкрософт", vbTextCompare) = 0 And _
                   InStr(1, strProvider, "standard", vbTextCompare) = 0 Then

                    With lvDevices.ListItems.Add(, , strDevHwid)
                        .SubItems(1) = strDevName
                        .SubItems(2) = strDevStatus
                        .SubItems(3) = strDevDriverLocal
                        .SubItems(4) = strProvider
                        .SubItems(5) = strCompatID
                        .SubItems(6) = strOrigHwid
                        .SubItems(7) = strInDPacks
                    End With

                    lngNumRow = lngNumRow + 1

                End If

                ' All - not in base
            Case 4

                If LenB(strInDPacks) = 0 Then

                    With lvDevices.ListItems.Add(, , strDevHwid)
                        .SubItems(1) = strDevName
                        .SubItems(2) = strDevStatus
                        .SubItems(3) = strDevDriverLocal
                        .SubItems(4) = strProvider
                        .SubItems(5) = strCompatID
                        .SubItems(6) = strOrigHwid
                        .SubItems(7) = strInDPacks
                    End With

                    lngNumRow = lngNumRow + 1

                End If

                ' Microsoft - not in base
            Case 5

                If InStr(1, strProvider, "microsoft", vbTextCompare) Or _
                   InStr(1, strProvider, "майкрософт", vbTextCompare) Or _
                   InStr(1, strProvider, "standard", vbTextCompare) Then
                    If LenB(strInDPacks) = 0 Then

                        With lvDevices.ListItems.Add(, , strDevHwid)
                            .SubItems(1) = strDevName
                            .SubItems(2) = strDevStatus
                            .SubItems(3) = strDevDriverLocal
                            .SubItems(4) = strProvider
                            .SubItems(5) = strCompatID
                            .SubItems(6) = strOrigHwid
                            .SubItems(7) = strInDPacks
                        End With

                        lngNumRow = lngNumRow + 1

                    End If

                End If

                ' OEM - not in base
            Case 6

                If InStr(1, strProvider, "microsoft", vbTextCompare) = 0 And _
                   InStr(1, strProvider, "майкрософт", vbTextCompare) = 0 And _
                   InStr(1, strProvider, "standard", vbTextCompare) = 0 Then
                    If LenB(strInDPacks) = 0 Then

                        With lvDevices.ListItems.Add(, , strDevHwid)
                            .SubItems(1) = strDevName
                            .SubItems(2) = strDevStatus
                            .SubItems(3) = strDevDriverLocal
                            .SubItems(4) = strProvider
                            .SubItems(5) = strCompatID
                            .SubItems(6) = strOrigHwid
                            .SubItems(7) = strInDPacks
                        End With

                        lngNumRow = lngNumRow + 1

                    End If

                End If

        End Select

    Next

End Sub

Private Sub LoadListbyMode()

Dim lngModeList                         As Long

Dim mbOpt1                              As Boolean
Dim mbOpt2                              As Boolean
Dim mbOpt3                              As Boolean
Dim mbOpt4                              As Boolean

    mbOpt1 = CBool(optGrp1.Value)
    mbOpt2 = CBool(optGrp2.Value)
    mbOpt3 = optGrp3.Value
    mbOpt4 = optGrp4.Value

    ' Microsoft
    If mbOpt1 And Not mbOpt2 Then
        'All
        If mbOpt3 Then
            lngModeList = 1
            'NotInBase
        Else
            lngModeList = 5

        End If

        ' OEM
    ElseIf Not mbOpt1 And mbOpt2 Then
        'All
        If mbOpt3 Then
            lngModeList = 2
            'NotInBase
        Else
            lngModeList = 6

        End If
        ' Ничего
    ElseIf Not mbOpt1 And Not mbOpt2 Then
        lngModeList = 9999

        ' Microsoft+OEM
    Else

        'All
        If mbOpt3 Then
            lngModeList = 3
            'NotInBase
        Else
            lngModeList = 4
        End If

    End If

    If Not (lvDevices Is Nothing) Then
        lvDevices.ListItems.Clear

    End If

    If lngModeList <> 9999 Then
        LoadList_Device False, lngModeList
    End If

    LoadFormCaption

    FindCheckCountList

End Sub

Private Sub Localise(StrPathFile As String)

' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.Caption = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
    'Кнопки
    cmdOK.Caption = LocaliseString(StrPathFile, strFormName, "cmdOK", cmdOK.Caption)
    frGroup.Caption = LocaliseString(StrPathFile, strFormName, "frGroup", frGroup.Caption)
    frFindDrvInternet.Caption = LocaliseString(StrPathFile, strFormName, "frFindDrvInternet", frFindDrvInternet.Caption)
    chkParseHwid.Caption = LocaliseString(StrPathFile, strFormName, "chkParseHwid", chkParseHwid.Caption)
    cmdGoSite.Caption = LocaliseString(StrPathFile, strFormName, "cmdGoSite", cmdGoSite.Caption)
    optGrp1.Caption = LocaliseString(StrPathFile, strFormName, "optGrp1", optGrp1.Caption)
    optGrp2.Caption = LocaliseString(StrPathFile, strFormName, "optGrp2", optGrp2.Caption)
    optGrp3.Caption = LocaliseString(StrPathFile, strFormName, "optGrp3", optGrp3.Caption)
    optGrp4.Caption = LocaliseString(StrPathFile, strFormName, "optGrp4", optGrp4.Caption)
    lblWait.Caption = LocaliseString(StrPathFile, strFormName, "lblWait", lblWait.Caption)
    cmdReNewHW.Caption = LocaliseString(StrPathFile, strFormName, "cmdReNewHW", cmdReNewHW.Caption)
    cmdBackUpDrivers.Caption = LocaliseString(StrPathFile, strFormName, "cmdBackUpDrivers", cmdBackUpDrivers.Caption)

End Sub

Private Sub lvDevices_ItemClick(ByVal Item As LvwListItem, ByVal Button As Integer)
    txtFindText.Text = ParseHwid(Item.Text)
End Sub

Private Sub lvDevices_ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)

Dim strOrigHwid                         As String

    If Button = vbLeftButton Then
        txtFindText.Text = ParseHwid(Item.Text)
        strOrigHwid = Item.SubItems(6)
        OpenDeviceProp strOrigHwid
    End If

End Sub

Private Sub lvDevices_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        OpenContextMenu Me, Me.mnuContext

    End If

End Sub

Private Sub mnuContextDelete_Click()

Dim strOrigHwid                         As String
Dim mbDeleteDriverByHwidTemp            As Boolean

    strOrigHwid = lvDevices.SelectedItem
    mbDeleteDriverByHwidTemp = DeleteDriverbyHwid(strOrigHwid)

    If mbDeleteDriverByHwidTemp Then
        If Not mbDeleteDriverByHwid Then
            mbDeleteDriverByHwid = True
        End If

    End If

End Sub

Private Sub mnuContextProperties_Click()

Dim strOrigHwid                         As String
    strOrigHwid = lvDevices.ListItems(lvDevices.SelectedItem.Index).SubItems(6)
    OpenDeviceProp strOrigHwid

End Sub

Private Sub OpenDeviceProp(ByVal strHwid As String)

Dim cmdString                           As String
Dim cmdStringParams                     As String
Dim nRetShellEx                         As Boolean

    cmdString = "rundll32.exe"
    cmdStringParams = "devmgr.dll,DeviceProperties_RunDLL /DeviceID " & strHwid
    DebugMode "cmdString: " & cmdString
    DebugMode "cmdStringParams: " & cmdStringParams
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL, cmdStringParams)
    DebugMode "cmdString: " & nRetShellEx

End Sub

Private Sub optGrp1_Click()
    LoadListbyMode
End Sub

Private Sub optGrp2_Click()
    LoadListbyMode
End Sub

Private Sub optGrp3_Click()
    LoadListbyMode
End Sub

Public Sub optGrp4_Click()
    LoadListbyMode
End Sub

Private Function ParseHwid(strValuer As String) As String

Dim strValuer_x()                       As String
Dim miSubSys                            As Long
Dim miREV                               As Long
Dim miMI                                As Long
Dim miCC                                As Long

    ' Удаление дубликатов
    If chkParseHwid.Value Then

        ' разбиваем по "\"
        If InStr(strValuer, "\") Then
            strValuer_x = Split(strValuer, "\")
            strValuer = strValuer_x(0) & "\" & strValuer_x(1)
        
            miSubSys = InStr(strValuer, "&SUBSYS")
    
            If miSubSys > 0 Then
                strValuer = Left$(strValuer, miSubSys - 1)
            End If
    
            miREV = InStr(strValuer, "&REV_")
    
            If miREV > 0 Then
                strValuer = Left$(strValuer, miREV - 1)
            End If
    
            miMI = InStr(strValuer, "&MI_")
    
            If miMI > 0 Then
                strValuer = Left$(strValuer, miMI - 1)
            End If
    
            miCC = InStr(strValuer, "&CC_")
    
            If miCC > 0 Then
                strValuer = Left$(strValuer, miCC - 1)
            End If
        End If

    End If

    ParseHwid = strValuer

End Function

Private Sub txtFindText_Change()
    cmdGoSite.Enabled = LenB(txtFindText.Text) > 0
End Sub

Private Sub lvDevices_ColumnClick(ByVal ColumnHeader As LvwColumnHeader)
Dim i                                   As Long
    With lvDevices
        .Sorted = False
        .SortKey = ColumnHeader.Index - 1
        If ComCtlsSupportLevel() >= 1 Then
            For i = 1 To .ColumnHeaders.Count
                If i <> ColumnHeader.Index Then
                    .ColumnHeaders(i).SortArrow = LvwColumnHeaderSortArrowNone
                Else
                    If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowNone Then
                        ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown
                    Else
                        If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown Then
                            ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp
                        ElseIf ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp Then
                            ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown
                        End If
                    End If
                End If
            Next i
            Select Case ColumnHeader.SortArrow
                Case LvwColumnHeaderSortArrowDown, LvwColumnHeaderSortArrowNone
                    .SortOrder = LvwSortOrderAscending
                Case LvwColumnHeaderSortArrowUp
                    .SortOrder = LvwSortOrderDescending
            End Select
            .SelectedColumn = ColumnHeader
        Else
            For i = 1 To .ColumnHeaders.Count
                If i <> ColumnHeader.Index Then
                    .ColumnHeaders(i).Icon = 0
                Else
                    If ColumnHeader.Icon = 0 Then
                        ColumnHeader.Icon = 1
                    Else
                        If ColumnHeader.Icon = 2 Then
                            ColumnHeader.Icon = 1
                        ElseIf ColumnHeader.Icon = 1 Then
                            ColumnHeader.Icon = 2
                        End If
                    End If
                End If
            Next i
            Select Case ColumnHeader.Icon
                Case 1, 0
                    .SortOrder = LvwSortOrderAscending
                Case 2
                    .SortOrder = LvwSortOrderDescending
            End Select
        End If
        .Sorted = True
        If Not .SelectedItem Is Nothing Then .SelectedItem.EnsureVisible
    End With
End Sub

