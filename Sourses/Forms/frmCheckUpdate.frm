VERSION 5.00
Begin VB.Form frmCheckUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Обновление: Обнаружена новая версия программы"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCheckUpdate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11340
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.ComboBoxW cmbVersions 
      Height          =   315
      Left            =   5100
      TabIndex        =   2
      Top             =   450
      Width           =   1335
      _ExtentX        =   2355
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
      Style           =   2
      Sorted          =   -1  'True
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Default         =   -1  'True
      Height          =   650
      Left            =   9345
      TabIndex        =   10
      Top             =   5160
      Width           =   1935
      _ExtentX        =   3413
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
   Begin prjDIADBS.ctlJCbutton cmdHistory 
      Height          =   650
      Left            =   4700
      TabIndex        =   8
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "История изменений"
      CaptionEffects  =   0
      PictureAlign    =   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdUpdateFull 
      Height          =   650
      Left            =   2415
      TabIndex        =   7
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Скачать дистрибутив"
      CaptionEffects  =   0
      PictureAlign    =   0
      DropDownSymbol  =   6
      DropDownSeparator=   -1  'True
      DropDownEnable  =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdUpdate 
      Height          =   650
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Скачать обновление"
      CaptionEffects  =   0
      PictureAlign    =   0
      DropDownSymbol  =   6
      DropDownSeparator=   -1  'True
      DropDownEnable  =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdDonate 
      Height          =   650
      Left            =   6990
      TabIndex        =   9
      Top             =   5160
      Width           =   2220
      _ExtentX        =   3916
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
   Begin prjDIADBS.RichTextBox rtfDescription 
      Height          =   4275
      Left            =   120
      TabIndex        =   4
      Top             =   800
      Visible         =   0   'False
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   7541
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragDrop     =   0   'False
      Locked          =   -1  'True
      HideSelection   =   0   'False
      MultiLine       =   -1  'True
      ScrollBars      =   2
      WantReturn      =   -1  'True
      TextRTF         =   "frmCheckUpdate.frx":000C
   End
   Begin prjDIADBS.LabelW lblWait 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   11160
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "Идет загрузка данных с официального сайта. Пожалуйста, подождите...."
   End
   Begin prjDIADBS.LabelW lblVersionList 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   450
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Список изменений для версии:"
   End
   Begin prjDIADBS.LabelW lblWWW 
      Height          =   315
      Left            =   8100
      TabIndex        =   3
      Top             =   450
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12582912
      MousePointer    =   4
      Alignment       =   1
      BackStyle       =   0
      Caption         =   "www.adia-project.net"
   End
   Begin prjDIADBS.LabelW lblVersion 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   11085
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Последняя версия программы: "
   End
   Begin VB.Menu mnuContextMenu1 
      Caption         =   "Контекстное меню 1"
      Begin VB.Menu mnuContextLinkFull 
         Caption         =   "Посетить сайт 1"
         Index           =   0
      End
      Begin VB.Menu mnuContextLinkFull 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContextLinkFull 
         Caption         =   "Посетить сайт 2"
         Index           =   2
      End
      Begin VB.Menu mnuContextLinkFull 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextLinkFull 
         Caption         =   "Посетить сайт 3"
         Index           =   4
      End
   End
   Begin VB.Menu mnuContextMenu2 
      Caption         =   "Контекстное меню 2"
      Begin VB.Menu mnuContextLinkUpdate 
         Caption         =   "Посетить сайт 1"
         Index           =   0
      End
      Begin VB.Menu mnuContextLinkUpdate 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContextLinkUpdate 
         Caption         =   "Посетить сайт 2"
         Index           =   2
      End
      Begin VB.Menu mnuContextLinkUpdate 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextLinkUpdate 
         Caption         =   "Посетить сайт 3"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmCheckUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbFirstStartUpdate As Boolean
Private strFormName        As String

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
'! Procedure   (Функция)   :   Sub LoadButtonLinkFull
'! Description (Описание)  :   [Контекстное меню для кнопки скачать дистрибутив]
'! Parameters  (Переменные):   strMassivLink() (String)
'!--------------------------------------------------------------------------------
Private Sub LoadButtonLinkFull(strMassivLink() As String)

    Dim strMirrorText As String

    If cmbVersions.ListIndex > -1 Then

        ' Отличия работы если русский или английский
        Select Case strPCLangCurrentID

            Case "0419"
                strMirrorText = "Зеркало"

            Case Else
                strMirrorText = "Mirror"
        End Select

        If InStr(1, strMassivLink(cmbVersions.ListIndex, 0), "http", vbTextCompare) Then
            cmdUpdateFull.DropDownEnable = True
        ElseIf InStr(1, strMassivLink(cmbVersions.ListIndex, 2), "http", vbTextCompare) Then
            cmdUpdateFull.DropDownEnable = True
        Else
            cmdUpdateFull.DropDownEnable = False
        End If

        If cmdUpdateFull.DropDownEnable Then
                            
            If InStr(1, strMassivLink(cmbVersions.ListIndex, 2), "http", vbTextCompare) = 0 Then
                mnuContextLinkFull(2).Enabled = False
            End If

            If InStr(1, strMassivLink(cmbVersions.ListIndex, 4), "http", vbTextCompare) = 0 Then
                mnuContextLinkFull(4).Enabled = False
            End If

            If LenB(strMassivLink(cmbVersions.ListIndex, 1)) Then
                mnuContextLinkFull(0).Visible = True
                mnuContextLinkFull(1).Visible = True
                SetUniMenu 0, 0, -1, mnuContextMenu1, strMassivLink(cmbVersions.ListIndex, 1)
            Else
                mnuContextLinkFull(0).Visible = False
                mnuContextLinkFull(1).Visible = False
            End If

            If LenB(strMassivLink(cmbVersions.ListIndex, 3)) Then
                mnuContextLinkFull(1).Visible = True
                mnuContextLinkFull(2).Visible = True
                SetUniMenu 0, 2, -1, mnuContextMenu1, strMassivLink(cmbVersions.ListIndex, 3)
            Else
                mnuContextLinkFull(1).Visible = False
                mnuContextLinkFull(2).Visible = False
            End If

            If LenB(strMassivLink(cmbVersions.ListIndex, 5)) Then
                mnuContextLinkFull(3).Visible = True
                mnuContextLinkFull(4).Visible = True
                SetUniMenu 0, 4, -1, mnuContextMenu1, strMassivLink(cmbVersions.ListIndex, 5)
            Else
                mnuContextLinkFull(3).Visible = False
                mnuContextLinkFull(4).Visible = False
            End If
        End If

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadButtonLinkUpdate
'! Description (Описание)  :   [Контекстное меню для кнопки скачать обновление]
'! Parameters  (Переменные):   strMassivLink() (String)
'!--------------------------------------------------------------------------------
Private Sub LoadButtonLinkUpdate(strMassivLink() As String)

    Dim strMirrorText As String

    If cmbVersions.ListIndex > -1 Then

        ' Отличия работы если русский или английский
        Select Case strPCLangCurrentID

            Case "0419"
                strMirrorText = "Зеркало"

            Case Else
                strMirrorText = "Mirror"
        End Select

        If InStr(1, strMassivLink(cmbVersions.ListIndex, 0), "http", vbTextCompare) Then
            cmdUpdate.DropDownEnable = True
        ElseIf InStr(1, strMassivLink(cmbVersions.ListIndex, 2), "http", vbTextCompare) Then
            cmdUpdate.DropDownEnable = True
        Else
            cmdUpdate.DropDownEnable = False
        End If

        If cmdUpdate.DropDownEnable Then
                            
            If InStr(1, strMassivLink(cmbVersions.ListIndex, 2), "http", vbTextCompare) = 0 Then
                mnuContextLinkUpdate(2).Enabled = False
            End If

            If InStr(1, strMassivLink(cmbVersions.ListIndex, 4), "http", vbTextCompare) = 0 Then
                mnuContextLinkUpdate(4).Enabled = False
            End If

            If LenB(strMassivLink(cmbVersions.ListIndex, 1)) Then
                mnuContextLinkUpdate(0).Visible = True
                mnuContextLinkUpdate(1).Visible = True
                SetUniMenu 1, 0, -1, mnuContextMenu2, strMassivLink(cmbVersions.ListIndex, 1)
            Else
                mnuContextLinkUpdate(0).Visible = False
                mnuContextLinkUpdate(1).Visible = False
            End If

            If LenB(strMassivLink(cmbVersions.ListIndex, 3)) Then
                mnuContextLinkUpdate(1).Visible = True
                mnuContextLinkUpdate(2).Visible = True
                SetUniMenu 1, 2, -1, mnuContextMenu2, strMassivLink(cmbVersions.ListIndex, 3)
            Else
                mnuContextLinkUpdate(1).Visible = False
                mnuContextLinkUpdate(2).Visible = False
            End If

            If LenB(strMassivLink(cmbVersions.ListIndex, 5)) Then
                mnuContextLinkUpdate(3).Visible = True
                mnuContextLinkUpdate(4).Visible = True
                SetUniMenu 1, 4, -1, mnuContextMenu2, strMassivLink(cmbVersions.ListIndex, 5)
            Else
                mnuContextLinkUpdate(3).Visible = False
                mnuContextLinkUpdate(4).Visible = False
            End If
        End If

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadDescriptionAndLinks
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadDescriptionAndLinks()

    Dim strDescriptionTemp As String

    ' Отличия работы если русский или английский
    Select Case strPCLangCurrentID

        Case "0419"
            strDescriptionTemp = Replace$(strDescription, vbLf, vbNewLine)

        Case Else
            strDescriptionTemp = Replace$(strDescription_en, vbLf, vbNewLine)
    End Select

    ' Кнопка Скачать обновление
    LoadButtonLinkUpdate strLink
    ' Кнопка Скачать дистрибутив
    LoadButtonLinkFull strLinkFull

    ' Описание изменений
    If LenB(strDescriptionTemp) Then
        rtfDescription.TextRTF = strDescriptionTemp
    Else
        rtfDescription.TextRTF = "Error on load ChangeLog. Please inform the developer"
    End If

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
    ' Кнопки
    cmdUpdate.Caption = LocaliseString(strPathFile, strFormName, "cmdUpdate", cmdUpdate.Caption)
    cmdUpdateFull.Caption = LocaliseString(strPathFile, strFormName, "cmdUpdateFull", cmdUpdateFull.Caption)
    cmdHistory.Caption = LocaliseString(strPathFile, strFormName, "cmdHistory", cmdHistory.Caption)
    cmdDonate.Caption = LocaliseString(strPathFile, strFormName, "cmdDonate", cmdDonate.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
    ' Лейблы
    lblVersion.Caption = LocaliseString(strPathFile, strFormName, "lblVersion", lblVersion.Caption) & strSpace & strVersion & " (" & strDateProg & ")"
    ' Меню
    LocaliseMenu strPathFile

    If InStr(1, strRelease, "beta", vbTextCompare) Then
        lblVersion.Caption = lblVersion.Caption & " This version may be Unstable!!!"
        lblVersion.ForeColor = vbRed
    End If

    lblVersionList.Caption = LocaliseString(strPathFile, strFormName, "lblVersionList", lblVersionList.Caption)
    lblWait.Caption = LocaliseString(strPathFile, strFormName, "lblWait", lblWait.Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LocaliseMenu
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub LocaliseMenu(ByVal strPathFile As String)
    SetUniMenu -1, 0, -1, mnuContextMenu1, LocaliseString(strPathFile, strFormName, "cmdUpdateFull", cmdUpdateFull.Caption)
    SetUniMenu -1, 1, -1, mnuContextMenu2, LocaliseString(strPathFile, strFormName, "cmdUpdate", cmdUpdate.Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbVersions_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbVersions_Click()

    With cmbVersions

        If .ListIndex > -1 Then
            strDescription = strUpdDescription(.ListIndex, 0)
            strDescription_en = strUpdDescription(.ListIndex, 1)
        Else
            strDescription = vbNullString
            strDescription_en = vbNullString
        End If

    End With

    LoadDescriptionAndLinks
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDonate_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdDonate_Click()
    frmDonate.Show vbModal, Me
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
'! Procedure   (Функция)   :   Sub cmdHistory_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdHistory_Click()

    Select Case strPCLangCurrentID

        Case "0419"
            RunUtilsShell strLinkHistory, False

        Case Else
            RunUtilsShell strLinkHistory_en, False
    End Select
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdUpdate_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdUpdate_Click()
    RunUtilsShell strLink(cmbVersions.ListIndex, 0), False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdUpdateFull_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdUpdateFull_Click()
    RunUtilsShell strLinkFull(cmbVersions.ListIndex, 0), False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Activate
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()

    Dim ii As Long

    If mbFirstStartUpdate Then
        lblWait.Visible = True
        DoEvents
        ' Загрузка данных с сайта
        LoadUpdateData
        DoEvents
        ' установка параметров для кнопок
        LoadDescriptionAndLinks
        ' Показываем список изменений
        lblWait.Visible = False
        rtfDescription.Visible = True
        cmbVersions.Left = lblVersionList.Left + lblVersionList.Width + 50

        For ii = LBound(strUpdVersions) To UBound(strUpdVersions)
            cmbVersions.AddItem strUpdVersions(ii), ii
        Next

        cmbVersions.ListIndex = 0
        
        rtfDescription.SetFocus
    End If

    mbFirstStartUpdate = False
    cmdUpdate.Enabled = True
    cmdUpdateFull.Enabled = True
    cmdHistory.Enabled = True
    cmdDonate.Enabled = True
    cmdExit.Enabled = True
    cmbVersions.Enabled = True
    mnuContextMenu1.Enabled = True
    mnuContextMenu2.Enabled = True
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

    cmdUpdate.Enabled = False
    cmdUpdateFull.Enabled = False
    cmdHistory.Enabled = False
    cmdDonate.Enabled = False
    cmdExit.Enabled = False
    cmbVersions.Enabled = False
    mnuContextMenu1.Enabled = False
    mnuContextMenu2.Enabled = False
    
    mbFirstStartUpdate = True
    lblWait.Visible = True
    DoEvents
    lblWait.Left = 100
    lblWait.Width = Me.Width - 200
    LoadIconImage2Object cmdExit, "BTN_EXIT", strPathImageMainWork
    LoadIconImage2Object cmdUpdate, "BTN_UPDATE", strPathImageMainWork
    LoadIconImage2Object cmdUpdateFull, "BTN_UPDATEFULL", strPathImageMainWork
    LoadIconImage2Object cmdHistory, "BTN_HISTORY", strPathImageMainWork
    LoadIconImage2Object cmdDonate, "BTN_DONATE", strPathImageMainWork

    ' Локализация приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    ' Контекстное меню для кнопки скачать дистрибутив
    cmdUpdateFull.SetPopupMenu mnuContextMenu1

    ' Контекстное меню для кнопки скачать обновление
    cmdUpdate.SetPopupMenu mnuContextMenu2

End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdUpdate.UnsetPopupMenu
    cmdUpdateFull.UnsetPopupMenu
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lblWWW_MouseDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub lblWWW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RunUtilsShell strUrl_MainWWWSite, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextLinkFull_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuContextLinkFull_Click(Index As Integer)
    
    Select Case Index

        Case 0
            RunUtilsShell strLinkFull(cmbVersions.ListIndex, 0), False

        Case 2
            RunUtilsShell strLinkFull(cmbVersions.ListIndex, 2), False
            
        Case 4
            RunUtilsShell strLinkFull(cmbVersions.ListIndex, 4), False
    End Select
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextLinkUpdate_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuContextLinkUpdate_Click(Index As Integer)
    
    Select Case Index

        Case 0
            RunUtilsShell strLink(cmbVersions.ListIndex, 0), False

        Case 2
            RunUtilsShell strLink(cmbVersions.ListIndex, 2), False

        Case 4
            RunUtilsShell strLink(cmbVersions.ListIndex, 4), False
    End Select
    
End Sub
