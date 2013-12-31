VERSION 5.00
Begin VB.Form frmDonate 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Поблагодарить автора"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDonate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjDIADBS.ctlXpButton cmdExit 
      Height          =   850
      Left            =   7320
      TabIndex        =   0
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
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
   Begin prjDIADBS.RichTextBox DonateRTF 
      Height          =   5250
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   9260
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
      FileName        =   "frmDonate.frx":000C
      Text            =   "frmDonate.frx":002C
      TextRTF         =   "frmDonate.frx":004C
   End
   Begin prjDIADBS.ctlXpButton cmdSMSCoin 
      Height          =   850
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Donate via SMSCoin"
      PicturePosition =   2
      ButtonStyle     =   3
      Picture         =   "frmDonate.frx":01A8
      PictureWidth    =   48
      PictureHeight   =   48
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      MaskColor       =   16645372
   End
   Begin prjDIADBS.ctlXpButton cmdPayPal 
      Height          =   850
      Left            =   2280
      TabIndex        =   3
      Top             =   5400
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Donate via"
      PicturePosition =   2
      ButtonStyle     =   3
      Picture         =   "frmDonate.frx":0B11
      PictureWidth    =   73
      PictureHeight   =   38
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      MaskColor       =   16777215
   End
   Begin prjDIADBS.ctlXpButton cmdYandexMoney 
      Height          =   850
      Left            =   4560
      TabIndex        =   4
      Top             =   5400
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Donate via"
      PicturePosition =   2
      ButtonStyle     =   3
      Picture         =   "frmDonate.frx":2C0B
      PictureWidth    =   61
      PictureHeight   =   32
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      MaskColor       =   16185078
   End
End
Attribute VB_Name = "frmDonate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Минимальные размеры формы
Private lngFormWidthMin  As Long
Private lngFormHeightMin As Long
Private strFormName      As String

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CheckEditDonate
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub CheckEditDonate(StrPathFile As String)

    Dim strMD5TextRtf         As String
    Dim strDONATE_MD5RTF_temp As String

    strMD5TextRtf = GetMD5(StrPathFile)
    DebugMode "DonateInfo: " & strMD5TextRtf

    Select Case strPCLangCurrentID

        Case "0419"
            strDONATE_MD5RTF_temp = strDONATE_MD5RTF

        Case Else
            strDONATE_MD5RTF_temp = strDONATE_MD5RTF_Eng
    End Select

    If InStr(1, strMD5TextRtf, strDONATE_MD5RTF_temp, vbTextCompare) = 0 Then
        DebugMode "DonateInfo: NotValid"
        MsgBox strMessages(40), vbInformation, strProductName
        Unload Me
    End If

    DonateRTF.Visible = True
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
'! Procedure   (Функция)   :   Sub cmdPayPal_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdPayPal_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=10349042"
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdSMSCoin_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdSMSCoin_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    Select Case strPCLangCurrentID

        Case "0419"
            cmdString = "http://donate.smscoin.com/js/smsdonate/index.html?sid=403169"

        Case Else
            cmdString = "http://donate.smscoin.com/js/smsdonate/index_en.html?sid=403169"
    End Select

    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdYandexMoney_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdYandexMoney_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = "https://money.yandex.ru/embed/shop.xml?uid=41001626648736&amp;writer=seller&amp;targets=donate+to+adia-project&amp;default-sum=50&amp;button-text=04&amp;comment=on&amp;hint=%22Please,%20write%20your%20comments%22"
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
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

    SetBtnFontProperties cmdExit
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Activate
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()
    LoadDonate
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
        SetIcon .hWnd, "frmDonate", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
        lngFormWidthMin = .Width
        lngFormHeightMin = .Height
    End With

    LoadIconImage2Btn cmdExit, "BTN_EXIT", strPathImageMainWork
    DonateRTF.Visible = False

    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_QueryUnload
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Cancel (Integer)
'                              UnloadMode (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmDonate = Nothing
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Resize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Resize()

    On Error Resume Next

    With Me

        Dim miDeltaFrm As Long

        If OsCurrVersionStruct.VerFull >= "6.0" Then
            miDeltaFrm = 125
        Else

            If mbAppThemed Then
                miDeltaFrm = 0
            Else
                miDeltaFrm = 0
            End If
        End If

        If .WindowState <> vbMinimized Then
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

            cmdExit.Left = .Width - cmdExit.Width - 200 - miDeltaFrm
            cmdExit.Top = .Height - cmdExit.Height - 550 - miDeltaFrm
            DonateRTF.Width = .Width - DonateRTF.Left - 200 - miDeltaFrm
            DonateRTF.Height = cmdExit.Top - DonateRTF.Top - 100
            cmdPayPal.Top = cmdExit.Top
            cmdSMSCoin.Top = cmdExit.Top
            cmdYandexMoney.Top = cmdExit.Top
            cmdPayPal.Left = cmdSMSCoin.Left + cmdSMSCoin.Width + 110
            cmdYandexMoney.Left = cmdPayPal.Left + cmdPayPal.Width + 110
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadDonate
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadDonate()

    Dim strPathDonate As String

    Select Case strPCLangCurrentID

        Case "0419"
            strPathDonate = PathCollect(strToolsDocs_Path & "\0419\donate.rtf")

        Case Else
            strPathDonate = PathCollect(strToolsDocs_Path & "\0409\donate.rtf")
    End Select

    If PathExists(strPathDonate) Then
        DonateRTF.LoadFile strPathDonate
    Else
        MsgBox strMessages(41), vbInformation, strProductName
        Unload Me
    End If

    ' Проверка файла Donate на неправомерное изменение
    CheckEditDonate strPathDonate
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal StrPathFile As String)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.Caption = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
    'Кнопки
    cmdExit.Caption = LocaliseString(StrPathFile, strFormName, "cmdExit", cmdExit.Caption)
End Sub
