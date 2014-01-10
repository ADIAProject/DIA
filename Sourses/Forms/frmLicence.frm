VERSION 5.00
Begin VB.Form frmLicence 
   Caption         =   "Лицензионное соглашение"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
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
   Icon            =   "frmLicence.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin prjDIADBS.CheckBoxW chkAgreeLicence 
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   5580
      Width           =   5295
      _ExtentX        =   9340
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
      Caption         =   "frmLicence.frx":000C
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.RichTextBox LicenceRTF 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   9551
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
      FileName        =   "frmLicence.frx":0082
      Text            =   "frmLicence.frx":00A2
      TextRTF         =   "frmLicence.frx":00C2
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   700
      Left            =   7560
      TabIndex        =   2
      Top             =   5520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1244
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
      Caption         =   "ОК"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Height          =   700
      Left            =   5640
      TabIndex        =   3
      Top             =   5520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1244
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
      Caption         =   "Отмена"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
End
Attribute VB_Name = "frmLicence"
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
'! Procedure   (Функция)   :   Sub CheckEditLicense
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub CheckEditLicense(StrPathFile As String)

    Dim strMD5TextRtf       As String
    Dim strEULA_MD5RTF_temp As String

    strMD5TextRtf = GetMD5(StrPathFile)
    DebugMode "LicenceInfo: " & strMD5TextRtf

    Select Case strPCLangCurrentID

        Case "0419"
            strEULA_MD5RTF_temp = strEULA_MD5RTF

        Case Else
            strEULA_MD5RTF_temp = strEULA_MD5RTF_Eng
    End Select

    If StrComp(strMD5TextRtf, strEULA_MD5RTF_temp, vbTextCompare) <> 0 Then
        If Not mbSilentRun Then
            DebugMode "LicenceInfo: NotValid"

            If MsgBox(strMessages(11), vbYesNo + vbQuestion, strProductName) = vbNo Then
                Unload Me
            End If
        End If

        DebugMode "The Source text of the file of the license agreement was changed!!! The most Further functioning(working) the program impossible. Address to developer or download anew distribution program of the program."

    End If

    DebugMode "LicenceText: End"
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
    SetBtnFontProperties cmdOK
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkAgreeLicence_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkAgreeLicence_Click()
    cmdOK.Enabled = chkAgreeLicence.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdExit_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()

    If mbFirstStart Then
        Unload Me
        'End
    Else
        Unload Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdOK_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
    ' если принимаем соглашение, записываем параметры в реестр
    SaveSetting App.ProductName, "Licence", "Show at Startup", Not CBool(chkAgreeLicence.Value)
    SaveSetting App.ProductName, "Licence", "EULA_DATE", strEULA_Version
    ' Загружаем основную форму
    frmLicence.Hide
    Set frmMain = New frmMain
    frmMain.Show
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_KeyDown
'! Description (Описание)  :   [обработка нажатий клавиш клавиатуры]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        cmdExit_Click
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
        SetIcon .hWnd, "frmLicence", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
        lngFormWidthMin = .Width
        lngFormHeightMin = .Height
    End With

    If mbFirstStart Then
        cmdOK.Visible = True
        chkAgreeLicence.Visible = True
        cmdOK.Enabled = False
    Else
        cmdOK.Visible = False
        chkAgreeLicence.Visible = False
    End If

    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork
    
    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If
    
    LoadLicence

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_QueryUnload
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Cancel (Integer)
'                              UnloadMode (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    Set frmLicence = Nothing
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Resize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Resize()

    Dim miDeltaFrm As Long

    On Error Resume Next

    With Me

        If .WindowState <> vbMinimized Then
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

            cmdExit.Left = .Width - cmdExit.Width - 200 - miDeltaFrm
            cmdExit.Top = .Height - cmdExit.Height - 600 - miDeltaFrm
            LicenceRTF.Width = .Width - LicenceRTF.Left - 200 - miDeltaFrm
            LicenceRTF.Height = cmdExit.Top - LicenceRTF.Top - 100
            cmdOK.Left = cmdExit.Left - cmdOK.Width - 110
            cmdOK.Top = cmdExit.Top
            chkAgreeLicence.Left = 100
            chkAgreeLicence.Top = cmdExit.Top
        End If

    End With

End Sub

'Private Sub LicenceRTF_Click()
'
'Dim lngRetVal                           As Long
'Dim strBuffer                           As String
'Dim intInStr                            As Integer
'Dim intLo                               As Integer
'
'    lngRetVal = SendMessage(LicenceRTF.hWnd, EM_GETSEL, 0, 0)
'    HiWord (lngRetVal) + 1
'    intLo = LoWord(lngRetVal) + 1
'    intInStr = InStrRev(LicenceRTF.Text, " ", intLo)
'
'    If intInStr = 0 Then
'        strBuffer = Left$(LicenceRTF.Text, intLo)
'    Else
'        strBuffer = Mid$(LicenceRTF.Text, intInStr + 1)
'
'    End If
'
'    strBuffer = Trim$(strBuffer)
'    intInStr = InStr(strBuffer, " ")
'
'    If intInStr <> 0 Then
'        strBuffer = Left$(strBuffer, intInStr - 1)
'
'    End If
'
'    Select Case True
'
'        Case InStr(strBuffer, "http:")
'
'        Case InStr(strBuffer, "file:")
'
'        Case InStr(strBuffer, "mailto:")
'
'        Case InStr(strBuffer, "ftp:")
'
'        Case InStr(strBuffer, "https:")
'
'        Case InStr(strBuffer, "gopher:")
'
'        Case InStr(strBuffer, "prospero:")
'
'        Case InStr(strBuffer, "telnet:")
'
'        Case InStr(strBuffer, "news:")
'
'        Case InStr(strBuffer, "wais:")
'
'        Case Else
'            Exit Sub
'
'    End Select
'
'    'to run
'    ShellExecute Me.hWnd, "OPEN", strBuffer, vbNullString, vbNullString, 5
'
'End Sub
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadLicence
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadLicence()

    Dim strPathLicence As String

    strPathLicence = strWorkTempBackSL & "licence.rtf"

    Select Case strPCLangCurrentID

        Case "0419"
            strPathLicence = PathCollect(strToolsDocs_Path & "\0419\licence.rtf")

        Case Else
            strPathLicence = PathCollect(strToolsDocs_Path & "\0409\licence.rtf")
    End Select

    If PathExists(strPathLicence) Then
        LicenceRTF.LoadFile strPathLicence
        
        ' Проверка лицензии на неправомерное изменение
        CheckEditLicense strPathLicence
    Else

        If Not mbSilentRun Then
            MsgBox strMessages(39), vbInformation, strProductName
        End If

        Unload Me
    End If

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
    ' Чекбокс
    chkAgreeLicence.Caption = LocaliseString(StrPathFile, strFormName, "chkAgreeLicence", chkAgreeLicence.Caption)
    'Кнопки
    cmdOK.Caption = LocaliseString(StrPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(StrPathFile, strFormName, "cmdExit", cmdExit.Caption)
End Sub

'Private Sub LicenceRTF_LinkEvent(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal LinkStart As Long, ByVal LinkEnd As Long)
'Debug.Print LinkStart & " " & LinkEnd
'Debug.Print Mid$(LicenceRTF.Text, LinkStart, (LinkEnd - LinkStart))
'End Sub
