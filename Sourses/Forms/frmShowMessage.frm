VERSION 5.00
Begin VB.Form frmShowMessage 
   Caption         =   "Сообщение программы"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShowMessage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   8910
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.TextBoxW txtMessageText 
      Height          =   4035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   7223
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
      Text            =   "frmShowMessage.frx":000C
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3
      CueBanner       =   "frmShowMessage.frx":003E
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   750
      Left            =   7020
      TabIndex        =   1
      Top             =   4080
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
      Caption         =   "ОК"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Height          =   750
      Left            =   5100
      TabIndex        =   2
      Top             =   4080
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
      Caption         =   "Отмена"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
End
Attribute VB_Name = "frmShowMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngFormWidthMin  As Long
Private lngFormHeightMin As Long
Private strFormName      As String

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
'! Procedure   (Функция)   :   Sub cmdExit_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    lngShowMessageResult = vbNo
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdOK_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
    lngShowMessageResult = vbYes
    Unload Me
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
        SetIcon .hWnd, "frmShowMessage", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
        lngFormWidthMin = .Width
        lngFormHeightMin = .Height
    End With

    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal StrPathFile As String)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
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
            txtMessageText.Height = cmdExit.Top - 100
            txtMessageText.Top = 25
            txtMessageText.Left = 25
            txtMessageText.Width = .Width - miDeltaFrm - 200
            cmdOK.Left = cmdExit.Left - cmdOK.Width - 110
            cmdOK.Top = cmdExit.Top
        End If

    End With

End Sub
