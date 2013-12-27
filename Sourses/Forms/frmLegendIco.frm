VERSION 5.00
Begin VB.Form frmLegendIco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Подробное описание обозначений кнопок пакетов драйверов"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLegendIco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox imgOkAttentionOLD 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox imgOkAttentionNew 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   7
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox imgOK 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox imgOkAttention 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.PictureBox imgOkNew 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.PictureBox imgOkOld 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
   Begin VB.PictureBox imgNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   5520
      Width           =   495
   End
   Begin VB.PictureBox imgNoDB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   1
      Top             =   6480
      Width           =   495
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   705
      Left            =   9900
      TabIndex        =   8
      Top             =   7140
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
   Begin prjDIADBS.LabelW lblDescription 
      Height          =   795
      Index           =   7
      Left            =   720
      TabIndex        =   9
      Top             =   6360
      Width           =   10995
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "Описание статусов иконок"
   End
   Begin prjDIADBS.LabelW lblDescription 
      Height          =   795
      Index           =   6
      Left            =   720
      TabIndex        =   10
      Top             =   5400
      Width           =   10995
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "Описание статусов иконок"
   End
   Begin prjDIADBS.LabelW lblDescription 
      Height          =   795
      Index           =   5
      Left            =   720
      TabIndex        =   11
      Top             =   4440
      Width           =   10995
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "Описание статусов иконок"
   End
   Begin prjDIADBS.LabelW lblDescription 
      Height          =   795
      Index           =   4
      Left            =   720
      TabIndex        =   12
      Top             =   3480
      Width           =   10995
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "Описание статусов иконок"
   End
   Begin prjDIADBS.LabelW lblDescription 
      Height          =   795
      Index           =   3
      Left            =   720
      TabIndex        =   13
      Top             =   2520
      Width           =   10995
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "Описание статусов иконок"
   End
   Begin prjDIADBS.LabelW lblDescription 
      Height          =   795
      Index           =   2
      Left            =   720
      TabIndex        =   14
      Top             =   1560
      Width           =   10995
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "Описание статусов иконок"
   End
   Begin prjDIADBS.LabelW lblDescription 
      Height          =   555
      Index           =   1
      Left            =   720
      TabIndex        =   15
      Top             =   840
      Width           =   10995
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "Описание статусов иконок"
   End
   Begin prjDIADBS.LabelW lblDescription 
      Height          =   555
      Index           =   0
      Left            =   720
      TabIndex        =   16
      Top             =   120
      Width           =   10995
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "Описание статусов иконок"
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   120
      X2              =   11760
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   120
      X2              =   11760
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   120
      X2              =   11760
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   120
      X2              =   11760
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   11760
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   11760
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmLegendIco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFormName As String

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdOK_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
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
        Unload Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FontCharsetChange
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub FontCharsetChange()

    ' Выставляем шрифт
    With Me.Font
        .Name = strOtherForm_FontName
        .Size = lngOtherForm_FontSize
        .Charset = lngDialog_Charset
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Load
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()

    Dim i As Long

    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmLegendIco", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    '--------------------- Статусные Иконки
    LoadIconImage2Object imgNoDB, "BTN_NO_DB", strPathImageStatusButtonWork
    LoadIconImage2Object imgOK, "BTN_OK", strPathImageStatusButtonWork
    LoadIconImage2Object imgOkAttention, "BTN_OK_ATTENTION", strPathImageStatusButtonWork
    LoadIconImage2Object imgOkAttentionNew, "BTN_OK_ATTENTION_NEW", strPathImageStatusButtonWork
    LoadIconImage2Object imgOkAttentionOLD, "BTN_OK_ATTENTION_OLD", strPathImageStatusButtonWork
    LoadIconImage2Object imgOkNew, "BTN_OK_NEW", strPathImageStatusButtonWork
    LoadIconImage2Object imgOkOld, "BTN_OK_OLD", strPathImageStatusButtonWork
    LoadIconImage2Object imgNo, "BTN_NO_DRV", strPathImageStatusButtonWork
    imgOK.BorderStyle = 0
    imgOkAttention.BorderStyle = 0
    imgOkNew.BorderStyle = 0
    imgOkOld.BorderStyle = 0
    imgOkAttentionNew.BorderStyle = 0
    imgOkAttentionOLD.BorderStyle = 0
    imgNo.BorderStyle = 0
    imgNoDB.BorderStyle = 0

    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork

    For i = 0 To 7
        lblDescription(i).Caption = arrTTipStatusIcon(i)
    Next

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
    cmdOK.Caption = LocaliseString(StrPathFile, strFormName, "cmdOK", cmdOK.Caption)
End Sub
