VERSION 5.00
Begin VB.Form frmFontDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Locate Font and Color ..."
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjDIADBS.TextBoxW txtFont 
      Height          =   495
      Left            =   60
      TabIndex        =   11
      Top             =   1260
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "frmFontDialog.frx":0000
      Alignment       =   2
      CueBanner       =   "frmFontDialog.frx":0052
   End
   Begin prjDIADBS.OptionButtonW opt3 
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   2700
      Visible         =   0   'False
      Width           =   1215
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
      Value           =   0   'False
      Caption         =   "frmFontDialog.frx":0072
   End
   Begin prjDIADBS.OptionButtonW opt2 
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   2700
      Visible         =   0   'False
      Width           =   1215
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
      Value           =   0   'False
      Caption         =   "frmFontDialog.frx":009E
   End
   Begin prjDIADBS.OptionButtonW opt1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2700
      Visible         =   0   'False
      Width           =   1215
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
      Caption         =   "frmFontDialog.frx":00D2
   End
   Begin prjDIADBS.SpinBox txtFontSize 
      Height          =   315
      Left            =   1860
      TabIndex        =   4
      Top             =   420
      Width           =   675
      _ExtentX        =   1191
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
      Min             =   6
      Max             =   20
      Value           =   6
      AllowOnlyNumbers=   -1  'True
   End
   Begin prjDIADBS.ctlColorButton ctlFontColor 
      Height          =   330
      Left            =   1980
      TabIndex        =   3
      Top             =   780
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   582
      Icon            =   "frmFontDialog.frx":00FC
   End
   Begin prjDIADBS.CheckBoxW chkItalic 
      Height          =   315
      Left            =   2700
      TabIndex        =   2
      Top             =   675
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmFontDialog.frx":0256
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.CheckBoxW chkBold 
      Height          =   255
      Left            =   2700
      TabIndex        =   1
      Top             =   420
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmFontDialog.frx":0282
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.ctlFontCombo ctlFontCombo 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      PreviewText     =   "ctlFontCombo1"
      ComboFontSize   =   10
      ButtonOverColor =   0
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
   Begin prjDIADBS.CheckBoxW chkUnderline 
      Height          =   315
      Left            =   2700
      TabIndex        =   5
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmFontDialog.frx":02AA
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   750
      Left            =   2280
      TabIndex        =   9
      Top             =   1860
      Width           =   2100
      _ExtentX        =   3704
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
      Left            =   60
      TabIndex        =   10
      Top             =   1860
      Width           =   2100
      _ExtentX        =   3704
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
      Caption         =   "Выход без сохранения"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.LabelW lblFontSize 
      Height          =   375
      Left            =   60
      TabIndex        =   12
      Top             =   420
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
      Caption         =   "Размер шрифта"
   End
   Begin prjDIADBS.LabelW lblFontColor 
      Height          =   375
      Left            =   60
      TabIndex        =   13
      Top             =   840
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
      Caption         =   "Цвет шрифта"
   End
End
Attribute VB_Name = "frmFontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFormName                     As String

Private Sub chkBold_Click()
    ctlFontCombo.ComboFontBold = chkBold.Value = 1
    txtFont.Font.Bold = chkBold.Value

End Sub

Private Sub chkItalic_Click()
    ctlFontCombo.ComboFontItalic = chkItalic.Value = 1
    txtFont.Font.Italic = chkItalic.Value

End Sub

Private Sub ctlFontColor_Click()
    txtFont.ForeColor = ctlFontColor.BackColor

End Sub

Private Sub chkUnderline_Click()
    txtFont.Font.Underline = chkUnderline.Value

End Sub

Private Sub ctlFontCombo_FontNotFound(FontName As String)
    MsgBox "Cant find this font: " & FontName

End Sub

Private Sub ctlFontCombo_SelectedFontChanged(NewFontName As String)
    txtFont.Font.Name = NewFontName
    ctlFontCombo.ClearUsedList
    ctlFontCombo.AddToUsedList NewFontName

End Sub

Private Sub Form_Activate()
    ctlFontCombo.SelectedFont = txtFont.Font.Name
    txtFontSize.Value = txtFont.Font.Size
    ctlFontCombo.PreviewText = txtFont.Text
    ctlFontCombo.AddToUsedList txtFont.Font.Name
    chkBold.Value = txtFont.Font.Bold
    chkItalic.Value = txtFont.Font.Italic
    chkUnderline.Value = txtFont.Font.Underline

End Sub

Private Sub Form_Load()

    SetupVisualStyles Me


    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmFontDialog", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With


    ' Устанавливаем картинки кнопок и убираем описание кнопок
    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork

    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange

    End If

    txtFontSize.Min = 6

End Sub

Private Sub Localise(ByVal StrPathFile As String)

' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.Caption = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
    ' Лэйблы
    lblFontSize.Caption = LocaliseString(StrPathFile, strFormName, "lblFontSize", lblFontSize.Caption)
    lblFontColor.Caption = LocaliseString(StrPathFile, strFormName, "lblFontColor", lblFontColor.Caption)
    chkBold.Caption = LocaliseString(StrPathFile, strFormName, "chkBold", chkBold.Caption)
    chkItalic.Caption = LocaliseString(StrPathFile, strFormName, "chkItalic", chkItalic.Caption)
    chkUnderline.Caption = LocaliseString(StrPathFile, strFormName, "chkUnderline", chkUnderline.Caption)
    txtFont.Text = LocaliseString(StrPathFile, strFormName, "txtFont", txtFont.Text)
    'Кнопки
    cmdOK.Caption = LocaliseString(StrPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(StrPathFile, strFormName, "cmdExit", cmdExit.Caption)

End Sub

Private Sub FontCharsetChange()
' Выставляем шрифт
    With Me.Font
        .Name = strOtherForm_FontName
        .Size = lngOtherForm_FontSize
        .Charset = lngDialog_Charset
    End With

    SetButtonProperties cmdExit
    SetButtonProperties cmdOK

End Sub

Private Sub txtFont_Change()
    ctlFontCombo.PreviewText = txtFont.Text
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
    SaveOptions
    Unload Me
End Sub

'! -----------------------------------------------------------
'!  Функция     :  SaveOptions
'!  Переменные  :
'!  Описание    :
'! -----------------------------------------------------------
Private Sub SaveOptions()

    With txtFont

        If opt1.Value Then
            strDialogTab_FontName = .Font.Name
            miDialogTab_FontSize = .Font.Size
            mbDialogTab_Underline = .Font.Underline
            mbDialogTab_Strikethru = .Font.Strikethrough
            mbDialogTab_Bold = .Font.Bold
            mbDialogTab_Italic = .Font.Italic
            lngDialogTab_Color = .ForeColor

        ElseIf opt2.Value Then
            strDialogTab2_FontName = .Font.Name
            miDialogTab2_FontSize = .Font.Size
            mbDialogTab2_Underline = .Font.Underline
            mbDialogTab2_Strikethru = .Font.Strikethrough
            mbDialogTab2_Bold = .Font.Bold
            mbDialogTab2_Italic = .Font.Italic
            lngDialogTab2_Color = .ForeColor
        Else
            strDialog_FontName = .Font.Name
            miDialog_FontSize = .Font.Size
            mbDialog_Underline = .Font.Underline
            mbDialog_Strikethru = .Font.Strikethrough
            mbDialog_Bold = .Font.Bold
            mbDialog_Italic = .Font.Italic
            lngDialog_Color = .ForeColor
            SetButtonProperties frmOptions.cmdFutureButton
            frmOptions.cmdFutureButton.Refresh
        End If

    End With

End Sub

Private Sub txtFontSize_Change()
    txtFont.Font.Size = txtFontSize.Value
End Sub

Private Sub txtFontSize_TextChange()
    txtFont.Font.Size = txtFontSize.Value
End Sub
