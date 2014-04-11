VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройки программы"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   13725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   13725
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.ImageList ImageListOptions 
      Left            =   240
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      InitListImages  =   "frmOptions.frx":000C
   End
   Begin prjDIADBS.ctlJCFrames frOptions 
      Height          =   5300
      Left            =   50
      Top             =   25
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Настройки"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ctlJCbutton cmdOK 
         Height          =   645
         Left            =   75
         TabIndex        =   1
         Top             =   3735
         Width           =   2850
         _ExtentX        =   5027
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
         BackColor       =   16765357
         Caption         =   "Сохранить изменения и выйти"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdExit 
         Default         =   -1  'True
         Height          =   645
         Left            =   75
         TabIndex        =   2
         Top             =   4515
         Width           =   2850
         _ExtentX        =   5027
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
         BackColor       =   16765357
         Caption         =   "Выход без сохранения"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ListView lvOptions 
         Height          =   3195
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   5636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Icons           =   "frmOptions.frx":002C
         SmallIcons      =   "frmOptions.frx":0058
         ColumnHeaderIcons=   "frmOptions.frx":0098
         BorderStyle     =   1
         View            =   2
         Arrange         =   3
         LabelEdit       =   2
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         ClickableColumnHeaders=   0   'False
         TrackSizeColumnHeaders=   0   'False
         ResizableColumnHeaders=   0   'False
      End
   End
   Begin prjDIADBS.ctlJCFrames frMain 
      Height          =   5300
      Left            =   3105
      Top             =   25
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Основные настройки программы"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.CheckBoxW chkRemoveTemp 
         Height          =   210
         Left            =   435
         TabIndex        =   16
         Top             =   3600
         Width           =   7920
         _ExtentX        =   8281
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
         Caption         =   "frmOptions.frx":00C4
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkUpdate 
         Height          =   210
         Left            =   435
         TabIndex        =   4
         Top             =   800
         Width           =   3240
         _ExtentX        =   5715
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
         Caption         =   "frmOptions.frx":013C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkReadDPName 
         Height          =   210
         Left            =   435
         TabIndex        =   8
         Top             =   1850
         Width           =   7920
         _ExtentX        =   11430
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
         Caption         =   "frmOptions.frx":0198
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkConvertDPName 
         Height          =   210
         Left            =   435
         TabIndex        =   7
         Top             =   1500
         Width           =   7920
         _ExtentX        =   13758
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
         Caption         =   "frmOptions.frx":020C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkHideOtherProcess 
         Height          =   210
         Left            =   435
         TabIndex        =   12
         Top             =   2550
         Width           =   7920
         _ExtentX        =   6350
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
         Caption         =   "frmOptions.frx":02DC
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkTempPath 
         Height          =   210
         Left            =   435
         TabIndex        =   14
         Top             =   3250
         Width           =   3255
         _ExtentX        =   5741
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
         Caption         =   "frmOptions.frx":0342
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkUpdateBeta 
         Height          =   210
         Left            =   3780
         TabIndex        =   5
         Top             =   800
         Width           =   4560
         _ExtentX        =   8043
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
         Caption         =   "frmOptions.frx":0392
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSilentDll 
         Height          =   210
         Left            =   435
         TabIndex        =   6
         Top             =   1150
         Width           =   7920
         _ExtentX        =   13970
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
         Caption         =   "frmOptions.frx":0408
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSearchOnStart 
         Height          =   210
         Left            =   435
         TabIndex        =   9
         Top             =   2200
         Width           =   5280
         _ExtentX        =   9313
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
         Caption         =   "frmOptions.frx":04A4
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.SpinBox txtPauseAfterSearch 
         Height          =   255
         Left            =   7710
         TabIndex        =   11
         Top             =   2200
         Width           =   660
         _ExtentX        =   1164
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
         AllowOnlyNumbers=   -1  'True
      End
      Begin prjDIADBS.ctlUcPickBox ucTempPath 
         Height          =   315
         Left            =   3780
         TabIndex        =   15
         Top             =   3200
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         DefaultExt      =   ""
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.ctlJCbutton optRezim_Intellect 
         Height          =   510
         Left            =   420
         TabIndex        =   18
         Top             =   4300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   900
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   14935011
         Caption         =   "Установка (Совместимые драйвера)"
         CaptionEffects  =   0
         Mode            =   2
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         ColorScheme     =   3
      End
      Begin prjDIADBS.ctlJCbutton optRezim_Upd 
         Height          =   510
         Left            =   5700
         TabIndex        =   20
         Top             =   4300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   900
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   14935011
         Caption         =   "Создание или обновление базы драйверов"
         CaptionEffects  =   0
         Mode            =   2
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         ColorScheme     =   3
      End
      Begin prjDIADBS.ctlJCbutton optRezim_Ust 
         Height          =   510
         Left            =   3060
         TabIndex        =   19
         Top             =   4300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   900
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   14935011
         Caption         =   "Установка (Полная - весь пакет)"
         CaptionEffects  =   0
         Mode            =   2
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         ColorScheme     =   3
      End
      Begin prjDIADBS.LabelW lblPauseAfterSearch 
         Height          =   225
         Left            =   5400
         TabIndex        =   10
         Top             =   2200
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Пауза после поиска: "
      End
      Begin prjDIADBS.LabelW lblOptionsTemp 
         Height          =   270
         Left            =   180
         TabIndex        =   13
         Top             =   2900
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Работа с временными файлами"
      End
      Begin prjDIADBS.LabelW lblOptionsStart 
         Height          =   270
         Left            =   180
         TabIndex        =   3
         Top             =   465
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Действия при запуске программы"
      End
      Begin prjDIADBS.LabelW lblRezim 
         Height          =   270
         Left            =   180
         TabIndex        =   17
         Top             =   3950
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Режим работы при старте программы"
      End
   End
   Begin prjDIADBS.ctlJCFrames frMain2 
      Height          =   5295
      Left            =   3300
      Top             =   300
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Основные настройки программы 2"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin VB.CommandButton cmdDriverVer 
         Caption         =   "?"
         Height          =   255
         Left            =   300
         TabIndex        =   24
         Top             =   1550
         Width           =   255
      End
      Begin prjDIADBS.OptionButtonW optCompareByVersion 
         Height          =   255
         Left            =   300
         TabIndex        =   27
         Top             =   2250
         Width           =   8100
         _ExtentX        =   14288
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
         Value           =   0   'False
         Caption         =   "frmOptions.frx":0522
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optCompareByDate 
         Height          =   255
         Left            =   300
         TabIndex        =   26
         Top             =   1900
         Width           =   8100
         _ExtentX        =   14288
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
         Caption         =   "frmOptions.frx":05A4
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtCompareVersionDRV 
         Height          =   1005
         Left            =   300
         TabIndex        =   28
         Top             =   2600
         Width           =   8100
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
         BackColor       =   -2147483633
         ForeColor       =   255
         BorderStyle     =   0
         Text            =   "frmOptions.frx":0652
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         CueBanner       =   "frmOptions.frx":0846
      End
      Begin prjDIADBS.CheckBoxW chkDateFormatRus 
         Height          =   210
         Left            =   300
         TabIndex        =   22
         Top             =   850
         Width           =   8100
         _ExtentX        =   14288
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
         Caption         =   "frmOptions.frx":0866
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkCreateRP 
         Height          =   210
         Left            =   300
         TabIndex        =   21
         Top             =   500
         Width           =   8100
         _ExtentX        =   14288
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
         Caption         =   "frmOptions.frx":08E0
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkCompatiblesHWID 
         Height          =   210
         Left            =   300
         TabIndex        =   23
         Top             =   1200
         Width           =   8100
         _ExtentX        =   14288
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
         Caption         =   "frmOptions.frx":0968
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblCompareVersionDRV 
         Height          =   225
         Left            =   600
         TabIndex        =   25
         Top             =   1550
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Сравнение версий драйверов"
      End
   End
   Begin prjDIADBS.ctlJCFrames frMainTools 
      Height          =   5295
      Left            =   3480
      Top             =   615
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Расположение основных утилит (Tools)"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ctlUcPickBox ucDevCon86Path 
         Height          =   315
         Left            =   2520
         TabIndex        =   29
         Top             =   450
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucDevCon64Path 
         Height          =   315
         Left            =   2520
         TabIndex        =   31
         Top             =   850
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucDevCon86Pathw2k 
         Height          =   315
         Left            =   2520
         TabIndex        =   33
         Top             =   1250
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucDPInst86Path 
         Height          =   315
         Left            =   2520
         TabIndex        =   35
         Top             =   1650
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucDPInst64Path 
         Height          =   315
         Left            =   2520
         TabIndex        =   37
         Top             =   2050
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucArchPath 
         Height          =   315
         Left            =   2520
         TabIndex        =   39
         Top             =   2450
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucCmdDevconPath 
         Height          =   315
         Left            =   2520
         TabIndex        =   41
         Top             =   2850
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.ctlJCbutton cmdPathDefault 
         Height          =   495
         Left            =   4900
         TabIndex        =   43
         Top             =   3300
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Сбросить настройки расположения утилит"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.LabelW lblDevCon64 
         Height          =   315
         Left            =   100
         TabIndex        =   32
         Top             =   875
         Width           =   2350
         _ExtentX        =   4154
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "DevCon.exe (64-bit)"
         WordWrap        =   0   'False
      End
      Begin prjDIADBS.LabelW lblDevCon86w2k 
         Height          =   315
         Left            =   100
         TabIndex        =   34
         Top             =   1275
         Width           =   2350
         _ExtentX        =   4154
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "DevCon.exe (for Windows 2k)"
         WordWrap        =   0   'False
      End
      Begin prjDIADBS.LabelW lblCmdDevconPath 
         Height          =   315
         Left            =   100
         TabIndex        =   42
         Top             =   2875
         Width           =   2350
         _ExtentX        =   4154
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "devcon_c.cmd"
         WordWrap        =   0   'False
      End
      Begin prjDIADBS.LabelW lblArc 
         Height          =   315
         Left            =   100
         TabIndex        =   40
         Top             =   2475
         Width           =   2350
         _ExtentX        =   4154
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "7za"
         WordWrap        =   0   'False
      End
      Begin prjDIADBS.LabelW lblDPInst64 
         Height          =   315
         Left            =   100
         TabIndex        =   38
         Top             =   2075
         Width           =   2350
         _ExtentX        =   4154
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "DPInst.exe (64-bit)"
         WordWrap        =   0   'False
      End
      Begin prjDIADBS.LabelW lblDPInst86 
         Height          =   315
         Left            =   100
         TabIndex        =   36
         Top             =   1675
         Width           =   2350
         _ExtentX        =   4154
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "DPInst.exe (32-bit)"
         WordWrap        =   0   'False
      End
      Begin prjDIADBS.LabelW lblDevCon86 
         Height          =   315
         Left            =   100
         TabIndex        =   30
         Top             =   475
         Width           =   2350
         _ExtentX        =   4154
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "DevCon.exe (32-bit)"
         WordWrap        =   0   'False
      End
   End
   Begin prjDIADBS.ctlJCFrames frOtherTools 
      Height          =   5295
      Left            =   3675
      Top             =   930
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Вспомогательные утилиты (Отображаются в меню ""Утилиты"")"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ListView lvUtils 
         Height          =   3855
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   6800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Icons           =   "frmOptions.frx":09E2
         SmallIcons      =   "frmOptions.frx":0A0E
         ColumnHeaderIcons=   "frmOptions.frx":0A3A
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
      Begin prjDIADBS.ctlJCbutton cmdAddUtil 
         Height          =   650
         Left            =   120
         TabIndex        =   45
         Top             =   4440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Добавить"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdEditUtil 
         Height          =   650
         Left            =   2160
         TabIndex        =   46
         Top             =   4455
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Изменить"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdDelUtil 
         Height          =   650
         Left            =   4200
         TabIndex        =   47
         Top             =   4455
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Удалить"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
   End
   Begin prjDIADBS.ctlJCFrames frOS 
      Height          =   5295
      Left            =   3885
      Top             =   1245
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Поддерживаемые ОС"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ListView lvOS 
         Height          =   2000
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   8355
         _ExtentX        =   14737
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
         Icons           =   "frmOptions.frx":0A66
         SmallIcons      =   "frmOptions.frx":0A92
         ColumnHeaderIcons=   "frmOptions.frx":0ABE
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
      Begin prjDIADBS.TextBoxW txtExcludeHWID 
         Height          =   495
         Left            =   120
         TabIndex        =   50
         Top             =   2820
         Width           =   8355
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
         Text            =   "frmOptions.frx":0AEA
         MultiLine       =   -1  'True
         ScrollBars      =   2
         CueBanner       =   "frmOptions.frx":0B0A
      End
      Begin prjDIADBS.CheckBoxW chkLoadFinishFile 
         Height          =   345
         Left            =   120
         TabIndex        =   52
         Top             =   4000
         Width           =   8355
         _ExtentX        =   14737
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
         Caption         =   "frmOptions.frx":0B2A
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkRecursion 
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3700
         Width           =   8355
         _ExtentX        =   14737
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
         Caption         =   "frmOptions.frx":0BF6
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdAddOS 
         Height          =   650
         Left            =   120
         TabIndex        =   53
         Top             =   4440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Добавить"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdEditOS 
         Height          =   650
         Left            =   2160
         TabIndex        =   54
         Top             =   4455
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Изменить"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdDelOS 
         Height          =   650
         Left            =   4200
         TabIndex        =   55
         Top             =   4455
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Удалить"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkDP_is_aFolder 
         Height          =   255
         Left            =   120
         TabIndex        =   150
         Top             =   3400
         Width           =   8355
         _ExtentX        =   14737
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
         Caption         =   "frmOptions.frx":0C8E
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblExcludeHWID 
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   2535
         Width           =   8355
         _ExtentX        =   14737
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
         BackStyle       =   0
         Caption         =   "Исключать следующие HWID (перечислять через "";"") из обработки (поддерживается маска ""*""):"
      End
   End
   Begin prjDIADBS.ctlJCFrames frDesign 
      Height          =   5295
      Left            =   4065
      Top             =   1560
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Оформление"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ctlColorButton ctlStatusBtnBackColor 
         Height          =   330
         Left            =   6720
         TabIndex        =   80
         Top             =   3780
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   582
         Icon            =   "frmOptions.frx":0D28
         BackColor       =   14016736
      End
      Begin prjDIADBS.ComboBoxW cmbButtonStyleColor 
         Height          =   315
         Left            =   3960
         TabIndex        =   79
         Top             =   3780
         Width           =   2595
         _ExtentX        =   4233
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
         Text            =   "frmOptions.frx":12AE
         CueBanner       =   "frmOptions.frx":12CE
         Sorted          =   -1  'True
      End
      Begin prjDIADBS.ComboBoxW cmbButtonStyle 
         Height          =   315
         Left            =   600
         TabIndex        =   78
         Top             =   3780
         Width           =   2595
         _ExtentX        =   4233
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
         Text            =   "frmOptions.frx":12EE
         CueBanner       =   "frmOptions.frx":130E
         Sorted          =   -1  'True
      End
      Begin VB.PictureBox imgOK 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0FFFF&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   7920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin prjDIADBS.CheckBoxW chkFutureButton 
         Height          =   210
         Left            =   450
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2400
         Width           =   210
         _ExtentX        =   370
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
         Caption         =   "frmOptions.frx":132E
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ComboBoxW cmbImageMain 
         Height          =   315
         Left            =   615
         TabIndex        =   84
         Top             =   4845
         Width           =   3000
         _ExtentX        =   5292
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
         Locked          =   -1  'True
         Text            =   "frmOptions.frx":134E
         CueBanner       =   "frmOptions.frx":136E
         Sorted          =   -1  'True
      End
      Begin prjDIADBS.ComboBoxW cmbImageStatus 
         Height          =   315
         Left            =   3960
         TabIndex        =   85
         Top             =   4845
         Width           =   3000
         _ExtentX        =   4233
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
         Locked          =   -1  'True
         Text            =   "frmOptions.frx":138E
         CueBanner       =   "frmOptions.frx":13AE
         Sorted          =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkButtonTextUpCase 
         Height          =   210
         Left            =   3510
         TabIndex        =   67
         Top             =   1530
         Width           =   4920
         _ExtentX        =   8678
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
         Caption         =   "frmOptions.frx":13CE
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.SpinBox txtButtonHeight 
         Height          =   255
         Left            =   1485
         TabIndex        =   58
         Top             =   765
         Width           =   1575
         _ExtentX        =   2778
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
         Max             =   3000
         AllowOnlyNumbers=   -1  'True
      End
      Begin prjDIADBS.SpinBox txtButtonWidth 
         Height          =   255
         Left            =   1485
         TabIndex        =   62
         Top             =   1125
         Width           =   1575
         _ExtentX        =   2778
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
         Max             =   3000
         AllowOnlyNumbers=   -1  'True
      End
      Begin prjDIADBS.SpinBox txtButton2BtnL 
         Height          =   255
         Left            =   6405
         TabIndex        =   60
         Top             =   765
         Width           =   1575
         _ExtentX        =   2778
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
         Max             =   1000
         AllowOnlyNumbers=   -1  'True
      End
      Begin prjDIADBS.SpinBox txtButton2BtnT 
         Height          =   255
         Left            =   6405
         TabIndex        =   64
         Top             =   1125
         Width           =   1575
         _ExtentX        =   2778
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
         Max             =   1000
         AllowOnlyNumbers=   -1  'True
      End
      Begin prjDIADBS.SpinBox txtButtonLeft 
         Height          =   255
         Left            =   1485
         TabIndex        =   66
         Top             =   1470
         Width           =   1575
         _ExtentX        =   2778
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
         Max             =   2000
         AllowOnlyNumbers=   -1  'True
      End
      Begin prjDIADBS.SpinBox txtButtonTop 
         Height          =   255
         Left            =   1485
         TabIndex        =   70
         Top             =   1815
         Width           =   1575
         _ExtentX        =   2778
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
         Max             =   2000
         AllowOnlyNumbers=   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorButton 
         Height          =   645
         Left            =   3480
         TabIndex        =   73
         Top             =   2145
         Width           =   2400
         _ExtentX        =   4313
         _ExtentY        =   1138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Установить цвет и шрифт текста кнопки"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkButtonDisable 
         Height          =   390
         Left            =   390
         TabIndex        =   75
         Top             =   2940
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":1442
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorToolTip 
         Height          =   645
         Left            =   6060
         TabIndex        =   74
         Top             =   2145
         Width           =   2400
         _ExtentX        =   4313
         _ExtentY        =   1138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Установить цвет и шрифт текста подсказок"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFutureButton 
         Height          =   555
         Left            =   390
         TabIndex        =   71
         Top             =   2220
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         CheckExist      =   -1  'True
         BackColor       =   14933984
         Caption         =   "Кнопка пакета драйверов"
         CaptionEffects  =   0
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         ColorScheme     =   2
      End
      Begin prjDIADBS.LabelW lblButtonStyleColor 
         Height          =   315
         Left            =   3960
         TabIndex        =   77
         Top             =   3480
         Width           =   3495
         _ExtentX        =   6165
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
         BackStyle       =   0
         Caption         =   "Стиль оформления кнопки"
      End
      Begin prjDIADBS.LabelW lblButtonStyle 
         Height          =   315
         Left            =   660
         TabIndex        =   76
         Top             =   3480
         Width           =   3075
         _ExtentX        =   5424
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
         BackStyle       =   0
         Caption         =   "Стиль оформления кнопки"
      End
      Begin prjDIADBS.ToolTip TT 
         Left            =   3480
         Top             =   1800
         _ExtentX        =   450
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Title           =   "frmOptions.frx":14B8
      End
      Begin prjDIADBS.LabelW lblTheme 
         Height          =   225
         Left            =   360
         TabIndex        =   81
         Top             =   4260
         Width           =   7875
         _ExtentX        =   13150
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Набор оформления программы (изменение основных иконок, и иконок статуса кнопок)"
      End
      Begin prjDIADBS.LabelW lblImageStatus 
         Height          =   255
         Left            =   3960
         TabIndex        =   83
         Top             =   4545
         Width           =   3000
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
         Caption         =   "Иконки для кнопок статуса"
      End
      Begin prjDIADBS.LabelW lblImageMain 
         Height          =   255
         Left            =   615
         TabIndex        =   82
         Top             =   4545
         Width           =   3000
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
         Caption         =   "Основные картинки"
      End
      Begin prjDIADBS.LabelW lblButtonWidth 
         Height          =   210
         Left            =   630
         TabIndex        =   61
         Top             =   1125
         Width           =   645
         _ExtentX        =   1270
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
         BackStyle       =   0
         Caption         =   "Ширина:"
         AutoSize        =   -1  'True
      End
      Begin prjDIADBS.LabelW lblButtonHeight 
         Height          =   210
         Left            =   630
         TabIndex        =   57
         Top             =   765
         Width           =   630
         _ExtentX        =   1191
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
         BackStyle       =   0
         Caption         =   "Высота:"
         AutoSize        =   -1  'True
      End
      Begin prjDIADBS.LabelW lblButtonTop 
         Height          =   210
         Left            =   630
         TabIndex        =   68
         Top             =   1815
         Width           =   615
         _ExtentX        =   1191
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
         BackStyle       =   0
         Caption         =   "Сверху:"
         AutoSize        =   -1  'True
      End
      Begin prjDIADBS.LabelW lblButtonLeft 
         Height          =   210
         Left            =   630
         TabIndex        =   65
         Top             =   1470
         Width           =   525
         _ExtentX        =   1032
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
         BackStyle       =   0
         Caption         =   "Слева:"
         AutoSize        =   -1  'True
      End
      Begin prjDIADBS.LabelW lblButton2BtnT 
         Height          =   255
         Left            =   3525
         TabIndex        =   63
         Top             =   1125
         Width           =   2865
         _ExtentX        =   5054
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
         BackStyle       =   0
         Caption         =   "Интервал по вертикали:"
         AutoSize        =   -1  'True
      End
      Begin prjDIADBS.LabelW lblButton2BtnL 
         Height          =   255
         Left            =   3525
         TabIndex        =   59
         Top             =   765
         Width           =   2850
         _ExtentX        =   5027
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
         BackStyle       =   0
         Caption         =   "Интервал по горизонтали:"
         AutoSize        =   -1  'True
      End
      Begin prjDIADBS.LabelW lblSizeButton 
         Height          =   210
         Left            =   390
         TabIndex        =   56
         Top             =   465
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Свойства кнопок"
      End
   End
   Begin prjDIADBS.ctlJCFrames frDesign2 
      Height          =   5295
      Left            =   4260
      Top             =   1845
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Оформление 2"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.SpinBox txtTabPerRowCount 
         Height          =   255
         Left            =   3330
         TabIndex        =   95
         Top             =   1755
         Width           =   675
         _ExtentX        =   1191
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
         Min             =   2
         Max             =   20
         Value           =   2
         TextAlignment   =   1
      End
      Begin prjDIADBS.CheckBoxW chkTabBlock 
         Height          =   210
         Left            =   390
         TabIndex        =   96
         Top             =   2085
         Width           =   7995
         _ExtentX        =   14102
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
         Caption         =   "frmOptions.frx":14D8
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkTabHide 
         Height          =   210
         Left            =   390
         TabIndex        =   97
         Top             =   2400
         Width           =   7995
         _ExtentX        =   14102
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
         Caption         =   "frmOptions.frx":159A
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkLoadUnSupportedOS 
         Height          =   210
         Left            =   390
         TabIndex        =   98
         Top             =   2715
         Width           =   7995
         _ExtentX        =   14102
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
         Caption         =   "frmOptions.frx":1648
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorTabOS 
         Height          =   795
         Left            =   390
         TabIndex        =   99
         Top             =   3030
         Width           =   2850
         _ExtentX        =   5027
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
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Установить цвет и шрифт текста закладки"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorTabDrivers 
         Height          =   795
         Left            =   390
         TabIndex        =   101
         Top             =   4320
         Width           =   2850
         _ExtentX        =   5027
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
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Установить цвет и шрифт текста закладки"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkFormMaximaze 
         Height          =   210
         Left            =   3315
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   720
         Width           =   4920
         _ExtentX        =   8678
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
         Caption         =   "frmOptions.frx":16C4
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.SpinBox txtFormHeight 
         Height          =   255
         Left            =   1275
         TabIndex        =   88
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
         Max             =   25000
         AllowOnlyNumbers=   -1  'True
      End
      Begin prjDIADBS.SpinBox txtFormWidth 
         Height          =   255
         Left            =   1275
         TabIndex        =   91
         Top             =   1065
         Width           =   1575
         _ExtentX        =   2778
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
         Max             =   25000
         AllowOnlyNumbers=   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkFormSizeSave 
         Height          =   210
         Left            =   3315
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1065
         Width           =   4920
         _ExtentX        =   8678
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
         Caption         =   "frmOptions.frx":172A
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblSizeForm 
         Height          =   210
         Left            =   180
         TabIndex        =   86
         Top             =   420
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Размеры основного окна"
      End
      Begin prjDIADBS.LabelW lblFormHeight 
         Height          =   210
         Left            =   435
         TabIndex        =   87
         Top             =   720
         Width           =   630
         _ExtentX        =   1191
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
         BackStyle       =   0
         Caption         =   "Высота:"
         AutoSize        =   -1  'True
      End
      Begin prjDIADBS.LabelW lblFormWidth 
         Height          =   210
         Left            =   435
         TabIndex        =   90
         Top             =   1065
         Width           =   645
         _ExtentX        =   1270
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
         BackStyle       =   0
         Caption         =   "Ширина:"
         AutoSize        =   -1  'True
      End
      Begin prjDIADBS.LabelW lblTabPerRowCount 
         Height          =   225
         Left            =   390
         TabIndex        =   94
         Top             =   1755
         Width           =   2730
         _ExtentX        =   5054
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
         BackStyle       =   0
         Caption         =   "Кол-во вкладок ОС на одну строку: "
         AutoSize        =   -1  'True
      End
      Begin prjDIADBS.LabelW lblTabControl 
         Height          =   225
         Left            =   150
         TabIndex        =   93
         Top             =   1440
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "TabControl - Поддерживаемые ОС"
      End
      Begin prjDIADBS.LabelW lblTabControl2 
         Height          =   225
         Left            =   120
         TabIndex        =   100
         Top             =   3960
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "TabControl 2 - Группы драйверов"
      End
   End
   Begin prjDIADBS.ctlJCFrames frDpInstParam 
      Height          =   5295
      Left            =   4440
      Top             =   2160
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Параметры запуска DPInst"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin VB.CommandButton cmdLegacyMode 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   105
         ToolTipText     =   "More on MSDN..."
         Top             =   660
         Width           =   255
      End
      Begin VB.CommandButton cmdPromptIfDriverIsNotBetter 
         Caption         =   "?"
         Height          =   255
         Left            =   2640
         TabIndex        =   108
         ToolTipText     =   "More on MSDN..."
         Top             =   1305
         Width           =   255
      End
      Begin VB.CommandButton cmdForceIfDriverIsNotBetter 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   111
         ToolTipText     =   "More on MSDN..."
         Top             =   1905
         Width           =   255
      End
      Begin VB.CommandButton cmdSuppressAddRemovePrograms 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   114
         ToolTipText     =   "More on MSDN..."
         Top             =   2460
         Width           =   255
      End
      Begin VB.CommandButton cmdSuppressWizard 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   117
         ToolTipText     =   "More on MSDN..."
         Top             =   2955
         Width           =   255
      End
      Begin VB.CommandButton cmdQuietInstall 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   120
         ToolTipText     =   "More on MSDN..."
         Top             =   3510
         Width           =   255
      End
      Begin VB.CommandButton cmdScanHardware 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   123
         ToolTipText     =   "More on MSDN..."
         Top             =   4005
         Width           =   255
      End
      Begin prjDIADBS.TextBoxW txtCmdStringDPInst 
         Height          =   330
         Left            =   2895
         TabIndex        =   126
         Top             =   4875
         Width           =   5535
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
         Text            =   "frmOptions.frx":178E
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":17AE
      End
      Begin prjDIADBS.CheckBoxW chkLegacyMode 
         Height          =   210
         Left            =   120
         TabIndex        =   104
         Top             =   660
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":17CE
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkPromptIfDriverIsNotBetter 
         Height          =   210
         Left            =   120
         TabIndex        =   107
         Top             =   1305
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":1802
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkForceIfDriverIsNotBetter 
         Height          =   210
         Left            =   120
         TabIndex        =   110
         Top             =   1905
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":1854
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSuppressAddRemovePrograms 
         CausesValidation=   0   'False
         Height          =   210
         Left            =   120
         TabIndex        =   113
         Top             =   2460
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":18A4
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSuppressWizard 
         Height          =   210
         Left            =   120
         TabIndex        =   116
         Top             =   2955
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":18F6
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkQuietInstall 
         Height          =   210
         Left            =   120
         TabIndex        =   119
         Top             =   3510
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":1932
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkScanHardware 
         Height          =   210
         Left            =   120
         TabIndex        =   122
         Top             =   4005
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":196A
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblCmdStringDPInst 
         Height          =   210
         Left            =   135
         TabIndex        =   125
         Top             =   4875
         Width           =   2685
         _ExtentX        =   4736
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
         BackStyle       =   0
         Caption         =   "Итоговые параметры запуска "
      End
      Begin prjDIADBS.LabelW lblDescription 
         Height          =   255
         Left            =   2865
         TabIndex        =   103
         Top             =   350
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Описание  параметра"
      End
      Begin prjDIADBS.LabelW lblParam 
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   350
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Параметр"
      End
      Begin prjDIADBS.LabelW lblPromptIfDriverIsNotBetter 
         Height          =   570
         Left            =   2925
         TabIndex        =   109
         Top             =   1305
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   1005
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
         Caption         =   "display a dialog box if a new driver is not a better match to a device than a driver that is currently installed on the device"
      End
      Begin prjDIADBS.LabelW lblLegacyMode 
         Height          =   645
         Left            =   2925
         TabIndex        =   106
         Top             =   660
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   1138
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
         Caption         =   "install unsigned drivers and driver packages that have missing files"
      End
      Begin prjDIADBS.LabelW lblForceIfDriverIsNotBetter 
         Height          =   510
         Left            =   2925
         TabIndex        =   112
         Top             =   1905
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   900
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
         Caption         =   "install a driver on a device even if the driver that is currently installed on the device is a better match than the new driver"
      End
      Begin prjDIADBS.LabelW lblSuppressAddRemovePrograms 
         Height          =   450
         Left            =   2925
         TabIndex        =   115
         Top             =   2460
         Width           =   5580
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
         Caption         =   "suppress the addition of Add or Remove Programs entries that represent the drivers and driver package"
      End
      Begin prjDIADBS.LabelW lblSuppressWizard 
         Height          =   450
         Left            =   2925
         TabIndex        =   118
         Top             =   2955
         Width           =   5550
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
         Caption         =   "configures DPInst to suppress the display of wizard pages and other user messages that DPInst generates."
      End
      Begin prjDIADBS.LabelW lblQuietInstall 
         Height          =   450
         Left            =   2925
         TabIndex        =   121
         Top             =   3510
         Width           =   5550
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
         Caption         =   "configures DPInst to suppress the display of wizard pages and most other user messages."
      End
      Begin prjDIADBS.LabelW lblScanHardware 
         Height          =   900
         Left            =   2925
         TabIndex        =   124
         Top             =   4005
         Width           =   5550
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
         Caption         =   $"frmOptions.frx":19A2
      End
   End
   Begin prjDIADBS.ctlJCFrames frDebug 
      Height          =   5295
      Left            =   4620
      Top             =   2460
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Отладочный режим"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.TextBoxW txtDebugLogName 
         Height          =   315
         Left            =   480
         TabIndex        =   136
         Top             =   2520
         Width           =   7815
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
         Text            =   "frmOptions.frx":1AA0
         CueBanner       =   "frmOptions.frx":1AE4
      End
      Begin prjDIADBS.TextBoxW txtMacrosDateDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   148
         Top             =   4905
         Width           =   1500
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
         Text            =   "frmOptions.frx":1B04
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1B30
      End
      Begin prjDIADBS.TextBoxW txtMacrosOSBITDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   146
         Top             =   4545
         Width           =   1500
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
         Text            =   "frmOptions.frx":1B50
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1B7E
      End
      Begin prjDIADBS.TextBoxW txtMacrosOSVERDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   144
         Top             =   4185
         Width           =   1500
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
         Text            =   "frmOptions.frx":1B9E
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1BCC
      End
      Begin prjDIADBS.TextBoxW txtMacrosPCModelDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   142
         Top             =   3825
         Width           =   1500
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
         Text            =   "frmOptions.frx":1BEC
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1C1E
      End
      Begin prjDIADBS.TextBoxW txtMacrosPCNameDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   140
         Top             =   3465
         Width           =   1500
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
         Text            =   "frmOptions.frx":1C3E
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1C6E
      End
      Begin prjDIADBS.CheckBoxW chkDebug 
         Height          =   210
         Left            =   495
         TabIndex        =   129
         Top             =   750
         Width           =   4440
         _ExtentX        =   7832
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
         Caption         =   "frmOptions.frx":1C8E
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlUcPickBox ucDebugLogPath 
         Height          =   315
         Left            =   480
         TabIndex        =   134
         Top             =   1890
         Width           =   7845
         _ExtentX        =   10821
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         DefaultExt      =   ""
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.CheckBoxW chkDebugLog2AppPath 
         Height          =   210
         Left            =   495
         TabIndex        =   132
         Top             =   1350
         Width           =   7920
         _ExtentX        =   11245
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
         Caption         =   "frmOptions.frx":1CDE
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkDebugTime2File 
         Height          =   210
         Left            =   495
         TabIndex        =   131
         Top             =   1050
         Width           =   7920
         _ExtentX        =   11245
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
         Caption         =   "frmOptions.frx":1D5E
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.SpinBox txtDebugLogLevel 
         Height          =   255
         Left            =   7680
         TabIndex        =   130
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
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
         Min             =   1
         Value           =   1
      End
      Begin prjDIADBS.LabelW lblDebugLogLevel 
         Height          =   255
         Left            =   4680
         TabIndex        =   128
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Уровень отладки:"
      End
      Begin prjDIADBS.LabelW lblMacrosDateDebug 
         Height          =   375
         Left            =   2400
         TabIndex        =   149
         Top             =   4905
         Width           =   5775
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
         Caption         =   "Дата и время создания лог-файла"
      End
      Begin prjDIADBS.LabelW lblMacrosOSBitDebug 
         Height          =   375
         Left            =   2400
         TabIndex        =   147
         Top             =   4545
         Width           =   5775
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
         Caption         =   "Архитектура операционной системы, в виде x32[64]"
      End
      Begin prjDIADBS.LabelW lblMacrosOSVerDebug 
         Height          =   375
         Left            =   2400
         TabIndex        =   145
         Top             =   4185
         Width           =   5775
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
         Caption         =   "Версия операционной системы в виде wnt5[6]"
      End
      Begin prjDIADBS.LabelW lblMacrosPCModelDebug 
         Height          =   375
         Left            =   2400
         TabIndex        =   143
         Top             =   3825
         Width           =   5775
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
         Caption         =   "Модель компьютера/материнской платы"
      End
      Begin prjDIADBS.LabelW lblMacrosParamDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   138
         Top             =   3150
         Width           =   1755
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Параметр"
      End
      Begin prjDIADBS.LabelW lblMacrosDescriptionDebug 
         Height          =   255
         Left            =   2400
         TabIndex        =   139
         Top             =   3150
         Width           =   5865
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Описание  параметра"
      End
      Begin prjDIADBS.LabelW lblMacrosPCNameDebug 
         Height          =   375
         Left            =   2400
         TabIndex        =   141
         Top             =   3465
         Width           =   5775
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
         Caption         =   "Краткое имя компьютера, без доменного суффикса"
      End
      Begin prjDIADBS.LabelW lblMacrosTypeDebug 
         Height          =   285
         Left            =   480
         TabIndex        =   137
         Top             =   2865
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Доступные макроподстановки для имени лог-файла:"
      End
      Begin prjDIADBS.LabelW lblDebugLogPath 
         Height          =   285
         Left            =   480
         TabIndex        =   133
         Top             =   1575
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Каталог для создания log-файлов:"
      End
      Begin prjDIADBS.LabelW lblDebug 
         Height          =   270
         Left            =   240
         TabIndex        =   127
         Top             =   420
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Настройки отладочного режима"
      End
      Begin prjDIADBS.LabelW lblDebugLogName 
         Height          =   285
         Left            =   495
         TabIndex        =   135
         Top             =   2225
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Каталог для создания log-файлов:"
      End
   End
   Begin prjDIADBS.ctlJCFrames frOther 
      Height          =   5295
      Left            =   4830
      Top             =   2775
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strItemOptions1     As String 'Основные настройки
Private strItemOptions2     As String 'Поддерживаемые ОС
Private strItemOptions3     As String 'Рабочие утилиты
Private strItemOptions4     As String 'Вспомогательные утилиты
Private strItemOptions5     As String 'Оформление программы
Private strItemOptions6     As String 'Параметры запуска DPInst
Private strItemOptions8     As String 'Основные настройки 2
Private strItemOptions9     As String 'Оформление программы 2
Private strItemOptions10    As String 'Отладочный режим
Private strTableOSHeader1   As String
Private strTableOSHeader2   As String
Private strTableOSHeader3   As String
Private strTableOSHeader4   As String
Private strTableOSHeader5   As String
Private strTableOSHeader6   As String
Private strTableOSHeader7   As String
Private strTableOSHeader8   As String
Private strTableOSHeader9   As String
Private strTableUtilHeader1 As String
Private strTableUtilHeader2 As String
Private strTableUtilHeader3 As String
Private strTableUtilHeader4 As String
Private strFormName         As String

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
'! Procedure   (Функция)   :   Sub ChangeButtonProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ChangeButtonProperties()

    With cmdFutureButton
        .Width = txtButtonWidth.Value
        .Height = txtButtonHeight.Value

        If chkButtonTextUpCase.Value Then
            .Caption = UCase$(LocaliseString(strPCLangCurrentPath, strFormName, "cmdFutureButton", .Caption))
        Else
            .Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdFutureButton", .Caption)
        End If

    End With

    With chkFutureButton
        .Left = cmdFutureButton.Left + 50
        .Top = cmdFutureButton.Top + (txtButtonHeight.Value - .Height) / 2
    End With

    SetBtnStatusFontProperties cmdFutureButton
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DebugCtlEnable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub DebugCtlEnable(ByVal mbEnable As Boolean)
    chkDebugTime2File.Enabled = mbEnable
    txtDebugLogName.Enabled = mbEnable
    ucDebugLogPath.Enabled = mbEnable
    chkDebugLog2AppPath.Enabled = mbEnable
    txtDebugLogLevel.Enabled = mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DebugCtlEnableLog2App
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub DebugCtlEnableLog2App(ByVal mbEnable As Boolean)
    ucDebugLogPath.Enabled = mbEnable
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
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FormLoadAction
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub FormLoadAction()

    ' Локализация приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    ' загрузить список опций
    LoadList_lvOptions
    ' Заполнить опции
    ReadOptions
    ' установить опции шрифта и цвета для будущей кнопки
    Set cmdFutureButton.PictureNormal = imgOK.Picture
    cmdFutureButton.ForeColor = lngFontBtn_Color
    SetBtnStatusFontProperties cmdFutureButton
    'Загрузить подсказку
    LoadToolTip
    DoEvents
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub InitializeObjectProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub InitializeObjectProperties()
    cmdFutureButton.CheckExist = True
    chkFutureButton.ZOrder 0
    
    cmbButtonStyle.ListIndex = lngStatusBtnStyle
    cmbButtonStyleColor.ListIndex = lngStatusBtnStyleColor
    ctlStatusBtnBackColor.Value = lngStatusBtnBackColor
    
    ' изменение шрифта и текста
    ChangeButtonProperties
End Sub

Private Sub LoadComboBtnStyle()
    
    With cmbButtonStyle
        .Clear
        .AddItem "Standard", 0
        .AddItem "Flat", 1
        .AddItem "WindowsXP", 2
        .AddItem "VistaAero", 3
        .AddItem "OfficeXP", 4
        .AddItem "Office2003", 5
        .AddItem "XPToolbar", 6
        .AddItem "VistaToolbar", 7
        .AddItem "Outlook2007", 8
        .AddItem "InstallShield", 9
        .AddItem "GelButton", 10
        .AddItem "3DHover", 11
        .AddItem "FlatHover", 12
        .AddItem "WindowsTheme", 13
    End With
    
    With cmbButtonStyleColor
        .Clear
        .AddItem "Blue", 0
        .AddItem "OliveGreen", 1
        .AddItem "Silver", 2
        .AddItem "Custom", 3
    End With
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadSkinListCombo
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   cmbName (ComboBox)
'                              strImagePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadSkinListCombo(cmbName As Object, strImagePath As String)

    Dim strListFolderTemp() As FindListStruct
    Dim i                   As Integer

    strListFolderTemp = SearchFoldersInRoot(strImagePath, "*")

    With cmbName
        .Clear

        For i = 0 To UBound(strListFolderTemp)
            .AddItem strListFolderTemp(i).Name, i
        Next

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadToolTip
'! Description (Описание)  :   [Инициализация подсказки для "будущей" кнопки]
'! Parameters  (Переменные):   ColumnHeader (LvwColumnHeader)
'!--------------------------------------------------------------------------------
Private Sub LoadToolTip()
Dim strTTText As String
Dim strTTipTextTitle As String

    strTTipTextTitle = LocaliseString(strPCLangCurrentPath, frmMain.Name, "ToolTipTextTitle", "Файл пакета драйверов:")
    strTTText = "d:\DIA\driverpacks.net\All\" & str2vbNewLine & _
                "dp_chipset_wnt5_x86-32_1209.7z" & vbNewLine & _
                "File Size: 4,33 МБ" & vbNewLine & _
                "Class of the Drivers: System" & str2vbNewLine & _
                "DRIVERS AVAILABLE TO INSTALL:" & vbNewLine & _
                "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbNewLine & _
                "HWID                  | PATH      | INFFILE      | VERSION(DB)            | ! | VERSION(PC)              | STATUS | DEVICE NAME                                                                                " & vbNewLine & _
                "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbNewLine & _
                "*PNP0103              | D\C\I\    | dmi_pci.inf  | 11/07/2008,7.0.1.1011  | > | 07/01/2001,5.1.2600.5512 | 0      | High Precision Event Timer                                                                 " & vbNewLine & _
                "PCI\VEN_8086&DEV_0100 | D\C\I\    | snb2009.inf  | 03/10/2011,9.2.0.1026  | < | 09/10/2012,9.2.0.1031    | 1      | 2nd Generation Intel(R) Core(TM) Processor Family DRAM Controller - 0100                   " & vbNewLine & _
                "PCI\VEN_8086&DEV_1C10 | D\C\I\    | cougcore.inf | 11/20/2010,9.2.0.1016  | < | 09/10/2012,9.2.0.1031    | 1      | Intel(R) 6 Series/C200 Series Chipset Family PCI Express Root Port 1 - 1C10                " & vbNewLine & _
                "PCI\VEN_8086&DEV_1C1A | D\C\I\    | cougcore.inf | 11/20/2010,9.2.0.1016  | < | 09/10/2012,9.2.0.1031    | 1      | Intel(R) 6 Series/C200 Series Chipset Family PCI Express Root Port 6 - 1C1A                " & vbNewLine & _
                "PCI\VEN_8086&DEV_1C1C | D\C\I\    | cougcore.inf | 11/20/2010,9.2.0.1016  | < | 09/10/2012,9.2.0.1031    | 1      | Intel(R) 6 Series/C200 Series Chipset Family PCI Express Root Port 7 - 1C1C                " & vbNewLine & _
                "PCI\VEN_8086&DEV_1C1E | D\C\I\    | cougcore.inf | 11/20/2010,9.2.0.1016  | < | 09/10/2012,9.2.0.1031    | 1      | Intel(R) 6 Series/C200 Series Chipset Family PCI Express Root Port 8 - 1C1E                " & vbNewLine & _
                "PCI\VEN_8086&DEV_1C22 | D\C\I\    | cougsmb.inf  | 09/10/2010,9.2.0.1011  | < | 09/10/2012,9.2.0.1031    | 0      | Intel(R) 6 Series/C200 Series Chipset Family SMBus Controller - 1C22                       " & vbNewLine & _
                "PCI\VEN_8086&DEV_1C24 | D\C\I\    | cougcore.inf | 11/20/2010,9.2.0.1016  | = | 11/20/2010,9.2.0.1016    | 0      | Intel(R) 6 Series/C200 Series Chipset Family Thermal Control - 1C24                        " & vbNewLine & _
                "PCI\VEN_8086&DEV_1C26 | D\C\I\    | cougusb.inf  | 07/31/2010,9.2.0.1031  | < | 09/10/2012,9.2.0.1031    | 1      | Intel(R) 6 Series/C200 Series Chipset Family USB Enhanced Host Controller - 1C26           " & vbNewLine & _
                "PCI\VEN_8086&DEV_1C2D | D\C\I\    | cougusb.inf  | 07/31/2010,9.2.0.1031  | < | 09/10/2012,9.2.0.1031    | 1      | Intel(R) 6 Series/C200 Series Chipset Family USB Enhanced Host Controller - 1C2D           " & vbNewLine & _
                "PCI\VEN_8086&DEV_1C3A | D\C\I6\   | heci.inf     | 09/22/2011,7.1.21.1134 | < | 12/17/2012,9.0.0.1287    | 1      | Intel(R) Management Engine Interface                                                       " & vbNewLine & _
                "PCI\VEN_8086&DEV_1C3A | D\C\I\    | cougme.inf   | 04/14/2011,1.2.0.1030  | < | 12/17/2012,9.0.0.1287    | 1      | Intel(R) 6 Series/C200 Series Management Engine Interface - 1C3A                           " & vbNewLine & _
                "PCI\VEN_8086&DEV_1C4A | D\C\I\    | cougcore.inf | 11/20/2010,9.2.0.1016  | < | 09/10/2012,9.2.0.1031    | 1      | Intel(R) H67 Express Chipset Family LPC Interface Controller - 1C4A                        " & vbNewLine & _
                "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

    ' Изменяем параметры Всплывающей подсказки для кнопки
    With TT
        .MaxTipWidth = lngRightWorkArea
        .SetDelayTime TipDelayTimeInitial, 400
        .SetDelayTime TipDelayTimeShow, 15000
        .Title = strTTipTextTitle
        .Tools.Add cmdFutureButton.hWnd, , strTTText
        SetTTFontProperties TT
    End With
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal strPathFile As String)

    Dim strFormNameMain As String

    strFormNameMain = frmMain.Name
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.CaptionW = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    frOptions.Caption = LocaliseString(strPathFile, strFormName, "frOptions", frOptions.Caption)
    ' Описание режимов
    optRezim_Intellect.Caption = LocaliseString(strPathFile, strFormNameMain, "RezimIntellect", optRezim_Intellect.Caption)
    optRezim_Ust.Caption = LocaliseString(strPathFile, strFormNameMain, "RezimUst", optRezim_Ust.Caption)
    optRezim_Upd.Caption = LocaliseString(strPathFile, strFormNameMain, "RezimUpd", optRezim_Upd.Caption)
    optRezim_Intellect.ToolTipText = LocaliseString(strPathFile, strFormNameMain, "RezimIntellectTip", optRezim_Intellect.ToolTipText)
    optRezim_Ust.ToolTipText = LocaliseString(strPathFile, strFormNameMain, "RezimUstTip", optRezim_Ust.ToolTipText)
    optRezim_Upd.ToolTipText = LocaliseString(strPathFile, strFormNameMain, "RezimUpdTip", optRezim_Upd.ToolTipText)
    strItemOptions1 = LocaliseString(strPathFile, strFormName, "ItemOptions1", "Основные настройки")
    strItemOptions2 = LocaliseString(strPathFile, strFormName, "ItemOptions2", "Поддерживаемые ОС")
    strItemOptions3 = LocaliseString(strPathFile, strFormName, "ItemOptions3", "Рабочие утилиты")
    strItemOptions4 = LocaliseString(strPathFile, strFormName, "ItemOptions4", "Вспомогательные утилиты")
    strItemOptions5 = LocaliseString(strPathFile, strFormName, "ItemOptions5", "Оформление программы")
    strItemOptions6 = LocaliseString(strPathFile, strFormName, "ItemOptions6", "Параметры запуска DPInst")
    strItemOptions8 = LocaliseString(strPathFile, strFormName, "ItemOptions8", "Основные настройки 2")
    strItemOptions9 = LocaliseString(strPathFile, strFormName, "ItemOptions9", "Оформление программы 2")
    strItemOptions10 = LocaliseString(strPathFile, strFormName, "ItemOptions10", "Отладочный режим")
    cmdOK.Caption = LocaliseString(strPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
    frMain.Caption = LocaliseString(strPathFile, strFormName, "frMain", frMain.Caption)
    lblOptionsStart.Caption = LocaliseString(strPathFile, strFormName, "lblOptionsStart", lblOptionsStart.Caption)
    chkUpdate.Caption = LocaliseString(strPathFile, strFormName, "chkUpdate", chkUpdate.Caption)
    chkUpdateBeta.Caption = LocaliseString(strPathFile, strFormName, "chkUpdateBeta", chkUpdateBeta.Caption)
    chkConvertDPName.Caption = LocaliseString(strPathFile, strFormName, "chkConvertDPName", chkConvertDPName.Caption)
    chkReadDPName.Caption = LocaliseString(strPathFile, strFormName, "chkReadDPName", chkReadDPName.Caption)
    chkHideOtherProcess.Caption = LocaliseString(strPathFile, strFormName, "chkHideOtherProcess", chkHideOtherProcess.Caption)
    lblOptionsTemp.Caption = LocaliseString(strPathFile, strFormName, "lblOptionsTemp", lblOptionsTemp.Caption)
    chkTempPath.Caption = LocaliseString(strPathFile, strFormName, "chkTempPath", chkTempPath.Caption)
    chkRemoveTemp.Caption = LocaliseString(strPathFile, strFormName, "chkRemoveTemp", chkRemoveTemp.Caption)
    lblDebug.Caption = LocaliseString(strPathFile, strFormName, "lblDebug", lblDebug.Caption)
    chkDebug.Caption = LocaliseString(strPathFile, strFormName, "chkDebug", chkDebug.Caption)
    lblRezim.Caption = LocaliseString(strPathFile, strFormName, "lblRezim", lblRezim.Caption)
    lblDebugLogPath.Caption = LocaliseString(strPathFile, strFormName, "lblDebugLogPath", lblDebugLogPath.Caption)
    frMainTools.Caption = LocaliseString(strPathFile, strFormName, "frMainTools", frMainTools.Caption)
    cmdPathDefault.Caption = LocaliseString(strPathFile, strFormName, "cmdPathDefault", cmdPathDefault.Caption)
    frOtherTools.Caption = LocaliseString(strPathFile, strFormName, "frOtherTools", frOtherTools.Caption)
    cmdAddUtil.Caption = LocaliseString(strPathFile, strFormName, "cmdAddUtil", cmdAddUtil.Caption)
    cmdEditUtil.Caption = LocaliseString(strPathFile, strFormName, "cmdEditUtil", cmdEditUtil.Caption)
    cmdDelUtil.Caption = LocaliseString(strPathFile, strFormName, "cmdDelUtil", cmdDelUtil.Caption)
    frOS.Caption = LocaliseString(strPathFile, strFormName, "frOS", frOS.Caption)
    chkRecursion.Caption = LocaliseString(strPathFile, strFormName, "chkRecursion", chkRecursion.Caption)
    chkLoadFinishFile.Caption = LocaliseString(strPathFile, strFormName, "chkLoadFinishFile", chkLoadFinishFile.Caption)
    cmdAddOS.Caption = LocaliseString(strPathFile, strFormName, "cmdAddOS", cmdAddOS.Caption)
    cmdEditOS.Caption = LocaliseString(strPathFile, strFormName, "cmdEditOS", cmdEditOS.Caption)
    cmdDelOS.Caption = LocaliseString(strPathFile, strFormName, "cmdDelOS", cmdDelOS.Caption)
    lblExcludeHWID.Caption = LocaliseString(strPathFile, strFormName, "lblExcludeHWID", lblExcludeHWID.Caption)
    frDesign.Caption = LocaliseString(strPathFile, strFormName, "frDesign", frDesign.Caption)
    frDesign2.Caption = LocaliseString(strPathFile, strFormName, "frDesign2", frDesign2.Caption)
    lblSizeForm.Caption = LocaliseString(strPathFile, strFormName, "lblSizeForm", lblSizeForm.Caption)
    lblFormHeight.Caption = LocaliseString(strPathFile, strFormName, "lblFormHeight", lblFormHeight.Caption)
    lblFormWidth.Caption = LocaliseString(strPathFile, strFormName, "lblFormWidth", lblFormWidth.Caption)
    chkFormMaximaze.Caption = LocaliseString(strPathFile, strFormName, "chkFormMaximaze", chkFormMaximaze.Caption)
    chkFormSizeSave.Caption = LocaliseString(strPathFile, strFormName, "chkFormSizeSave", chkFormSizeSave.Caption)
    lblSizeButton.Caption = LocaliseString(strPathFile, strFormName, "lblSizeButton", lblSizeButton.Caption)
    lblButtonHeight.Caption = LocaliseString(strPathFile, strFormName, "lblButtonHeight", lblButtonHeight.Caption)
    lblButtonWidth.Caption = LocaliseString(strPathFile, strFormName, "lblButtonWidth", lblButtonWidth.Caption)
    lblButtonLeft.Caption = LocaliseString(strPathFile, strFormName, "lblButtonLeft", lblButtonLeft.Caption)
    lblButtonTop.Caption = LocaliseString(strPathFile, strFormName, "lblButtonTop", lblButtonTop.Caption)
    lblButton2BtnL.Caption = LocaliseString(strPathFile, strFormName, "lblButton2BtnL", lblButton2BtnL.Caption)
    lblButton2BtnT.Caption = LocaliseString(strPathFile, strFormName, "lblButton2BtnT", lblButton2BtnT.Caption)
    chkButtonTextUpCase.Caption = LocaliseString(strPathFile, strFormName, "chkButtonTextUpCase", chkButtonTextUpCase.Caption)
    cmdFutureButton.Caption = LocaliseString(strPathFile, strFormName, "cmdFutureButton", cmdFutureButton.Caption)
    lblImageMain.Caption = LocaliseString(strPathFile, strFormName, "lblImageMain", lblImageMain.Caption)
    lblImageStatus.Caption = LocaliseString(strPathFile, strFormName, "lblImageStatus", lblImageStatus.Caption)
    lblTabControl.Caption = LocaliseString(strPathFile, strFormName, "lblTabControl", lblTabControl.Caption)
    lblTabControl2.Caption = LocaliseString(strPathFile, strFormName, "lblTabControl2", lblTabControl2.Caption)
    lblTabPerRowCount.Caption = LocaliseString(strPathFile, strFormName, "lblTabPerRowCount", lblTabPerRowCount.Caption)
    chkTabBlock.Caption = LocaliseString(strPathFile, strFormName, "chkTabBlock", chkTabBlock.Caption)
    chkTabHide.Caption = LocaliseString(strPathFile, strFormName, "chkTabHide", chkTabHide.Caption)
    frDpInstParam.Caption = LocaliseString(strPathFile, strFormName, "frDpInstParam", frDpInstParam.Caption)
    lblParam.Caption = LocaliseString(strPathFile, strFormName, "lblParam", lblParam.Caption)
    lblDescription.Caption = LocaliseString(strPathFile, strFormName, "lblDescription", lblDescription.Caption)
    lblLegacyMode.Caption = LocaliseString(strPathFile, strFormName, "lblLegacyMode", lblLegacyMode.Caption)
    lblPromptIfDriverIsNotBetter.Caption = LocaliseString(strPathFile, strFormName, "lblPromptIfDriverIsNotBetter", lblPromptIfDriverIsNotBetter.Caption)
    lblForceIfDriverIsNotBetter.Caption = LocaliseString(strPathFile, strFormName, "lblForceIfDriverIsNotBetter", lblForceIfDriverIsNotBetter.Caption)
    lblSuppressAddRemovePrograms.Caption = LocaliseString(strPathFile, strFormName, "lblSuppressAddRemovePrograms", lblSuppressAddRemovePrograms.Caption)
    lblSuppressWizard.Caption = LocaliseString(strPathFile, strFormName, "lblSuppressWizard", lblSuppressWizard.Caption)
    lblQuietInstall.Caption = LocaliseString(strPathFile, strFormName, "lblQuietInstall", lblQuietInstall.Caption)
    lblScanHardware.Caption = LocaliseString(strPathFile, strFormName, "lblScanHardware", lblScanHardware.Caption)
    lblCmdStringDPInst.Caption = LocaliseString(strPathFile, strFormName, "lblCmdStringDPInst", lblCmdStringDPInst.Caption)
    strTableOSHeader1 = LocaliseString(strPathFile, strFormName, "TableOSHeader1", "Версия")
    strTableOSHeader2 = LocaliseString(strPathFile, strFormName, "TableOSHeader2", "Наименование")
    strTableOSHeader3 = LocaliseString(strPathFile, strFormName, "TableOSHeader3", "Пакеты драйверов")
    strTableOSHeader4 = LocaliseString(strPathFile, strFormName, "TableOSHeader4", "База данных")
    strTableOSHeader5 = LocaliseString(strPathFile, strFormName, "TableOSHeader5", "x64")
    strTableOSHeader6 = LocaliseString(strPathFile, strFormName, "TableOSHeader6", "PhysX")
    strTableOSHeader7 = LocaliseString(strPathFile, strFormName, "TableOSHeader7", "Lang")
    strTableOSHeader8 = LocaliseString(strPathFile, strFormName, "TableOSHeader8", "ExludeFiles")
    strTableOSHeader9 = LocaliseString(strPathFile, strFormName, "TableOSHeader9", "ExludeFiles")
    strTableUtilHeader1 = LocaliseString(strPathFile, strFormName, "TableUtilHeader1", "Наименование")
    strTableUtilHeader2 = LocaliseString(strPathFile, strFormName, "TableUtilHeader2", "Путь")
    strTableUtilHeader3 = LocaliseString(strPathFile, strFormName, "TableUtilHeader3", "Параметр")
    strTableUtilHeader4 = LocaliseString(strPathFile, strFormName, "TableUtilHeader4", "Путь x64")
    frMain2.Caption = LocaliseString(strPathFile, strFormName, "frMain2", frMain2.Caption)
    lblCompareVersionDRV.Caption = LocaliseString(strPathFile, strFormName, "lblCompareVersionDRV", lblCompareVersionDRV.Caption)
    optCompareByDate.Caption = LocaliseString(strPathFile, strFormName, "optCompareByDate", optCompareByDate.Caption)
    optCompareByVersion.Caption = LocaliseString(strPathFile, strFormName, "optCompareByVersion", optCompareByVersion.Caption)
    txtCompareVersionDRV.Text = LocaliseString(strPathFile, strFormName, "txtCompareVersionDRV", txtCompareVersionDRV.Text)
    chkSilentDll.Caption = LocaliseString(strPathFile, strFormName, "chkSilentDll", chkSilentDll.Caption)
    chkDateFormatRus.Caption = LocaliseString(strPathFile, strFormName, "chkDateFormatRus", chkDateFormatRus.Caption)
    chkSearchOnStart.Caption = LocaliseString(strPathFile, strFormName, "chkSearchOnStart", chkSearchOnStart.Caption)
    lblPauseAfterSearch.Caption = LocaliseString(strPathFile, strFormName, "lblPauseAfterSearch", lblPauseAfterSearch.Caption)
    chkCreateRP.Caption = LocaliseString(strPathFile, strFormName, "chkCreateRP", chkCreateRP.Caption)
    chkCompatiblesHWID.Caption = LocaliseString(strPathFile, strFormName, "chkCompatiblesHWID", chkCompatiblesHWID.Caption)
    chkLoadUnSupportedOS.Caption = LocaliseString(strPathFile, strFormName, "chkLoadUnSupportedOS", chkLoadUnSupportedOS.Caption)
    chkDebugLog2AppPath.Caption = LocaliseString(strPathFile, strFormName, "chkDebugLog2AppPath", chkDebugLog2AppPath.Caption)
    frDebug.Caption = LocaliseString(strPathFile, strFormName, "frDebug", frDebug.Caption)
    lblMacrosTypeDebug.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosTypeDebug", lblMacrosTypeDebug.Caption)
    lblMacrosParamDebug.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosParamDebug", lblMacrosParamDebug.Caption)
    lblMacrosDescriptionDebug.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosDescriptionDebug", lblMacrosDescriptionDebug.Caption)
    lblMacrosPCNameDebug.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosPCNameDebug", lblMacrosPCNameDebug.Caption)
    lblMacrosPCModelDebug.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosPCModelDebug", lblMacrosPCModelDebug.Caption)
    lblMacrosOSVerDebug.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosOSVerDebug", lblMacrosOSVerDebug.Caption)
    lblMacrosOSBitDebug.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosOSBitDebug", lblMacrosOSBitDebug.Caption)
    lblMacrosDateDebug.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosDateDebug", lblMacrosDateDebug.Caption)
    chkDebugTime2File.Caption = LocaliseString(strPathFile, strFormName, "chkDebugTime2File", chkDebugTime2File.Caption)
    lblDebugLogName.Caption = LocaliseString(strPathFile, strFormName, "lblDebugLogName", lblDebugLogName.Caption)
    cmdFontColorButton.Caption = LocaliseString(strPathFile, strFormName, "cmdFontColorButton", cmdFontColorButton.Caption)
    cmdFontColorTabOS.Caption = LocaliseString(strPathFile, strFormName, "cmdFontColorTabOS", cmdFontColorTabOS.Caption)
    cmdFontColorTabDrivers.Caption = LocaliseString(strPathFile, strFormName, "cmdFontColorTabDrivers", cmdFontColorTabDrivers.Caption)
    chkButtonDisable.Caption = LocaliseString(strPathFile, strFormName, "chkButtonDisable", chkButtonDisable.Caption)
    lblTheme.Caption = LocaliseString(strPathFile, strFormName, "lblTheme", lblTheme.Caption)
    cmdFontColorToolTip.Caption = LocaliseString(strPathFile, strFormName, "cmdFontColorToolTip", cmdFontColorToolTip.Caption)
    lblDebugLogLevel.Caption = LocaliseString(strPathFile, strFormName, "lblDebugLogLevel", lblDebugLogLevel.Caption)
    ctlStatusBtnBackColor.DropDownCaption = LocaliseString(strPathFile, strFormName, "ctlStatusBtnBackColor", ctlStatusBtnBackColor.DropDownCaption)
    lblButtonStyle.Caption = LocaliseString(strPathFile, strFormName, "lblButtonStyle", lblButtonStyle.Caption)
    lblButtonStyleColor.Caption = LocaliseString(strPathFile, strFormName, "lblButtonStyleColor", lblButtonStyleColor.Caption)
    chkDP_is_aFolder.Caption = LocaliseString(strPathFile, strFormName, "chkDP_is_aFolder", chkDP_is_aFolder.Caption)
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ReadOptions
'! Description (Описание)  :   [Читаем настройки программы и заполняем поля]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ReadOptions()
    ' загрузить список ОС
    LoadList_lvOS
    ' загрузить список Утилит
    LoadList_lvUtils
    ' Остальные параметры
    chkUpdate.Value = mbUpdateCheck
    chkUpdateBeta.Value = mbUpdateCheckBeta
    chkSilentDll.Value = mbSilentDLL
    chkRemoveTemp.Value = mbDelTmpAfterClose
    chkDebug.Value = mbDebugStandart
    chkDebugTime2File.Value = mbDebugTime2File
    chkDebugLog2AppPath.Value = mbDebugLog2AppPath
    chkTabBlock.Value = mbTabBlock
    chkTabHide.Value = mbTabHide
    chkFormMaximaze.Value = mbStartMaximazed
    chkRecursion.Value = mbRecursion
    chkFormSizeSave.Value = mbSaveSizeOnExit
    chkTempPath.Value = mbTempPath
    ucTempPath.Path = strAlternativeTempPath
    chkHideOtherProcess.Value = mbHideOtherProcess
    ucDebugLogPath.Path = strDebugLogPathTemp
    txtDebugLogName.Text = strDebugLogNameTemp
    chkCreateRP.Value = mbCreateRestorePoint
    chkLoadUnSupportedOS.Value = mbLoadUnSupportedOS
    chkCompatiblesHWID.Value = mbCompatiblesHWID
    txtDebugLogLevel.Text = lngDetailMode
    chkDP_is_aFolder.Value = mbDP_Is_aFolder
    
    ' Режим при старте
    Select Case miStartMode

        Case 1
            optRezim_Upd.Value = False
            optRezim_Ust.Value = False
            optRezim_Intellect.Value = True

        Case 2
            optRezim_Upd.Value = False
            optRezim_Intellect.Value = False
            optRezim_Ust.Value = True

        Case 3
            optRezim_Ust.Value = False
            optRezim_Intellect.Value = False
            optRezim_Upd.Value = True
    End Select

    'MainForm
    txtFormHeight.Value = lngMainFormHeight
    txtFormWidth.Value = lngMainFormWidth
    txtExcludeHWID.Text = strExcludeHWID
    'Buttons
    txtButtonWidth.Value = lngButtonWidth
    txtButtonHeight.Value = lngButtonHeight
    txtButtonLeft.Value = lngButtonLeft
    txtButtonTop.Value = lngButtonTop
    txtButton2BtnL.Value = lngBtn2BtnLeft
    txtButton2BtnT.Value = lngBtn2BtnTop
    chkButtonTextUpCase.Value = mbButtonTextUpCase
    txtTabPerRowCount.Value = lngOSCountPerRow

    'Пути к программам
    If mbPatnAbs Then
        'Секция Devcon
        ucDevCon86Path.Path = strDevConExePath
        ucDevCon64Path.Path = strDevConExePath64
        ucDevCon86Pathw2k.Path = strDevConExePathW2k
        'Секция DPInst
        ucDPInst86Path.Path = strDPInstExePath86
        ucDPInst64Path.Path = strDPInstExePath64
        'Секция Arc
        ucArchPath.Path = strArh7zExePATH
        ucCmdDevconPath.Path = strDevconCmdPath
    Else
        'Секция Devcon
        ucDevCon86Path.Path = Replace$(strDevConExePath, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucDevCon64Path.Path = Replace$(strDevConExePath64, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucDevCon86Pathw2k.Path = Replace$(strDevConExePathW2k, strAppPathBackSL, vbNullString, , , vbTextCompare)
        'Секция DPInst
        ucDPInst86Path.Path = Replace$(strDPInstExePath86, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucDPInst64Path.Path = Replace$(strDPInstExePath64, strAppPathBackSL, vbNullString, , , vbTextCompare)
        'Секция Arc
        ucArchPath.Path = Replace$(strArh7zExePATH, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucCmdDevconPath.Path = Replace$(strDevconCmdPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
    End If

    ' Настройки DpInst
    chkLegacyMode.Value = mbDpInstLegacyMode
    chkPromptIfDriverIsNotBetter.Value = mbDpInstPromptIfDriverIsNotBetter
    chkForceIfDriverIsNotBetter.Value = mbDpInstForceIfDriverIsNotBetter
    chkSuppressAddRemovePrograms.Value = mbDpInstSuppressAddRemovePrograms
    chkSuppressWizard.Value = mbDpInstSuppressWizard
    chkQuietInstall.Value = mbDpInstQuietInstall
    chkScanHardware.Value = mbDpInstScanHardware
    ' Другие настройки
    txtCmdStringDPInst = CollectCmdString
    chkLoadFinishFile.Value = mbLoadFinishFile
    chkReadDPName.Value = mbReadDPName
    chkConvertDPName.Value = mbConvertDPName
    ' Загрузка списка скинов
    LoadSkinListCombo cmbImageMain, strPathImageMain
    LoadSkinListCombo cmbImageStatus, strPathImageStatusButton
    cmbImageMain.Text = strImageMainName
    cmbImageStatus.Text = strImageStatusButtonName
    LoadIconImage2Object imgOK, "BTN_OK", strPathImageStatusButtonWork
    ' Сравнение версий драйверов
    optCompareByDate.Value = mbCompareDrvVerByDate
    optCompareByVersion.Value = Not mbCompareDrvVerByDate
    ' Отображать дату в формате dd/mm/yyyy
    chkDateFormatRus.Value = mbDateFormatRus
    'поиск новых устройств при старте
    chkSearchOnStart.Value = mbSearchOnStart

    With txtPauseAfterSearch
        .Min = 0
        .Increment = 1
        .Value = lngPauseAfterSearch
    End With

    ' изменение активности элементов
    DebugCtlEnable CBool(chkDebug.Value)
    DebugCtlEnableLog2App Not CBool(chkDebugLog2AppPath.Value)
    TempCtlEnable CBool(chkTempPath.Value)
    UpdateCtlEnable CBool(chkUpdate.Value)
    TabCtlEnable Not CBool(chkTabHide.Value)
    Tab2CtlEnable CBool(chkTabBlock.Value)
    ' Инициализация параметров для изменения шрифта и цвета
    InitializeObjectProperties
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveOptions
'! Description (Описание)  :   [Сохранение настроек в ини-файл]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SaveOptions()

    Dim miRezim          As Long
    Dim cnt              As Long
    Dim lngOSCountNew    As Long
    Dim lngUtilsCountNew As Long
    Dim strSysIniTemp    As String
    Dim strLogNameTemp   As String

    If mbIsDriveCDRoom And Not mbLoadIniTmpAfterRestart Then
        If strSysIni <> strWorkTempBackSL & strSettingIniFile Then
            MsgBox strMessages(38), vbInformation + vbApplicationModal, strProductName

            Exit Sub

        End If

    ElseIf mbIsDriveCDRoom And mbLoadIniTmpAfterRestart Then
        strSysIniTemp = strWinTemp & "Settings_DIA_TMP.ini"
        SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP", True
        SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP_PATH", strSysIniTemp
    Else
        strSysIniTemp = strSysIni
    End If

    '**************************************************
    '***************** Запись настроек ****************
    '**************************************************
    ' Секция MAIN
    'Удаление TEMP при выходе
    IniWriteStrPrivate "Main", "DelTmpAfterClose", chkRemoveTemp.Value, strSysIniTemp
    ' Автообновление
    IniWriteStrPrivate "Main", "UpdateCheck", chkUpdate.Value, strSysIniTemp
    ' Автообновление Beta
    IniWriteStrPrivate "Main", "UpdateCheckBeta", chkUpdateBeta.Value, strSysIniTemp

    ' Режим запуска
    If optRezim_Intellect.Value Then
        miRezim = 1
    Else

        If optRezim_Ust.Value Then
            miRezim = 2
        Else

            If optRezim_Upd.Value Then
                miRezim = 3
            End If
        End If
    End If

    IniWriteStrPrivate "Main", "StartMode", miRezim, strSysIniTemp
    IniWriteStrPrivate "Main", "EULAAgree", Abs(mbEULAAgree), strSysIniTemp
    IniWriteStrPrivate "Main", "HideOtherProcess", chkHideOtherProcess.Value, strSysIniTemp
    IniWriteStrPrivate "Main", "AlternativeTemp", chkTempPath.Value, strSysIniTemp
    IniWriteStrPrivate "Main", "AlternativeTempPath", ucTempPath.Path, strSysIniTemp
    IniWriteStrPrivate "Main", "SilentDLL", chkSilentDll.Value, strSysIniTemp
    IniWriteStrPrivate "Main", "SearchOnStart", chkSearchOnStart.Value, strSysIniTemp
    IniWriteStrPrivate "Main", "PauseAfterSearch", txtPauseAfterSearch.Value, strSysIniTemp
    IniWriteStrPrivate "Main", "CreateRestorePoint", chkCreateRP.Value, strSysIniTemp
    IniWriteStrPrivate "Main", "IconMainSkin", cmbImageMain.Text, strSysIniTemp
    IniWriteStrPrivate "Main", "LoadIniTmpAfterRestart", Abs(mbLoadIniTmpAfterRestart), strSysIniTemp

    ' Секция Debug
    IniWriteStrPrivate "Debug", "DebugEnable", chkDebug.Value, strSysIniTemp
    IniWriteStrPrivate "Debug", "DebugLogPath", ucDebugLogPath.Path, strSysIniTemp
    strLogNameTemp = "DIA-LOG_%DATE%.txt"

    If LenB(txtDebugLogName.Text) Then
        If InStr(txtDebugLogName.Text, strDot) Then
            strLogNameTemp = txtDebugLogName.Text
        End If
    End If

    IniWriteStrPrivate "Debug", "DebugLogName", strLogNameTemp, strSysIniTemp
    IniWriteStrPrivate "Debug", "CleenHistory", 1, strSysIniTemp
    IniWriteStrPrivate "Debug", "Detailmode", txtDebugLogLevel.Text, strSysIniTemp
    IniWriteStrPrivate "Debug", "DebugLog2AppPath", chkDebugLog2AppPath.Value, strSysIniTemp
    IniWriteStrPrivate "Debug", "Time2File", Abs(mbDebugTime2File), strSysIniTemp
    
    'Секция Arc
    IniWriteStrPrivate "Arc", "PathExe", ucArchPath.Path, strSysIniTemp
    'Секция Devcon
    IniWriteStrPrivate "Devcon", "PathExe", ucDevCon86Path.Path, strSysIniTemp
    IniWriteStrPrivate "Devcon", "PathExe64", ucDevCon64Path.Path, strSysIniTemp
    IniWriteStrPrivate "Devcon", "PathExeW2K", ucDevCon86Pathw2k.Path, strSysIniTemp
    IniWriteStrPrivate "Devcon", "CollectHwidsCmd", ucCmdDevconPath.Path, strSysIniTemp
    'Секция DPInst
    IniWriteStrPrivate "DPInst", "PathExe", ucDPInst86Path.Path, strSysIniTemp
    IniWriteStrPrivate "DPInst", "PathExe64", ucDPInst64Path.Path, strSysIniTemp
    IniWriteStrPrivate "DPInst", "LegacyMode", chkLegacyMode.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "PromptIfDriverIsNotBetter", chkPromptIfDriverIsNotBetter.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "ForceIfDriverIsNotBetter", chkForceIfDriverIsNotBetter.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "SuppressAddRemovePrograms", chkSuppressAddRemovePrograms.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "SuppressWizard", chkSuppressWizard.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "QuietInstall", chkQuietInstall.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "ScanHardware", chkScanHardware.Value, strSysIniTemp
    
    'Секция OS
    'Число ОС
    lngOSCountNew = lvOS.ListItems.Count
    IniWriteStrPrivate "OS", "OSCount", lngOSCountNew, strSysIniTemp
    ' кол-во табов на строку
    IniWriteStrPrivate "OS", "OSCountPerRow", txtTabPerRowCount.Value, strSysIniTemp
    ' Рекурсивный перебор папок
    IniWriteStrPrivate "OS", "Recursion", chkRecursion.Value, strSysIniTemp
    ' Блокировать лишние вкладки
    IniWriteStrPrivate "OS", "TabBlock", chkTabBlock.Value, strSysIniTemp
    ' Скрывать лишние вкладки
    IniWriteStrPrivate "OS", "TabHide", chkTabHide.Value, strSysIniTemp
    ' Обрабатывать файлы Finish
    IniWriteStrPrivate "OS", "LoadFinishFile", chkLoadFinishFile.Value, strSysIniTemp
    ' Считывать класс пакета драйверов из Finish
    IniWriteStrPrivate "OS", "ReadClasses", Abs(mbReadClasses), strSysIniTemp
    ' Считывать класс пакета драйверов из Finish
    IniWriteStrPrivate "OS", "ReadDPName", chkReadDPName.Value, strSysIniTemp
    ' Считывать класс пакета драйверов из Finish
    IniWriteStrPrivate "OS", "ConvertDPName", chkConvertDPName.Value, strSysIniTemp
    IniWriteStrPrivate "OS", "ExcludeHWID", txtExcludeHWID.Text, strSysIniTemp
    ' Сравнение версий драйверов
    IniWriteStrPrivate "OS", "CompareDrvVerByDate", Abs(optCompareByDate.Value), strSysIniTemp
    IniWriteStrPrivate "OS", "DateFormatRus", chkDateFormatRus.Value, strSysIniTemp
    ' Обрабатывать совместимые HWID
    IniWriteStrPrivate "OS", "CompatiblesHWID", chkCompatiblesHWID.Value, strSysIniTemp
    IniWriteStrPrivate "OS", "CompatiblesHWIDCount", lngCompatiblesHWIDCount, strSysIniTemp
    ' Необрабатывать вкладки
    IniWriteStrPrivate "OS", "LoadUnSupportedOS", chkLoadUnSupportedOS.Value, strSysIniTemp
    IniWriteStrPrivate "OS", "CalcDriverScore", Abs(mbCalcDriverScore), strSysIniTemp
    IniWriteStrPrivate "OS", "SearchCompatibleDriverOtherOS", Abs(mbSearchCompatibleDriverOtherOS), strSysIniTemp
    IniWriteStrPrivate "OS", "MatchHWIDbyDPName", Abs(mbMatchHWIDbyDPName), strSysIniTemp
    IniWriteStrPrivate "OS", "DP_is_aFolder", chkDP_is_aFolder.Value, strSysIniTemp
    IniWriteStrPrivate "OS", "SortMethodShell", Abs(mbSortMethodShell), strSysIniTemp
    
    'Заполяем в цикле подсекции ОС
    For cnt = 1 To lngOSCountNew

        'Секция OS_N
        With lvOS.ListItems(cnt)
            IniWriteStrPrivate "OS_" & cnt, "Ver", .Text, strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "Name", .SubItems(1), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "drpFolder", .SubItems(2), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "devIDFolder", .SubItems(3), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "is64bit", .SubItems(4), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "ExcludeFileName", .SubItems(8), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "PathPhysX", .SubItems(5), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "PathLanguages", .SubItems(6), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "PathRuntimes", .SubItems(7), strSysIniTemp
        End With

    Next

    'Секция Utils
    'Число утилит
    lngUtilsCountNew = lvUtils.ListItems.Count
    IniWriteStrPrivate "Utils", "UtilsCount", lngUtilsCountNew, strSysIniTemp

    'Заполяем в цикле подсекции Утилиты
    For cnt = 1 To lngUtilsCountNew
        'Секция Utils_N
        With lvUtils.ListItems(cnt)
            IniWriteStrPrivate "Utils_" & cnt, "Name", .Text, strSysIniTemp
            IniWriteStrPrivate "Utils_" & cnt, "Path", .SubItems(1), strSysIniTemp
            IniWriteStrPrivate "Utils_" & cnt, "Path64", .SubItems(2), strSysIniTemp
            IniWriteStrPrivate "Utils_" & cnt, "Params", .SubItems(3), strSysIniTemp
        End With
    Next

    'Секция MainForm
    IniWriteStrPrivate "MainForm", "Width", txtFormWidth.Value, strSysIniTemp
    IniWriteStrPrivate "MainForm", "Height", txtFormHeight.Value, strSysIniTemp
    IniWriteStrPrivate "MainForm", "StartMaximazed", chkFormMaximaze.Value, strSysIniTemp
    mbSaveSizeOnExit = CBool(chkFormSizeSave.Value)
    IniWriteStrPrivate "MainForm", "SaveSizeOnExit", chkFormSizeSave.Value, strSysIniTemp
    IniWriteStrPrivate "MainForm", "FontName", strFontMainForm_Name, strSysIniTemp
    IniWriteStrPrivate "MainForm", "FontSize", lngFontMainForm_Size, strSysIniTemp
    IniWriteStrPrivate "MainForm", "HighlightColor", CStr(glHighlightColor), strSysIniTemp
    'Секция Buttons
    IniWriteStrPrivate "Button", "FontName", strFontBtn_Name, strSysIniTemp
    IniWriteStrPrivate "Button", "FontSize", miFontBtn_Size, strSysIniTemp
    IniWriteStrPrivate "Button", "FontUnderline", Abs(mbFontBtn_Underline), strSysIniTemp
    IniWriteStrPrivate "Button", "FontStrikethru", Abs(mbFontBtn_Strikethru), strSysIniTemp
    IniWriteStrPrivate "Button", "FontItalic", Abs(mbFontBtn_Italic), strSysIniTemp
    IniWriteStrPrivate "Button", "FontBold", Abs(mbFontBtn_Bold), strSysIniTemp
    IniWriteStrPrivate "Button", "FontColor", CStr(cmdFutureButton.ForeColor), strSysIniTemp
    IniWriteStrPrivate "Button", "Width", txtButtonWidth.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Height", txtButtonHeight.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Left", txtButtonLeft.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Top", txtButtonTop.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Btn2BtnLeft", txtButton2BtnL.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Btn2BtnTop", txtButton2BtnT.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "TextUpCase", chkButtonTextUpCase.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Style", cmbButtonStyle.ListIndex, strSysIniTemp
    IniWriteStrPrivate "Button", "IconStatusSkin", cmbImageStatus.Text, strSysIniTemp
    'Секция Tab
    IniWriteStrPrivate "Tab", "FontName", strFontTab_Name, strSysIniTemp
    IniWriteStrPrivate "Tab", "FontSize", miFontTab_Size, strSysIniTemp
    IniWriteStrPrivate "Tab", "FontUnderline", Abs(mbFontTab_Underline), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontStrikethru", Abs(mbFontTab_Strikethru), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontItalic", Abs(mbFontTab_Italic), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontBold", Abs(mbFontTab_Bold), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontColor", lngFontTab_Color, strSysIniTemp
    'Секция Tab2
    IniWriteStrPrivate "Tab2", "FontName", strFontTab2_Name, strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontSize", miFontTab2_Size, strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontUnderline", Abs(mbFontTab2_Underline), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontStrikethru", Abs(mbFontTab2_Strikethru), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontItalic", Abs(mbFontTab2_Italic), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontBold", Abs(mbFontTab2_Bold), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontColor", lngFontTab2_Color, strSysIniTemp
    IniWriteStrPrivate "Tab2", "StartMode", lngStartModeTab2, strSysIniTemp
    'Секция ToolTip
    IniWriteStrPrivate "ToolTip", "FontName", strFontTT_Name, strSysIniTemp
    IniWriteStrPrivate "ToolTip", "FontSize", miFontTT_Size, strSysIniTemp
    IniWriteStrPrivate "ToolTip", "FontUnderline", Abs(mbFontTT_Underline), strSysIniTemp
    IniWriteStrPrivate "ToolTip", "FontStrikethru", Abs(mbFontTT_Strikethru), strSysIniTemp
    IniWriteStrPrivate "ToolTip", "FontItalic", Abs(mbFontTT_Italic), strSysIniTemp
    IniWriteStrPrivate "ToolTip", "FontBold", Abs(mbFontTT_Bold), strSysIniTemp
    IniWriteStrPrivate "ToolTip", "FontColor", lngFontTT_Color, strSysIniTemp
    'Секция "NotebookVendor"
    IniWriteStrPrivate "NotebookVendor", "FilterCount", UBound(arrNotebookFilterList), strSysIniTemp

    For cnt = 0 To UBound(arrNotebookFilterList) - 1
        IniWriteStrPrivate "NotebookVendor", "Filter_" & cnt + 1, arrNotebookFilterList(cnt), strSysIniTemp
    Next

    ' Приводим Ini файл к читабельному виду
    NormIniFile strSysIniTemp
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Tab2CtlEnable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub Tab2CtlEnable(ByVal mbEnable As Boolean)
    chkTabHide.Enabled = mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub TabCtlEnable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub TabCtlEnable(ByVal mbEnable As Boolean)
    chkTabBlock.Enabled = mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub TempCtlEnable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub TempCtlEnable(ByVal mbEnable As Boolean)
    ucTempPath.Enabled = mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub TransferOSData
'! Description (Описание)  :   [Передача параметров ОС из спика в форму редактирования]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub TransferOSData()

    Dim i As Long

    With lvOS
        i = .SelectedItem.Index

        If i >= 0 Then

            frmOSEdit.txtOSVer.Text = .ListItems.item(i).Text
            frmOSEdit.txtOSName.Text = .ListItems.item(i).SubItems(1)
            frmOSEdit.ucPathDRP.Path = .ListItems.item(i).SubItems(2)
            frmOSEdit.ucPathDB.Path = .ListItems.item(i).SubItems(3)
            frmOSEdit.chk64bit.Value = CBool(.ListItems.item(i).SubItems(4))
    
            Select Case .ListItems.item(i).SubItems(4)
    
                Case 0
                    frmOSEdit.chk64bit.Value = False
                    frmOSEdit.chkNotCheckBitOS.Value = False
    
                Case 1
                    frmOSEdit.chk64bit.Value = True
                    frmOSEdit.chkNotCheckBitOS.Value = False
    
                Case 2
                    frmOSEdit.chk64bit.Value = False
                    frmOSEdit.chkNotCheckBitOS.Value = True
    
                Case 3
                    frmOSEdit.chk64bit.Value = True
                    frmOSEdit.chkNotCheckBitOS.Value = True
    
                Case Else
                    frmOSEdit.chk64bit.Value = False
                    frmOSEdit.chkNotCheckBitOS.Value = False
            End Select
    
            frmOSEdit.ucPhysXPath.Path = .ListItems.item(i).SubItems(5)
            frmOSEdit.ucLangPath.Path = .ListItems.item(i).SubItems(6)
            frmOSEdit.ucRuntimesPath.Path = .ListItems.item(i).SubItems(7)
            frmOSEdit.txtExcludeFileName.Text = .ListItems.item(i).SubItems(8)
            
            frmOSEdit.Show vbModal, Me
        End If
        
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub TransferUtilsData
'! Description (Описание)  :   [Передача параметров Утилит из спика в форму редактирования]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub TransferUtilsData()

    Dim i As Long

    With lvUtils
        i = .SelectedItem.Index

        If i >= 0 Then
            frmUtilsEdit.txtUtilName.Text = .ListItems.item(i).Text
            frmUtilsEdit.ucPathUtil.Path = .ListItems.item(i).SubItems(1)
            frmUtilsEdit.ucPathUtil64.Path = .ListItems.item(i).SubItems(2)
            frmUtilsEdit.txtParamUtil.Text = .ListItems.item(i).SubItems(3)
            
            frmUtilsEdit.Show vbModal, Me
        End If
        
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UpdateCtlEnable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub UpdateCtlEnable(ByVal mbEnable As Boolean)
    chkUpdateBeta.Enabled = mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkButtonDisable_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkButtonDisable_Click()
    cmdFutureButton.Enabled = CBool(chkButtonDisable.Value)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkButtonTextUpCase_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkButtonTextUpCase_Click()
    ChangeButtonProperties
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkDebug_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkDebug_Click()
    DebugCtlEnable CBool(chkDebug.Value)
    DebugCtlEnableLog2App Not CBool(chkDebugLog2AppPath.Value)

    If Not CBool(chkDebug.Value) Then
        If Not CBool(chkDebugLog2AppPath.Value) Then
            ucDebugLogPath.Enabled = False
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkDebugLog2AppPath_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkDebugLog2AppPath_Click()
    DebugCtlEnableLog2App Not CBool(chkDebugLog2AppPath.Value)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkForceIfDriverIsNotBetter_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkForceIfDriverIsNotBetter_Click()
    mbDpInstForceIfDriverIsNotBetter = CBool(chkForceIfDriverIsNotBetter.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkFormMaximaze_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkFormMaximaze_Click()

    If chkFormMaximaze.Value Then
        chkFormSizeSave.Value = vbUnchecked
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkFormSizeSave_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkFormSizeSave_Click()

    If chkFormSizeSave.Value Then
        chkFormMaximaze.Value = vbUnchecked
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkLegacyMode_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkLegacyMode_Click()
    mbDpInstLegacyMode = CBool(chkLegacyMode.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkPromptIfDriverIsNotBetter_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkPromptIfDriverIsNotBetter_Click()
    mbDpInstPromptIfDriverIsNotBetter = CBool(chkPromptIfDriverIsNotBetter.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkQuietInstall_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkQuietInstall_Click()
    mbDpInstQuietInstall = CBool(chkQuietInstall.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkScanHardware_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkScanHardware_Click()
    mbDpInstScanHardware = CBool(chkScanHardware.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkSuppressAddRemovePrograms_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkSuppressAddRemovePrograms_Click()
    mbDpInstSuppressAddRemovePrograms = CBool(chkSuppressAddRemovePrograms.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkSuppressWizard_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkSuppressWizard_Click()
    mbDpInstSuppressWizard = CBool(chkSuppressWizard.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkTabBlock_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkTabBlock_Click()
    Tab2CtlEnable CBool(chkTabBlock.Value)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkTabHide_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkTabHide_Click()
    TabCtlEnable Not CBool(chkTabHide.Value)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkTempPath_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkTempPath_Click()
    TempCtlEnable CBool(chkTempPath.Value)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkUpdate_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkUpdate_Click()
    UpdateCtlEnable CBool(chkUpdate.Value)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyle_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyle_Click()
Dim lngIndex As Long
    
    lngIndex = cmbButtonStyle.ListIndex
    If lngIndex > -1 Then
        cmdFutureButton.ButtonStyle = lngIndex
        Select Case lngIndex
            Case 0, 1, 4, 6, 7, 11, 12
                cmbButtonStyleColor.Enabled = False
                ctlStatusBtnBackColor.Visible = True
                cmbButtonStyleColor.ListIndex = 3
            Case 2, 3, 5, 8, 9, 10
                cmbButtonStyleColor.ListIndex = cmdFutureButton.ColorScheme
                cmbButtonStyleColor.Enabled = True
                cmbButtonStyleColor_Click
            Case 13
                cmbButtonStyleColor.Enabled = False
                ctlStatusBtnBackColor.Visible = False
        End Select
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyle_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyle_GotFocus()
    HighlightActiveControl Me, cmbButtonStyle, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyle_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyle_LostFocus()
    HighlightActiveControl Me, cmbButtonStyle, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyleColor_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyleColor_Click()
Dim lngIndex As Long

    lngIndex = cmbButtonStyleColor.ListIndex
    If lngIndex > -1 Then
        cmdFutureButton.ColorScheme = lngIndex
        If lngIndex < 3 Then
            ctlStatusBtnBackColor.Visible = False
        Else
            ctlStatusBtnBackColor.Visible = True
            cmdFutureButton.BackColor = ctlStatusBtnBackColor.Value
        End If
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyleColor_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyleColor_GotFocus()
    HighlightActiveControl Me, cmbButtonStyleColor, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyleColor_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyleColor_LostFocus()
    HighlightActiveControl Me, cmbButtonStyleColor, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbImageMain_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbImageMain_Click()

    If PathExists(strPathImageMain & cmbImageMain.Text) = False Then
        cmbImageMain.BackColor = vbRed
    Else
        cmbImageMain.BackColor = &H80000005
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbImageMain_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbImageMain_GotFocus()
    HighlightActiveControl Me, cmbImageMain, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbImageMain_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbImageMain_LostFocus()

    If PathExists(strPathImageMain & cmbImageMain.Text) = False Then
        cmbImageMain.BackColor = vbRed
    Else
        cmbImageMain.BackColor = &H80000005
    End If

    HighlightActiveControl Me, cmbImageMain, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbImageStatus_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbImageStatus_Click()

    Dim strPathImageStatusButtonWorkTemp As String

    If PathExists(strPathImageStatusButton & cmbImageStatus.Text) = False Then
        cmbImageStatus.BackColor = vbRed
    Else
        cmbImageStatus.BackColor = &H80000005
    End If

    strPathImageStatusButtonWorkTemp = strPathImageStatusButton & cmbImageStatus.Text
    LoadIconImage2Object imgOK, "BTN_OK", strPathImageStatusButtonWorkTemp
    Set cmdFutureButton.PictureNormal = imgOK.Picture
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbImageStatus_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbImageStatus_GotFocus()
    HighlightActiveControl Me, cmbImageStatus, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbImageStatus_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbImageStatus_LostFocus()

    If PathExists(strPathImageStatusButton & cmbImageStatus.Text) = False Then
        cmbImageStatus.BackColor = vbRed
    Else
        cmbImageStatus.BackColor = &H80000005
    End If

    HighlightActiveControl Me, cmbImageStatus, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdAddOS_Click
'! Description (Описание)  :   [кнопка добавления ОС]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdAddOS_Click()
    mbAddInList = True
    frmOSEdit.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdAddUtil_Click
'! Description (Описание)  :   [кнопка добавления утилиты]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdAddUtil_Click()
    mbAddInList = True
    frmUtilsEdit.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDelOS_Click
'! Description (Описание)  :   [кнопка удаление ОС]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdDelOS_Click()

    With lvOS

        If .ListItems.Count Then
            .ListItems.Remove (.SelectedItem.Index)
            lngLastIdOS = lngLastIdOS - 1
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDelUtil_Click
'! Description (Описание)  :   [кнопка удаление утилиты]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdDelUtil_Click()

    With lvUtils

        If .ListItems.Count Then
            .ListItems.Remove (.SelectedItem.Index)
            lngLastIdUtil = lngLastIdUtil - 1
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDriverVer_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdDriverVer_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff547394%28VS.85%29.aspx?ppud=4", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdEditOS_Click
'! Description (Описание)  :   [кнопка редактирование ОС]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdEditOS_Click()
    TransferOSData
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdEditUtil_Click
'! Description (Описание)  :   [кнопка редактирование утилиты]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdEditUtil_Click()
    TransferUtilsData
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdExit_Click
'! Description (Описание)  :   [Нажатие кнопки Выход. Выход без сохранения]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Me.Hide
    ChangeStatusTextAndDebug cmdExit.Caption
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdFontColorButton_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdFontColorButton_Click()

    With frmFontDialog
        .optControl(3).Value = True
        .txtFont.Font.Name = strFontBtn_Name
        .txtFont.Font.Size = miFontBtn_Size
        .txtFont.Font.Bold = mbFontBtn_Bold
        .txtFont.Font.Italic = mbFontBtn_Italic
        .txtFont.Font.Underline = mbFontBtn_Underline
        .txtFont.Font.Charset = lngFont_Charset
        .txtFont.ForeColor = lngFontBtn_Color
        .ctlFontColor.Value = lngFontBtn_Color
        .Show vbModal, Me
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdFontColorTabDrivers_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdFontColorTabDrivers_Click()

    With frmFontDialog
        .optControl(1).Value = True
        .txtFont.Font.Name = strFontTab2_Name
        .txtFont.Font.Size = miFontTab2_Size
        .txtFont.Font.Bold = mbFontTab2_Bold
        .txtFont.Font.Italic = mbFontTab2_Italic
        .txtFont.Font.Underline = mbFontTab2_Underline
        .txtFont.Font.Charset = lngFont_Charset
        .txtFont.ForeColor = lngFontTab2_Color
        .ctlFontColor.Value = lngFontTab2_Color
        .Show vbModal, Me
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdFontColorTabOS_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdFontColorTabOS_Click()

    With frmFontDialog
        .optControl(0).Value = True
        .txtFont.Font.Name = strFontTab_Name
        .txtFont.Font.Size = miFontTab_Size
        .txtFont.Font.Bold = mbFontTab_Bold
        .txtFont.Font.Italic = mbFontTab_Italic
        .txtFont.Font.Underline = mbFontTab_Underline
        .txtFont.Font.Charset = lngFont_Charset
        .txtFont.ForeColor = lngFontTab_Color
        .ctlFontColor.Value = lngFontTab_Color
        .Show vbModal, Me
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdFontColorToolTip_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdFontColorToolTip_Click()
    
    With frmFontDialog
        .optControl(2).Value = True
        .txtFont.Font.Name = strFontTT_Name
        .txtFont.Font.Size = miFontTT_Size
        .txtFont.Font.Bold = mbFontTT_Bold
        .txtFont.Font.Italic = mbFontTT_Italic
        .txtFont.Font.Underline = mbFontTT_Underline
        .txtFont.Font.Charset = lngFont_Charset
        .txtFont.ForeColor = lngFontTT_Color
        .ctlFontColor.Value = lngFontTT_Color
        .Show vbModal, Me
    End With
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdForceIfDriverIsNotBetter_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdForceIfDriverIsNotBetter_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff544948.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdLegacyMode_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdLegacyMode_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff548635.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdOK_Click
'! Description (Описание)  :   [Нажатие кнопки ОК. Применение настроек]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()

    Dim lngMsgRet As Long

    If mbIsDriveCDRoom And mbLoadIniTmpAfterRestart Then
        SaveOptions
        ChangeStatusTextAndDebug strMessages(36)
        lngMsgRet = MsgBox(strMessages(36) & strMessages(147), vbInformation + vbApplicationModal + vbYesNo, strProductName)
        mbRestartProgram = lngMsgRet = vbYes
    ElseIf FileExists(strSysIni) Then
        If Not FileisReadOnly(strSysIni) Then
            SaveOptions
            ChangeStatusTextAndDebug strMessages(36)
            lngMsgRet = MsgBox(strMessages(36) & strMessages(147), vbInformation + vbApplicationModal + vbYesNo, strProductName)
            mbRestartProgram = lngMsgRet = vbYes
        End If
    Else
        SaveOptions
        ChangeStatusTextAndDebug strMessages(36)
        lngMsgRet = MsgBox(strMessages(36) & strMessages(147), vbInformation + vbApplicationModal + vbYesNo, strProductName)
        mbRestartProgram = lngMsgRet = vbYes
    End If

    Me.Hide
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdPathDefault_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdPathDefault_Click()
    ucDevCon86Path.Path = "Tools\Devcon\devcon.exe"
    ucDevCon64Path.Path = "Tools\Devcon\devcon64.exe"
    ucDevCon86Pathw2k.Path = "Tools\Devcon\devconw2k.exe"
    'Секция DPInst
    ucDPInst86Path.Path = "Tools\DPInst\DPInst.exe"
    ucDPInst64Path.Path = "Tools\DPInst\DPInst64.exe"
    'Секция Arc
    ucArchPath.Path = "Tools\Arc\7za.exe"
    ucCmdDevconPath.Path = "Tools\Devcon\devcon_c.cmd"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdPromptIfDriverIsNotBetter_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdPromptIfDriverIsNotBetter_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff549759.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdQuietInstall_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdQuietInstall_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff549799.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdScanHardware_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdScanHardware_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff550761.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdSuppressAddRemovePrograms_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdSuppressAddRemovePrograms_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff553404.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdSuppressWizard_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdSuppressWizard_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff550803.aspx#setting_the_suppresswizard_flag", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ctlStatusBtnBackColor_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ctlStatusBtnBackColor_Click()
    If cmbButtonStyleColor.ListIndex = 3 Then
        cmdFutureButton.BackColor = ctlStatusBtnBackColor.Value
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_KeyDown
'! Description (Описание)  :   [Обработка нажатий клавиш клавиатуры сначала на форме]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        If MsgBox(strMessages(37), vbQuestion + vbYesNo, strProductName) = vbYes Then
            cmdExit_Click
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Load
'! Description (Описание)  :   [Загрузка формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmOptions", False
        .Height = 5850
        .Width = 11900
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    'Top frame position
    frOptions.Top = 50
    frDesign.Top = 50
    frDesign2.Top = 50
    frDpInstParam.Top = 50
    frMain.Top = 50
    frMain2.Top = 50
    frMainTools.Top = 50
    frOtherTools.Top = 50
    frOS.Top = 50
    frOther.Top = 50
    frDebug.Top = 50
    'Left frame position
    frDesign.Left = 3100
    frDesign2.Left = 3100
    frDpInstParam.Left = 3100
    frMain.Left = 3100
    frMain2.Left = 3100
    frMainTools.Left = 3100
    frOS.Left = 3100
    frOther.Left = 3100
    frOtherTools.Left = 3100
    frDebug.Left = 3100
    ' Устанавливаем минимальные значения для текстовых полей
    txtTabPerRowCount.Min = 2
    txtFormHeight.Min = lngMainFormHeightMin
    txtFormWidth.Min = lngMainFormWidthMin
    txtButtonHeight.Min = lngButtonHeightMin
    txtButtonWidth.Min = lngButtonWidthMin
    ' Устанавливаем картинки кнопок и убираем описание кнопок
    LoadIconImage2Object cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2Object cmdExit, "BTN_EXIT", strPathImageMainWork
    LoadIconImage2Object cmdAddUtil, "BTN_ADD", strPathImageMainWork
    LoadIconImage2Object cmdEditUtil, "BTN_EDIT", strPathImageMainWork
    LoadIconImage2Object cmdDelUtil, "BTN_DELETE", strPathImageMainWork
    LoadIconImage2Object cmdAddOS, "BTN_ADD", strPathImageMainWork
    LoadIconImage2Object cmdEditOS, "BTN_EDIT", strPathImageMainWork
    LoadIconImage2Object cmdDelOS, "BTN_DELETE", strPathImageMainWork
    LoadIconImage2Object cmdFontColorButton, "BTN_FONT", strPathImageMainWork
    LoadIconImage2Object cmdFontColorTabOS, "BTN_FONT", strPathImageMainWork
    LoadIconImage2Object cmdFontColorTabDrivers, "BTN_FONT", strPathImageMainWork
    LoadIconImage2Object cmdFontColorToolTip, "BTN_FONT", strPathImageMainWork
    'загружаем список стилей кнопок
    LoadComboBtnStyle
    ' Действия при загрузке формы
    FormLoadAction
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_QueryUnload
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Cancel (Integer)
'                              UnloadMode (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' Выгружаем из памяти форму и другие компоненты
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
        ChangeStatusTextAndDebug cmdExit.Caption
    Else
        Set frmOptions = Nothing
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Resize
'! Description (Описание)  :   [Изменение размеров формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Resize()

    On Error Resume Next

    If Me.WindowState <> vbMinimized Then
        SetTrayIcon NIM_DELETE, Me.hWnd, 0&, vbNullString
    Else
        SetTrayIcon NIM_ADD, Me.hWnd, Me.Icon, "Drivers Installer Assistant"
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadList_lvOptions
'! Description (Описание)  :   [Построение дерева настроек]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadList_lvOptions()
    
    ' Загружаем картинки в ImageList
    With ImageListOptions.ListImages
        If .Count = 0 Then
            .Add 1, , LoadIconImageFromPath("OPT_MAIN", strPathImageMainWork)
            .Add 2, , LoadIconImageFromPath("OPT_MAIN2", strPathImageMainWork)
            .Add 3, , LoadIconImageFromPath("OPT_OSLIST", strPathImageMainWork)
            .Add 4, , LoadIconImageFromPath("OPT_TOOLS_MAIN", strPathImageMainWork)
            .Add 5, , LoadIconImageFromPath("OPT_TOOLS_OTHER", strPathImageMainWork)
            .Add 6, , LoadIconImageFromPath("OPT_DESIGN", strPathImageMainWork)
            .Add 7, , LoadIconImageFromPath("OPT_DESIGN2", strPathImageMainWork)
            .Add 8, , LoadIconImageFromPath("OPT_DPINST", strPathImageMainWork)
            .Add 9, , LoadIconImageFromPath("OPT_DEVPARSER", strPathImageMainWork)
        End If
    End With
        
    ' Заполняем ListView названием опций программы
    With lvOptions
        With .ListItems
            If .Count = 0 Then
                .Add 1, , strItemOptions1, , 1
                .Add 2, , strItemOptions8, , 2
                .Add 3, , strItemOptions2, , 3
                .Add 4, , strItemOptions3, , 4
                .Add 5, , strItemOptions4, , 5
                .Add 6, , strItemOptions5, , 6
                .Add 7, , strItemOptions9, , 7
                .Add 8, , strItemOptions6, , 8
                .Add 9, , strItemOptions10, , 9
            Else
                .item(1).Text = strItemOptions1
                .item(2).Text = strItemOptions8
                .item(3).Text = strItemOptions2
                .item(4).Text = strItemOptions3
                .item(5).Text = strItemOptions4
                .item(6).Text = strItemOptions5
                .item(7).Text = strItemOptions9
                .item(8).Text = strItemOptions6
                .item(9).Text = strItemOptions10
            End If
        End With
    
        .ColumnWidth = .Width - 100
    End With
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadList_lvOS
'! Description (Описание)  :   [Построение спиcка ОС]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadList_lvOS()

    Dim i As Long

    With lvOS
        .ListItems.Clear

        If .ColumnHeaders.Count = 0 Then
            .ColumnHeaders.Add 1, , strTableOSHeader1, 80 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 2, , strTableOSHeader2, 150 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 3, , strTableOSHeader3, 150 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 4, , strTableOSHeader4, 120 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 5, , strTableOSHeader5, 30 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 6, , strTableOSHeader6, 50 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 7, , strTableOSHeader7, 50 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 8, , strTableOSHeader8, 50 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 9, , strTableOSHeader9, 120 * Screen.TwipsPerPixelX
        End If

        For i = 0 To lngOSCount - 1

            With .ListItems.Add(, , arrOSList(i).Ver)
                .SubItems(1) = arrOSList(i).Name
                .SubItems(2) = arrOSList(i).drpFolder
                .SubItems(3) = arrOSList(i).devIDFolder
                .SubItems(4) = arrOSList(i).is64bit
                .SubItems(5) = arrOSList(i).PathPhysX
                .SubItems(6) = arrOSList(i).PathLanguages
                .SubItems(7) = arrOSList(i).PathRuntimes
                .SubItems(8) = arrOSList(i).ExcludeFileName
            End With

        Next

    End With

    lngLastIdOS = lngOSCount
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadList_lvUtils
'! Description (Описание)  :   [Построение спика утилит]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadList_lvUtils()

    Dim i As Long

    With lvUtils
        .ListItems.Clear

        If .ColumnHeaders.Count = 0 Then
            .ColumnHeaders.Add 1, , strTableUtilHeader1, 200 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 2, , strTableUtilHeader2, 200 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 3, , strTableUtilHeader4, 200 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 4, , strTableUtilHeader3, 120 * Screen.TwipsPerPixelX
        End If

        For i = 0 To lngUtilsCount - 1

            With .ListItems.Add(, , arrUtilsList(i, 0))
                .SubItems(1) = arrUtilsList(i, 1)
                .SubItems(2) = arrUtilsList(i, 2)
                .SubItems(3) = arrUtilsList(i, 3)
            End With

        Next

    End With

    lngLastIdUtil = lngUtilsCount
End Sub

''!--------------------------------------------------------------------------------
''! Procedure   (Функция)   :   Sub lvOptions_ItemChanged
''! Description (Описание)  :   [При выборе опции происходит отображение соответсвующего окна]
''! Parameters  (Переменные):   iIndex (Long)
''!--------------------------------------------------------------------------------
Private Sub lvOptions_ItemSelect(ByVal item As LvwListItem, ByVal Selected As Boolean)

    If Selected Then
        Select Case item.Index
    
            Case 1
            'ItemOptions1=Основные настройки
                frMain.ZOrder 0
    
            Case 2
            ' ItemOptions8=Основные настройки 2
                frMain2.ZOrder 0
    
            Case 3
            'ItemOptions2=Поддерживаемые ОС
                frOS.ZOrder 0
    
            Case 4
            'ItemOptions3=Рабочие утилиты
                frMainTools.ZOrder 0
    
            Case 5
            'ItemOptions4=Вспомогательные утилиты
                frOtherTools.ZOrder 0
    
            Case 6
            'ItemOptions5=Оформление программы
                frDesign.ZOrder 0
    
            Case 7
            'ItemOptions9=Оформление программы 2
                frDesign2.ZOrder 0
    
            Case 8
            'ItemOptions6=Параметры запуска DPInst
                frDpInstParam.ZOrder 0
    
            Case 9
            'ItemOptions10=Отладочный режим
                frDebug.ZOrder 0
    
            Case Else
                frOther.ZOrder 0
        End Select
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvOS_ColumnClick
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   ColumnHeader (LvwColumnHeader)
'!--------------------------------------------------------------------------------
Private Sub lvOS_ColumnClick(ByVal ColumnHeader As LvwColumnHeader)

    Dim i As Long

    With lvOS
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvOS_ItemDblClick
'! Description (Описание)  :   [Двойнок клик по элементу списка вызывает форму редактирования]
'! Parameters  (Переменные):   Item (LvwListItem)
'                              Button (Integer)
'!--------------------------------------------------------------------------------
Private Sub lvOS_ItemDblClick(ByVal item As LvwListItem, ByVal Button As Integer)
    TransferOSData
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvUtils_ColumnClick
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   ColumnHeader (LvwColumnHeader)
'!--------------------------------------------------------------------------------
Private Sub lvUtils_ColumnClick(ByVal ColumnHeader As LvwColumnHeader)

    Dim i As Long

    With lvUtils
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvUtils_ItemDblClick
'! Description (Описание)  :   [Двойнок клик по элементу списка вызывает форму редактирования]
'! Parameters  (Переменные):   Item (LvwListItem)
'                              Button (Integer)
'!--------------------------------------------------------------------------------
Private Sub lvUtils_ItemDblClick(ByVal item As LvwListItem, ByVal Button As Integer)
    TransferUtilsData
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtButtonHeight_Change
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtButtonHeight_Change()
    ChangeButtonProperties
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtButtonWidth_Change
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtButtonWidth_Change()
    ChangeButtonProperties
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtCmdStringDPInst_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtCmdStringDPInst_GotFocus()
    HighlightActiveControl Me, txtCmdStringDPInst, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtCmdStringDPInst_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtCmdStringDPInst_LostFocus()
    HighlightActiveControl Me, txtCmdStringDPInst, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtDebugLogName_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtDebugLogName_GotFocus()
    HighlightActiveControl Me, txtDebugLogName, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtDebugLogName_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtDebugLogName_LostFocus()
    HighlightActiveControl Me, txtDebugLogName, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtExcludeHWID_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtExcludeHWID_GotFocus()
    HighlightActiveControl Me, txtExcludeHWID, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtExcludeHWID_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtExcludeHWID_LostFocus()
    HighlightActiveControl Me, txtExcludeHWID, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPath_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArchPath_Click()

    Dim strTempPath As String

    If ucArchPath.FileCount Then
        strTempPath = ucArchPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucArchPath.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPath_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArchPath_GotFocus()
    HighlightActiveControl Me, ucArchPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPath_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArchPath_LostFocus()
    HighlightActiveControl Me, ucArchPath, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucCmdDevconPath_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucCmdDevconPath_Click()

    Dim strTempPath As String

    If ucCmdDevconPath.FileCount Then
        strTempPath = ucCmdDevconPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucCmdDevconPath.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucCmdDevconPath_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucCmdDevconPath_GotFocus()
    HighlightActiveControl Me, ucCmdDevconPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucCmdDevconPath_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucCmdDevconPath_LostFocus()
    HighlightActiveControl Me, ucCmdDevconPath, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDebugLogPath_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDebugLogPath_Click()

    Dim strTempPath As String

    If ucDebugLogPath.FileCount Then
        strTempPath = ucDebugLogPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucDebugLogPath.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDebugLogPath_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDebugLogPath_GotFocus()
    HighlightActiveControl Me, ucDebugLogPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDebugLogPath_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDebugLogPath_LostFocus()
    HighlightActiveControl Me, ucDebugLogPath, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDevCon64Path_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon64Path_Click()

    Dim strTempPath As String

    If ucDevCon64Path.FileCount Then
        strTempPath = ucDevCon64Path.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucDevCon64Path.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDevCon64Path_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon64Path_GotFocus()
    HighlightActiveControl Me, ucDevCon64Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDevCon64Path_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon64Path_LostFocus()
    HighlightActiveControl Me, ucDevCon64Path, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDevCon86Path_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon86Path_Click()

    Dim strTempPath As String

    If ucDevCon86Path.FileCount Then
        strTempPath = ucDevCon86Path.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucDevCon86Path.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDevCon86Path_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon86Path_GotFocus()
    HighlightActiveControl Me, ucDevCon86Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDevCon86Path_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon86Path_LostFocus()
    HighlightActiveControl Me, ucDevCon86Path, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDevCon86Pathw2k_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon86Pathw2k_Click()

    Dim strTempPath As String

    If ucDevCon86Pathw2k.FileCount Then
        strTempPath = ucDevCon86Pathw2k.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucDevCon86Pathw2k.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDevCon86Pathw2k_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon86Pathw2k_GotFocus()
    HighlightActiveControl Me, ucDevCon86Pathw2k, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDevCon86Pathw2k_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon86Pathw2k_LostFocus()
    HighlightActiveControl Me, ucDevCon86Pathw2k, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst64Path_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst64Path_Click()

    Dim strTempPath As String

    If ucDPInst64Path.FileCount Then
        strTempPath = ucDPInst64Path.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucDPInst64Path.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst64Path_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst64Path_GotFocus()
    HighlightActiveControl Me, ucDPInst64Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst64Path_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst64Path_LostFocus()
    HighlightActiveControl Me, ucDPInst64Path, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst86Path_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst86Path_Click()

    Dim strTempPath As String

    If ucDPInst86Path.FileCount Then
        strTempPath = ucDPInst86Path.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucDPInst86Path.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst86Path_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst86Path_GotFocus()
    HighlightActiveControl Me, ucDPInst86Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst86Path_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst86Path_LostFocus()
    HighlightActiveControl Me, ucDPInst86Path, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucTempPath_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucTempPath_Click()

    Dim strTempPath As String

    If ucTempPath.FileCount Then
        strTempPath = ucTempPath.Path

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) Then
        ucTempPath.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucTempPath_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucTempPath_GotFocus()
    HighlightActiveControl Me, ucTempPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucTempPath_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucTempPath_LostFocus()
    HighlightActiveControl Me, ucTempPath, False
End Sub

