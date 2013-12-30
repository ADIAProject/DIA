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
   Begin prjDIADBS.ctlJCFrames frOptions 
      Height          =   5300
      Left            =   50
      Top             =   25
      Width           =   3000
      _extentx        =   5292
      _extenty        =   9340
      font            =   "frmOptions.frx":058A
      fillcolor       =   14016736
      roundedcorner   =   0   'False
      caption         =   "Настройки"
      textboxheight   =   18
      themecolor      =   2
      headerstyle     =   1
      backcolor       =   14016736
      Begin prjDIADBS.ctlJCbutton cmdOK 
         Height          =   750
         Left            =   75
         TabIndex        =   0
         Top             =   3500
         Width           =   2850
         _extentx        =   5027
         _extenty        =   1323
         font            =   "frmOptions.frx":05B2
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Сохранить изменения и выйти"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdExit 
         Height          =   735
         Left            =   75
         TabIndex        =   18
         Top             =   4400
         Width           =   2850
         _extentx        =   5027
         _extenty        =   1296
         font            =   "frmOptions.frx":05DA
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Выход без сохранения"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
   End
   Begin prjDIADBS.ctlJCFrames frMain 
      Height          =   5300
      Left            =   3105
      Top             =   25
      Width           =   8655
      _extentx        =   15266
      _extenty        =   9340
      font            =   "frmOptions.frx":0602
      fillcolor       =   14016736
      roundedcorner   =   0   'False
      caption         =   "Основные настройки программы"
      textboxheight   =   18
      themecolor      =   2
      headerstyle     =   1
      backcolor       =   14016736
      Begin prjDIADBS.CheckBoxW chkRemoveTemp 
         Height          =   210
         Left            =   435
         TabIndex        =   23
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
         Caption         =   "frmOptions.frx":062A
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkUpdate 
         Height          =   210
         Left            =   435
         TabIndex        =   24
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
         Caption         =   "frmOptions.frx":06A2
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkReadDPName 
         Height          =   210
         Left            =   435
         TabIndex        =   32
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
         Caption         =   "frmOptions.frx":06FE
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkConvertDPName 
         Height          =   210
         Left            =   435
         TabIndex        =   33
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
         Caption         =   "frmOptions.frx":0772
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkHideOtherProcess 
         Height          =   210
         Left            =   435
         TabIndex        =   54
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
         Caption         =   "frmOptions.frx":0842
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkTempPath 
         Height          =   210
         Left            =   435
         TabIndex        =   55
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
         Caption         =   "frmOptions.frx":08A8
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkUpdateBeta 
         Height          =   210
         Left            =   3780
         TabIndex        =   58
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
         Caption         =   "frmOptions.frx":08F8
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSilentDll 
         Height          =   210
         Left            =   435
         TabIndex        =   19
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
         Caption         =   "frmOptions.frx":096E
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSearchOnStart 
         Height          =   210
         Left            =   435
         TabIndex        =   20
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
         Caption         =   "frmOptions.frx":0A0A
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.SpinBox txtPauseAfterSearch 
         Height          =   255
         Left            =   7710
         TabIndex        =   21
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
         TabIndex        =   1
         Top             =   3200
         Width           =   4575
         _extentx        =   8070
         _extenty        =   556
         useautoforecolor=   0   'False
         defaultext      =   ""
         enabled         =   0   'False
         filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.ctlJCbutton optRezim_Intellect 
         Height          =   510
         Left            =   420
         TabIndex        =   59
         Top             =   4300
         Width           =   2505
         _extentx        =   4419
         _extenty        =   900
         font            =   "frmOptions.frx":0A88
         buttonstyle     =   8
         backcolor       =   14935011
         caption         =   "Установка (Совместимые драйвера)"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         mode            =   2
         picturealign    =   0
         tooltipbackcolor=   0
         colorscheme     =   3
      End
      Begin prjDIADBS.ctlJCbutton optRezim_Upd 
         Height          =   510
         Left            =   5700
         TabIndex        =   60
         Top             =   4300
         Width           =   2505
         _extentx        =   4419
         _extenty        =   900
         font            =   "frmOptions.frx":0AB0
         buttonstyle     =   8
         backcolor       =   14935011
         caption         =   "Создание или обновление базы драйверов"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         mode            =   2
         picturealign    =   0
         tooltipbackcolor=   0
         colorscheme     =   3
      End
      Begin prjDIADBS.ctlJCbutton optRezim_Ust 
         Height          =   510
         Left            =   3060
         TabIndex        =   61
         Top             =   4300
         Width           =   2505
         _extentx        =   4419
         _extenty        =   900
         font            =   "frmOptions.frx":0AD8
         buttonstyle     =   8
         backcolor       =   14935011
         caption         =   "Установка (Полная - весь пакет)"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         mode            =   2
         picturealign    =   0
         tooltipbackcolor=   0
         colorscheme     =   3
      End
      Begin prjDIADBS.LabelW lblPauseAfterSearch 
         Height          =   225
         Left            =   5400
         TabIndex        =   91
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
         TabIndex        =   92
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
         TabIndex        =   93
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
         TabIndex        =   94
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
      _extentx        =   15266
      _extenty        =   9340
      font            =   "frmOptions.frx":0B00
      fillcolor       =   14016736
      roundedcorner   =   0   'False
      caption         =   "Основные настройки программы 2"
      textboxheight   =   18
      themecolor      =   2
      headerstyle     =   1
      backcolor       =   14016736
      Begin VB.CommandButton cmdDriverVer 
         Caption         =   "?"
         Height          =   255
         Left            =   300
         TabIndex        =   88
         Top             =   1550
         Width           =   255
      End
      Begin prjDIADBS.OptionButtonW optCompareByVersion 
         Height          =   255
         Left            =   300
         TabIndex        =   63
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
         Caption         =   "frmOptions.frx":0B28
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optCompareByDate 
         Height          =   255
         Left            =   300
         TabIndex        =   62
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
         Caption         =   "frmOptions.frx":0BAA
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtCompareVersionDRV 
         Height          =   1005
         Left            =   300
         TabIndex        =   22
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
         Text            =   "frmOptions.frx":0C58
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         CueBanner       =   "frmOptions.frx":0E4C
      End
      Begin prjDIADBS.CheckBoxW chkDateFormatRus 
         Height          =   210
         Left            =   300
         TabIndex        =   2
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
         Caption         =   "frmOptions.frx":0E6C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkCreateRP 
         Height          =   210
         Left            =   300
         TabIndex        =   25
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
         Caption         =   "frmOptions.frx":0EE6
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkCompatiblesHWID 
         Height          =   210
         Left            =   300
         TabIndex        =   26
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
         Caption         =   "frmOptions.frx":0F6E
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblCompareVersionDRV 
         Height          =   225
         Left            =   600
         TabIndex        =   95
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
      _extentx        =   15266
      _extenty        =   9340
      font            =   "frmOptions.frx":0FE8
      fillcolor       =   14016736
      roundedcorner   =   0   'False
      caption         =   "Расположение основных утилит (Tools)"
      textboxheight   =   18
      themecolor      =   2
      headerstyle     =   1
      backcolor       =   14016736
      Begin prjDIADBS.ctlUcPickBox ucDevCon86Path 
         Height          =   315
         Left            =   2520
         TabIndex        =   27
         Top             =   450
         Width           =   6000
         _extentx        =   10583
         _extenty        =   556
         useautoforecolor=   0   'False
         defaultext      =   ""
         dialogtype      =   1
         enabled         =   0   'False
         fileflags       =   2621446
         filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucDevCon64Path 
         Height          =   315
         Left            =   2520
         TabIndex        =   29
         Top             =   850
         Width           =   6000
         _extentx        =   10583
         _extenty        =   556
         defaultext      =   ""
         dialogtype      =   1
         enabled         =   0   'False
         fileflags       =   2621446
         filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucDevCon86Pathw2k 
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   1250
         Width           =   6000
         _extentx        =   10583
         _extenty        =   556
         defaultext      =   ""
         dialogtype      =   1
         enabled         =   0   'False
         fileflags       =   2621446
         filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucDPInst86Path 
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   1650
         Width           =   6000
         _extentx        =   10583
         _extenty        =   556
         defaultext      =   ""
         dialogtype      =   1
         enabled         =   0   'False
         fileflags       =   2621446
         filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucDPInst64Path 
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   2050
         Width           =   6000
         _extentx        =   10583
         _extenty        =   556
         defaultext      =   ""
         dialogtype      =   1
         enabled         =   0   'False
         fileflags       =   2621446
         filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucArchPath 
         Height          =   315
         Left            =   2520
         TabIndex        =   6
         Top             =   2450
         Width           =   6000
         _extentx        =   10583
         _extenty        =   556
         defaultext      =   ""
         dialogtype      =   1
         enabled         =   0   'False
         fileflags       =   2621446
         filters         =   "Supported files|*.exe|EXE Files (*.exe)"
      End
      Begin prjDIADBS.ctlUcPickBox ucCmdDevconPath 
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Top             =   2850
         Width           =   6000
         _extentx        =   10583
         _extenty        =   556
         defaultext      =   ""
         dialogtype      =   1
         enabled         =   0   'False
         fileflags       =   2621446
         filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.ctlJCbutton cmdPathDefault 
         Height          =   495
         Left            =   4900
         TabIndex        =   64
         Top             =   3300
         Width           =   3615
         _extentx        =   6376
         _extenty        =   873
         font            =   "frmOptions.frx":1010
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Сбросить настройки расположения утилит"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
      Begin prjDIADBS.LabelW lblDevCon64 
         Height          =   315
         Left            =   100
         TabIndex        =   96
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
         TabIndex        =   97
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
         TabIndex        =   98
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
         TabIndex        =   99
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
         TabIndex        =   100
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
         TabIndex        =   101
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
         TabIndex        =   102
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
      _extentx        =   15266
      _extenty        =   9340
      font            =   "frmOptions.frx":1038
      fillcolor       =   14016736
      roundedcorner   =   0   'False
      caption         =   "Вспомогательные утилиты (Отображаются в меню ""Утилиты"")"
      textboxheight   =   18
      themecolor      =   2
      headerstyle     =   1
      backcolor       =   14016736
      Begin prjDIADBS.ListView lvUtils 
         Height          =   3855
         Left            =   120
         TabIndex        =   89
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
         Icons           =   "frmOptions.frx":1060
         SmallIcons      =   "frmOptions.frx":108C
         ColumnHeaderIcons=   "frmOptions.frx":10B8
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
         Height          =   750
         Left            =   120
         TabIndex        =   68
         Top             =   4440
         Width           =   1815
         _extentx        =   3201
         _extenty        =   1323
         font            =   "frmOptions.frx":10E4
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Добавить"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdEditUtil 
         Height          =   750
         Left            =   2160
         TabIndex        =   69
         Top             =   4455
         Width           =   1815
         _extentx        =   3201
         _extenty        =   1323
         font            =   "frmOptions.frx":110C
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Изменить"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdDelUtil 
         Height          =   750
         Left            =   4200
         TabIndex        =   70
         Top             =   4455
         Width           =   1815
         _extentx        =   3201
         _extenty        =   1323
         font            =   "frmOptions.frx":1134
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Удалить"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
   End
   Begin prjDIADBS.ctlJCFrames frOS 
      Height          =   5295
      Left            =   3885
      Top             =   1245
      Width           =   8655
      _extentx        =   15266
      _extenty        =   9340
      font            =   "frmOptions.frx":115C
      fillcolor       =   14016736
      roundedcorner   =   0   'False
      caption         =   "Поддерживаемые ОС"
      textboxheight   =   18
      themecolor      =   2
      headerstyle     =   1
      backcolor       =   14016736
      Begin prjDIADBS.ListView lvOS 
         Height          =   2350
         Left            =   120
         TabIndex        =   90
         Top             =   480
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   4154
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Icons           =   "frmOptions.frx":1184
         SmallIcons      =   "frmOptions.frx":11B0
         ColumnHeaderIcons=   "frmOptions.frx":11DC
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
         TabIndex        =   30
         Top             =   3180
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
         Text            =   "frmOptions.frx":1208
         MultiLine       =   -1  'True
         ScrollBars      =   2
         CueBanner       =   "frmOptions.frx":1228
      End
      Begin prjDIADBS.CheckBoxW chkLoadFinishFile 
         Height          =   345
         Left            =   135
         TabIndex        =   56
         Top             =   3990
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
         Caption         =   "frmOptions.frx":1248
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkRecursion 
         Height          =   255
         Left            =   135
         TabIndex        =   57
         Top             =   3735
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
         Caption         =   "frmOptions.frx":1314
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdAddOS 
         Height          =   750
         Left            =   120
         TabIndex        =   65
         Top             =   4440
         Width           =   1815
         _extentx        =   3201
         _extenty        =   1323
         font            =   "frmOptions.frx":13AC
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Добавить"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdEditOS 
         Height          =   750
         Left            =   2160
         TabIndex        =   66
         Top             =   4455
         Width           =   1815
         _extentx        =   3201
         _extenty        =   1323
         font            =   "frmOptions.frx":13D4
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Изменить"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdDelOS 
         Height          =   750
         Left            =   4200
         TabIndex        =   67
         Top             =   4455
         Width           =   1815
         _extentx        =   3201
         _extenty        =   1323
         font            =   "frmOptions.frx":13FC
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Удалить"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
      Begin prjDIADBS.LabelW lblExcludeHWID 
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   2900
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
      _extentx        =   15266
      _extenty        =   9340
      font            =   "frmOptions.frx":1424
      fillcolor       =   14016736
      roundedcorner   =   0   'False
      caption         =   "Оформление"
      textboxheight   =   18
      themecolor      =   2
      headerstyle     =   1
      backcolor       =   14016736
      Begin VB.PictureBox imgOK 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0FFFF&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   8040
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   4680
         Visible         =   0   'False
         Width           =   495
      End
      Begin prjDIADBS.CheckBoxW chkFutureButton 
         Height          =   210
         Left            =   780
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   3240
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
         Caption         =   "frmOptions.frx":144C
         Transparent     =   -1  'True
      End
      Begin VB.ComboBox cmbImageMain 
         Height          =   315
         ItemData        =   "frmOptions.frx":146C
         Left            =   615
         List            =   "frmOptions.frx":1479
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   4845
         Width           =   3000
      End
      Begin VB.ComboBox cmbImageStatus 
         Height          =   315
         ItemData        =   "frmOptions.frx":14B1
         Left            =   3960
         List            =   "frmOptions.frx":14BE
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   4845
         Width           =   3000
      End
      Begin prjDIADBS.CheckBoxW chkButtonTextUpCase 
         Height          =   210
         Left            =   3510
         TabIndex        =   31
         Top             =   2370
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
         Caption         =   "frmOptions.frx":14F6
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkFormMaximaze 
         Height          =   210
         Left            =   3495
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   660
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
         Caption         =   "frmOptions.frx":156A
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.SpinBox txtButtonHeight 
         Height          =   255
         Left            =   1485
         TabIndex        =   10
         Top             =   1605
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
         TabIndex        =   11
         Top             =   1965
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
         TabIndex        =   7
         Top             =   1605
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
         TabIndex        =   12
         Top             =   1965
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
         TabIndex        =   13
         Top             =   2310
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
         TabIndex        =   14
         Top             =   2655
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
      Begin prjDIADBS.SpinBox txtFormHeight 
         Height          =   255
         Left            =   1455
         TabIndex        =   15
         Top             =   660
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
         Left            =   1455
         TabIndex        =   16
         Top             =   1005
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
         Left            =   3495
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1005
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
         Caption         =   "frmOptions.frx":15D0
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorButton 
         Height          =   795
         Left            =   3480
         TabIndex        =   73
         Top             =   2805
         Width           =   2450
         _extentx        =   5027
         _extenty        =   1402
         font            =   "frmOptions.frx":1634
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Установить цвет и шрифт текста кнопки"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
      Begin prjDIADBS.CheckBoxW chkButtonDisable 
         Height          =   450
         Left            =   780
         TabIndex        =   74
         Top             =   3660
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":165C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlXpButton cmdFutureButton 
         Height          =   615
         Left            =   660
         TabIndex        =   75
         Top             =   3000
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Кнопка пакета драйверов"
         ButtonStyle     =   3
         PictureWidth    =   32
         PictureHeight   =   32
         ShowFocusRect   =   0   'False
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
         TextColor       =   0
         MenuExist       =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorToolTip 
         Height          =   795
         Left            =   6060
         TabIndex        =   141
         Top             =   2805
         Width           =   2445
         _extentx        =   3969
         _extenty        =   1402
         font            =   "frmOptions.frx":16D2
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Установить цвет и шрифт текста подсказок"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
      Begin prjDIADBS.ToolTip TT 
         Left            =   7320
         Top             =   4680
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
         Title           =   "frmOptions.frx":16FA
      End
      Begin prjDIADBS.LabelW lblTheme 
         Height          =   225
         Left            =   360
         TabIndex        =   104
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
         TabIndex        =   105
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
         TabIndex        =   106
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
         TabIndex        =   107
         Top             =   1965
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
         TabIndex        =   108
         Top             =   1605
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
         TabIndex        =   109
         Top             =   2655
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
         TabIndex        =   110
         Top             =   2310
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
         TabIndex        =   111
         Top             =   1965
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
         TabIndex        =   112
         Top             =   1605
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
         TabIndex        =   113
         Top             =   1305
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
      Begin prjDIADBS.LabelW lblFormWidth 
         Height          =   210
         Left            =   615
         TabIndex        =   114
         Top             =   1005
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
      Begin prjDIADBS.LabelW lblFormHeight 
         Height          =   210
         Left            =   615
         TabIndex        =   115
         Top             =   660
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
      Begin prjDIADBS.LabelW lblSizeForm 
         Height          =   210
         Left            =   360
         TabIndex        =   116
         Top             =   360
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
   End
   Begin prjDIADBS.ctlJCFrames frDesign2 
      Height          =   5295
      Left            =   4260
      Top             =   1845
      Width           =   8655
      _extentx        =   15266
      _extenty        =   9340
      font            =   "frmOptions.frx":171A
      fillcolor       =   14016736
      roundedcorner   =   0   'False
      caption         =   "Оформление 2"
      textboxheight   =   18
      themecolor      =   2
      headerstyle     =   1
      backcolor       =   14016736
      Begin prjDIADBS.SpinBox txtTabPerRowCount 
         Height          =   255
         Left            =   3330
         TabIndex        =   28
         Top             =   795
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
         TabIndex        =   35
         Top             =   1125
         Width           =   8000
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
         Caption         =   "frmOptions.frx":1742
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkTabHide 
         Height          =   210
         Left            =   390
         TabIndex        =   36
         Top             =   1440
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
         Caption         =   "frmOptions.frx":1804
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkLoadUnSupportedOS 
         Height          =   210
         Left            =   390
         TabIndex        =   53
         Top             =   1755
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
         Caption         =   "frmOptions.frx":18B2
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorTabOS 
         Height          =   795
         Left            =   390
         TabIndex        =   71
         Top             =   2070
         Width           =   2850
         _extentx        =   5027
         _extenty        =   1402
         font            =   "frmOptions.frx":192E
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Установить цвет и шрифт текста закладки"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorTabDrivers 
         Height          =   795
         Left            =   390
         TabIndex        =   72
         Top             =   3360
         Width           =   2850
         _extentx        =   5027
         _extenty        =   1402
         font            =   "frmOptions.frx":1956
         buttonstyle     =   8
         backcolor       =   16765357
         caption         =   "Установить цвет и шрифт текста закладки"
         pictureshadow   =   -1  'True
         picturepushonhover=   -1  'True
         captioneffects  =   0
         picturealign    =   0
         tooltipbackcolor=   0
      End
      Begin prjDIADBS.LabelW lblTabPerRowCount 
         Height          =   225
         Left            =   390
         TabIndex        =   117
         Top             =   795
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
         TabIndex        =   118
         Top             =   480
         Width           =   8200
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
         TabIndex        =   119
         Top             =   3000
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
      _extentx        =   15266
      _extenty        =   9340
      font            =   "frmOptions.frx":197E
      fillcolor       =   14016736
      roundedcorner   =   0   'False
      caption         =   "Параметры запуска DPInst"
      textboxheight   =   18
      themecolor      =   2
      headerstyle     =   1
      backcolor       =   14016736
      Begin VB.CommandButton cmdLegacyMode 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   37
         ToolTipText     =   "More on MSDN..."
         Top             =   660
         Width           =   255
      End
      Begin VB.CommandButton cmdPromptIfDriverIsNotBetter 
         Caption         =   "?"
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         ToolTipText     =   "More on MSDN..."
         Top             =   1305
         Width           =   255
      End
      Begin VB.CommandButton cmdForceIfDriverIsNotBetter 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   39
         ToolTipText     =   "More on MSDN..."
         Top             =   1905
         Width           =   255
      End
      Begin VB.CommandButton cmdSuppressAddRemovePrograms 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   40
         ToolTipText     =   "More on MSDN..."
         Top             =   2460
         Width           =   255
      End
      Begin VB.CommandButton cmdSuppressWizard 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   41
         ToolTipText     =   "More on MSDN..."
         Top             =   2955
         Width           =   255
      End
      Begin VB.CommandButton cmdQuietInstall 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   42
         ToolTipText     =   "More on MSDN..."
         Top             =   3510
         Width           =   255
      End
      Begin VB.CommandButton cmdScanHardware 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   51
         ToolTipText     =   "More on MSDN..."
         Top             =   4005
         Width           =   255
      End
      Begin prjDIADBS.TextBoxW txtCmdStringDPInst 
         Height          =   330
         Left            =   2895
         TabIndex        =   50
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
         Text            =   "frmOptions.frx":19A6
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":19C6
      End
      Begin prjDIADBS.CheckBoxW chkLegacyMode 
         Height          =   210
         Left            =   120
         TabIndex        =   43
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
         Caption         =   "frmOptions.frx":19E6
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkPromptIfDriverIsNotBetter 
         Height          =   210
         Left            =   120
         TabIndex        =   44
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
         Caption         =   "frmOptions.frx":1A1A
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkForceIfDriverIsNotBetter 
         Height          =   210
         Left            =   120
         TabIndex        =   45
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
         Caption         =   "frmOptions.frx":1A6C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSuppressAddRemovePrograms 
         CausesValidation=   0   'False
         Height          =   210
         Left            =   120
         TabIndex        =   46
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
         Caption         =   "frmOptions.frx":1ABC
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSuppressWizard 
         Height          =   210
         Left            =   120
         TabIndex        =   47
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
         Caption         =   "frmOptions.frx":1B0E
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkQuietInstall 
         Height          =   210
         Left            =   120
         TabIndex        =   48
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
         Caption         =   "frmOptions.frx":1B4A
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkScanHardware 
         Height          =   210
         Left            =   120
         TabIndex        =   49
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
         Caption         =   "frmOptions.frx":1B82
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblCmdStringDPInst 
         Height          =   210
         Left            =   135
         TabIndex        =   120
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
         TabIndex        =   121
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
         TabIndex        =   122
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
         TabIndex        =   123
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
         TabIndex        =   124
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
         TabIndex        =   125
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
         TabIndex        =   126
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
         TabIndex        =   127
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
         TabIndex        =   128
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
         TabIndex        =   129
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
         Caption         =   $"frmOptions.frx":1BBA
      End
   End
   Begin prjDIADBS.ctlJCFrames frDebug 
      Height          =   5295
      Left            =   4620
      Top             =   2460
      Width           =   8655
      _extentx        =   15266
      _extenty        =   9340
      font            =   "frmOptions.frx":1CB8
      fillcolor       =   14016736
      roundedcorner   =   0   'False
      caption         =   "Отладочный режим"
      textboxheight   =   18
      themecolor      =   2
      headerstyle     =   1
      backcolor       =   14016736
      Begin prjDIADBS.TextBoxW txtDebugLogName 
         Height          =   315
         Left            =   480
         TabIndex        =   87
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
         Text            =   "frmOptions.frx":1CE0
         CueBanner       =   "frmOptions.frx":1D24
      End
      Begin prjDIADBS.TextBoxW txtMacrosDate 
         Height          =   255
         Left            =   480
         TabIndex        =   82
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
         Text            =   "frmOptions.frx":1D44
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1D70
      End
      Begin prjDIADBS.TextBoxW txtMacrosOSBIT 
         Height          =   255
         Left            =   480
         TabIndex        =   81
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
         Text            =   "frmOptions.frx":1D90
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1DBE
      End
      Begin prjDIADBS.TextBoxW txtMacrosOSVER 
         Height          =   255
         Left            =   480
         TabIndex        =   80
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
         Text            =   "frmOptions.frx":1DDE
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1E0C
      End
      Begin prjDIADBS.TextBoxW txtMacrosPCModel 
         Height          =   255
         Left            =   480
         TabIndex        =   79
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
         Text            =   "frmOptions.frx":1E2C
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1E5E
      End
      Begin prjDIADBS.TextBoxW txtMacrosPCName 
         Height          =   255
         Left            =   480
         TabIndex        =   78
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
         Text            =   "frmOptions.frx":1E7E
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1EAE
      End
      Begin prjDIADBS.CheckBoxW chkDebug 
         Height          =   210
         Left            =   495
         TabIndex        =   83
         Top             =   750
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
         Caption         =   "frmOptions.frx":1ECE
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlUcPickBox ucDebugLogPath 
         Height          =   315
         Left            =   480
         TabIndex        =   84
         Top             =   1890
         Width           =   7845
         _extentx        =   10821
         _extenty        =   556
         useautoforecolor=   0   'False
         defaultext      =   ""
         enabled         =   0   'False
         filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.CheckBoxW chkDebugLog2AppPath 
         Height          =   210
         Left            =   495
         TabIndex        =   85
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
         Caption         =   "frmOptions.frx":1F1E
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkDebugTime2File 
         Height          =   210
         Left            =   495
         TabIndex        =   86
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
         Caption         =   "frmOptions.frx":1F9E
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblMacrosDate 
         Height          =   375
         Left            =   2400
         TabIndex        =   130
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
      Begin prjDIADBS.LabelW lblMacrosOSBit 
         Height          =   375
         Left            =   2400
         TabIndex        =   131
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
      Begin prjDIADBS.LabelW lblMacrosOSVer 
         Height          =   375
         Left            =   2400
         TabIndex        =   132
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
      Begin prjDIADBS.LabelW lblMacrosPCModel 
         Height          =   375
         Left            =   2400
         TabIndex        =   133
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
      Begin prjDIADBS.LabelW lblMacrosParam 
         Height          =   255
         Left            =   480
         TabIndex        =   134
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
      Begin prjDIADBS.LabelW lblMacrosDescription 
         Height          =   255
         Left            =   2400
         TabIndex        =   135
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
      Begin prjDIADBS.LabelW lblMacrosPCName 
         Height          =   375
         Left            =   2400
         TabIndex        =   136
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
      Begin prjDIADBS.LabelW lblMacrosType 
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
         TabIndex        =   138
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
         TabIndex        =   139
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
         TabIndex        =   140
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
      _extentx        =   15266
      _extenty        =   9340
      font            =   "frmOptions.frx":2020
      fillcolor       =   14016736
      roundedcorner   =   0   'False
      textboxheight   =   18
      themecolor      =   2
      headerstyle     =   1
      backcolor       =   14016736
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents lvOptions As cListView
Attribute lvOptions.VB_VarHelpID = -1

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

    frDebug.Font.Charset = lngFont_Charset
    frDesign.Font.Charset = lngFont_Charset
    frDesign2.Font.Charset = lngFont_Charset
    frDpInstParam.Font.Charset = lngFont_Charset
    frMain.Font.Charset = lngFont_Charset
    frMain2.Font.Charset = lngFont_Charset
    frMainTools.Font.Charset = lngFont_Charset
    frOptions.Font.Charset = lngFont_Charset
    frOS.Font.Charset = lngFont_Charset
    frOther.Font.Charset = lngFont_Charset
    frOtherTools.Font.Charset = lngFont_Charset
End Sub

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

    SetBtnFontProperties cmdFutureButton
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkButtonDisable_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkButtonDisable_Click()
    cmdFutureButton.EnabledCtrl = chkButtonDisable.Value
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
    DebugCtlEnable chkDebug.Value
    DebugCtlEnableLog2App Not chkDebugLog2AppPath.Value

    If Not chkDebug.Value Then
        If Not chkDebugLog2AppPath.Value Then
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
    DebugCtlEnableLog2App Not chkDebugLog2AppPath.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkForceIfDriverIsNotBetter_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkForceIfDriverIsNotBetter_Click()
    mbDpInstForceIfDriverIsNotBetter = chkForceIfDriverIsNotBetter.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkFormMaximaze_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkFormMaximaze_Click()

    If chkFormMaximaze.Value Then
        chkFormSizeSave.Value = False
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkFormSizeSave_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkFormSizeSave_Click()

    If chkFormSizeSave.Value Then
        chkFormMaximaze.Value = False
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkLegacyMode_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkLegacyMode_Click()
    mbDpInstLegacyMode = chkLegacyMode.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkPromptIfDriverIsNotBetter_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkPromptIfDriverIsNotBetter_Click()
    mbDpInstPromptIfDriverIsNotBetter = chkPromptIfDriverIsNotBetter.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkQuietInstall_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkQuietInstall_Click()
    mbDpInstQuietInstall = chkQuietInstall.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkScanHardware_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkScanHardware_Click()
    mbDpInstScanHardware = chkScanHardware.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkSuppressAddRemovePrograms_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkSuppressAddRemovePrograms_Click()
    mbDpInstSuppressAddRemovePrograms = chkSuppressAddRemovePrograms.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkSuppressWizard_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkSuppressWizard_Click()
    mbDpInstSuppressWizard = chkSuppressWizard.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkTabBlock_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkTabBlock_Click()
    Tab2CtlEnable chkTabBlock.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkTabHide_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkTabHide_Click()
    TabCtlEnable Not chkTabHide.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkTempPath_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkTempPath_Click()
    TempCtlEnable chkTempPath.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkUpdate_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkUpdate_Click()
    UpdateCtlEnable chkUpdate.Value
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
    Set cmdFutureButton.Picture = imgOK.Picture
    cmdFutureButton.Refresh
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

    Dim i As Long

    With lvOS

        If .ListItems.Count > 0 Then
            i = .SelectedItem.Index
            .ListItems.Remove (i)
            LastIdOS = LastIdOS - 1
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDelUtil_Click
'! Description (Описание)  :   [кнопка удаление утилиты]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdDelUtil_Click()

    Dim i As Long

    With lvUtils

        If .ListItems.Count > 0 Then
            i = .SelectedItem.Index
            .ListItems.Remove (i)
            LastIdUtil = LastIdUtil - 1
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDriverVer_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdDriverVer_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = Kavichki & "http://msdn.microsoft.com/en-us/library/ff547394%28VS.85%29.aspx?ppud=4" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
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
        .ForeColor = lngFontTab2_Color
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
        .Show vbModal, Me
    End With
    
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
        .Show vbModal, Me
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdForceIfDriverIsNotBetter_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdForceIfDriverIsNotBetter_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = Kavichki & "http://msdn.microsoft.com/en-us/library/ms793551.aspx" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdLegacyMode_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdLegacyMode_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = Kavichki & "http://msdn.microsoft.com/en-us/library/ms794322.aspx" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
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
    ElseIf Not FileisReadOnly(strSysIni) Then
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

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = Kavichki & "http://msdn.microsoft.com/en-us/library/ms793530.aspx" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdQuietInstall_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdQuietInstall_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = Kavichki & "http://msdn.microsoft.com/en-us/library/ms794300.aspx" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdScanHardware_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdScanHardware_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = Kavichki & "http://msdn.microsoft.com/en-us/library/ms794295.aspx" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdSuppressAddRemovePrograms_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdSuppressAddRemovePrograms_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = Kavichki & "http://msdn.microsoft.com/en-us/library/ms794270.aspx" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdSuppressWizard_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdSuppressWizard_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = Kavichki & "http://msdn.microsoft.com/en-us/library/ms791062.aspx" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
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
    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork
    LoadIconImage2BtnJC cmdAddUtil, "BTN_ADD", strPathImageMainWork
    LoadIconImage2BtnJC cmdEditUtil, "BTN_EDIT", strPathImageMainWork
    LoadIconImage2BtnJC cmdDelUtil, "BTN_DELETE", strPathImageMainWork
    LoadIconImage2BtnJC cmdAddOS, "BTN_ADD", strPathImageMainWork
    LoadIconImage2BtnJC cmdEditOS, "BTN_EDIT", strPathImageMainWork
    LoadIconImage2BtnJC cmdDelOS, "BTN_DELETE", strPathImageMainWork
    LoadIconImage2BtnJC cmdFontColorButton, "BTN_FONT", strPathImageMainWork
    LoadIconImage2BtnJC cmdFontColorTabOS, "BTN_FONT", strPathImageMainWork
    LoadIconImage2BtnJC cmdFontColorTabDrivers, "BTN_FONT", strPathImageMainWork
    LoadIconImage2BtnJC cmdFontColorToolTip, "BTN_FONT", strPathImageMainWork
    FormLoadAction
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FormLoadAction
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub FormLoadAction()

    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    ' загрузить список опций
    tvOptionsLoad
    ' Заполнить опции
    ReadOptions
    ' установить опции шрифта и цвета для будущей кнопки
    Set cmdFutureButton.Picture = imgOK.Picture
    cmdFutureButton.TextColor = lngFontBtn_Color
    SetBtnFontProperties cmdFutureButton
    'Загрузить подсказку
    LoadToolTip
    
    DoEvents
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
        Set lvOptions = Nothing
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
'! Procedure   (Функция)   :   Sub InitializeObjectProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub InitializeObjectProperties()

    With cmdFutureButton
        .CheckExist = True
    End With

    chkFutureButton.ZOrder 0
    ' изменение шрифта и текста
    ChangeButtonProperties
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadList_OS
'! Description (Описание)  :   [Построение спиcка ОС]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadList_OS()

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

    LastIdOS = lngOSCount
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadList_Utils
'! Description (Описание)  :   [Построение спика утилит]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadList_Utils()

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

    LastIdUtil = lngUtilsCount
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadListCombo
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   cmbName (ComboBox)
'                              strImagePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadListCombo(cmbName As ComboBox, strImagePath As String)

    Dim strListFolderTemp() As String
    Dim i                   As Integer

    strListFolderTemp = SearchFoldersInRoot(strImagePath, "*")

    With cmbName
        .Clear

        For i = LBound(strListFolderTemp, 2) To UBound(strListFolderTemp, 2)
            .AddItem strListFolderTemp(1, i), i
        Next

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal StrPathFile As String)

    Dim strFormNameMain As String

    strFormNameMain = frmMain.Name
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.Caption = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
    frOptions.Caption = LocaliseString(StrPathFile, strFormName, "frOptions", frOptions.Caption)
    ' Описание режимов
    optRezim_Intellect.Caption = LocaliseString(StrPathFile, strFormNameMain, "RezimIntellect", optRezim_Intellect.Caption)
    optRezim_Ust.Caption = LocaliseString(StrPathFile, strFormNameMain, "RezimUst", optRezim_Ust.Caption)
    optRezim_Upd.Caption = LocaliseString(StrPathFile, strFormNameMain, "RezimUpd", optRezim_Upd.Caption)
    optRezim_Intellect.ToolTipText = LocaliseString(StrPathFile, strFormNameMain, "RezimIntellectTip", optRezim_Intellect.ToolTipText)
    optRezim_Ust.ToolTipText = LocaliseString(StrPathFile, strFormNameMain, "RezimUstTip", optRezim_Ust.ToolTipText)
    optRezim_Upd.ToolTipText = LocaliseString(StrPathFile, strFormNameMain, "RezimUpdTip", optRezim_Upd.ToolTipText)
    strItemOptions1 = LocaliseString(StrPathFile, strFormName, "ItemOptions1", "Основные настройки")
    strItemOptions2 = LocaliseString(StrPathFile, strFormName, "ItemOptions2", "Поддерживаемые ОС")
    strItemOptions3 = LocaliseString(StrPathFile, strFormName, "ItemOptions3", "Рабочие утилиты")
    strItemOptions4 = LocaliseString(StrPathFile, strFormName, "ItemOptions4", "Вспомогательные утилиты")
    strItemOptions5 = LocaliseString(StrPathFile, strFormName, "ItemOptions5", "Оформление программы")
    strItemOptions6 = LocaliseString(StrPathFile, strFormName, "ItemOptions6", "Параметры запуска DPInst")
    strItemOptions8 = LocaliseString(StrPathFile, strFormName, "ItemOptions8", "Основные настройки 2")
    strItemOptions9 = LocaliseString(StrPathFile, strFormName, "ItemOptions9", "Оформление программы 2")
    strItemOptions10 = LocaliseString(StrPathFile, strFormName, "ItemOptions10", "Отладочный режим")
    cmdOK.Caption = LocaliseString(StrPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(StrPathFile, strFormName, "cmdExit", cmdExit.Caption)
    frMain.Caption = LocaliseString(StrPathFile, strFormName, "frMain", frMain.Caption)
    lblOptionsStart.Caption = LocaliseString(StrPathFile, strFormName, "lblOptionsStart", lblOptionsStart.Caption)
    chkUpdate.Caption = LocaliseString(StrPathFile, strFormName, "chkUpdate", chkUpdate.Caption)
    chkUpdateBeta.Caption = LocaliseString(StrPathFile, strFormName, "chkUpdateBeta", chkUpdateBeta.Caption)
    chkConvertDPName.Caption = LocaliseString(StrPathFile, strFormName, "chkConvertDPName", chkConvertDPName.Caption)
    chkReadDPName.Caption = LocaliseString(StrPathFile, strFormName, "chkReadDPName", chkReadDPName.Caption)
    chkHideOtherProcess.Caption = LocaliseString(StrPathFile, strFormName, "chkHideOtherProcess", chkHideOtherProcess.Caption)
    lblOptionsTemp.Caption = LocaliseString(StrPathFile, strFormName, "lblOptionsTemp", lblOptionsTemp.Caption)
    chkTempPath.Caption = LocaliseString(StrPathFile, strFormName, "chkTempPath", chkTempPath.Caption)
    chkRemoveTemp.Caption = LocaliseString(StrPathFile, strFormName, "chkRemoveTemp", chkRemoveTemp.Caption)
    lblDebug.Caption = LocaliseString(StrPathFile, strFormName, "lblDebug", lblDebug.Caption)
    chkDebug.Caption = LocaliseString(StrPathFile, strFormName, "chkDebug", chkDebug.Caption)
    lblRezim.Caption = LocaliseString(StrPathFile, strFormName, "lblRezim", lblRezim.Caption)
    lblDebugLogPath.Caption = LocaliseString(StrPathFile, strFormName, "lblDebugLogPath", lblDebugLogPath.Caption)
    frMainTools.Caption = LocaliseString(StrPathFile, strFormName, "frMainTools", frMainTools.Caption)
    cmdPathDefault.Caption = LocaliseString(StrPathFile, strFormName, "cmdPathDefault", cmdPathDefault.Caption)
    frOtherTools.Caption = LocaliseString(StrPathFile, strFormName, "frOtherTools", frOtherTools.Caption)
    cmdAddUtil.Caption = LocaliseString(StrPathFile, strFormName, "cmdAddUtil", cmdAddUtil.Caption)
    cmdEditUtil.Caption = LocaliseString(StrPathFile, strFormName, "cmdEditUtil", cmdEditUtil.Caption)
    cmdDelUtil.Caption = LocaliseString(StrPathFile, strFormName, "cmdDelUtil", cmdDelUtil.Caption)
    frOS.Caption = LocaliseString(StrPathFile, strFormName, "frOS", frOS.Caption)
    chkRecursion.Caption = LocaliseString(StrPathFile, strFormName, "chkRecursion", chkRecursion.Caption)
    chkLoadFinishFile.Caption = LocaliseString(StrPathFile, strFormName, "chkLoadFinishFile", chkLoadFinishFile.Caption)
    cmdAddOS.Caption = LocaliseString(StrPathFile, strFormName, "cmdAddOS", cmdAddOS.Caption)
    cmdEditOS.Caption = LocaliseString(StrPathFile, strFormName, "cmdEditOS", cmdEditOS.Caption)
    cmdDelOS.Caption = LocaliseString(StrPathFile, strFormName, "cmdDelOS", cmdDelOS.Caption)
    lblExcludeHWID.Caption = LocaliseString(StrPathFile, strFormName, "lblExcludeHWID", lblExcludeHWID.Caption)
    frDesign.Caption = LocaliseString(StrPathFile, strFormName, "frDesign", frDesign.Caption)
    frDesign2.Caption = LocaliseString(StrPathFile, strFormName, "frDesign2", frDesign2.Caption)
    lblSizeForm.Caption = LocaliseString(StrPathFile, strFormName, "lblSizeForm", lblSizeForm.Caption)
    lblFormHeight.Caption = LocaliseString(StrPathFile, strFormName, "lblFormHeight", lblFormHeight.Caption)
    lblFormWidth.Caption = LocaliseString(StrPathFile, strFormName, "lblFormWidth", lblFormWidth.Caption)
    chkFormMaximaze.Caption = LocaliseString(StrPathFile, strFormName, "chkFormMaximaze", chkFormMaximaze.Caption)
    chkFormSizeSave.Caption = LocaliseString(StrPathFile, strFormName, "chkFormSizeSave", chkFormSizeSave.Caption)
    lblSizeButton.Caption = LocaliseString(StrPathFile, strFormName, "lblSizeButton", lblSizeButton.Caption)
    lblButtonHeight.Caption = LocaliseString(StrPathFile, strFormName, "lblButtonHeight", lblButtonHeight.Caption)
    lblButtonWidth.Caption = LocaliseString(StrPathFile, strFormName, "lblButtonWidth", lblButtonWidth.Caption)
    lblButtonLeft.Caption = LocaliseString(StrPathFile, strFormName, "lblButtonLeft", lblButtonLeft.Caption)
    lblButtonTop.Caption = LocaliseString(StrPathFile, strFormName, "lblButtonTop", lblButtonTop.Caption)
    lblButton2BtnL.Caption = LocaliseString(StrPathFile, strFormName, "lblButton2BtnL", lblButton2BtnL.Caption)
    lblButton2BtnT.Caption = LocaliseString(StrPathFile, strFormName, "lblButton2BtnT", lblButton2BtnT.Caption)
    chkButtonTextUpCase.Caption = LocaliseString(StrPathFile, strFormName, "chkButtonTextUpCase", chkButtonTextUpCase.Caption)
    cmdFutureButton.Caption = LocaliseString(StrPathFile, strFormName, "cmdFutureButton", cmdFutureButton.Caption)
    lblImageMain.Caption = LocaliseString(StrPathFile, strFormName, "lblImageMain", lblImageMain.Caption)
    lblImageStatus.Caption = LocaliseString(StrPathFile, strFormName, "lblImageStatus", lblImageStatus.Caption)
    lblTabControl.Caption = LocaliseString(StrPathFile, strFormName, "lblTabControl", lblTabControl.Caption)
    lblTabControl2.Caption = LocaliseString(StrPathFile, strFormName, "lblTabControl2", lblTabControl2.Caption)
    lblTabPerRowCount.Caption = LocaliseString(StrPathFile, strFormName, "lblTabPerRowCount", lblTabPerRowCount.Caption)
    chkTabBlock.Caption = LocaliseString(StrPathFile, strFormName, "chkTabBlock", chkTabBlock.Caption)
    chkTabHide.Caption = LocaliseString(StrPathFile, strFormName, "chkTabHide", chkTabHide.Caption)
    frDpInstParam.Caption = LocaliseString(StrPathFile, strFormName, "frDpInstParam", frDpInstParam.Caption)
    lblParam.Caption = LocaliseString(StrPathFile, strFormName, "lblParam", lblParam.Caption)
    lblDescription.Caption = LocaliseString(StrPathFile, strFormName, "lblDescription", lblDescription.Caption)
    lblLegacyMode.Caption = LocaliseString(StrPathFile, strFormName, "lblLegacyMode", lblLegacyMode.Caption)
    lblPromptIfDriverIsNotBetter.Caption = LocaliseString(StrPathFile, strFormName, "lblPromptIfDriverIsNotBetter", lblPromptIfDriverIsNotBetter.Caption)
    lblForceIfDriverIsNotBetter.Caption = LocaliseString(StrPathFile, strFormName, "lblForceIfDriverIsNotBetter", lblForceIfDriverIsNotBetter.Caption)
    lblSuppressAddRemovePrograms.Caption = LocaliseString(StrPathFile, strFormName, "lblSuppressAddRemovePrograms", lblSuppressAddRemovePrograms.Caption)
    lblSuppressWizard.Caption = LocaliseString(StrPathFile, strFormName, "lblSuppressWizard", lblSuppressWizard.Caption)
    lblQuietInstall.Caption = LocaliseString(StrPathFile, strFormName, "lblQuietInstall", lblQuietInstall.Caption)
    lblScanHardware.Caption = LocaliseString(StrPathFile, strFormName, "lblScanHardware", lblScanHardware.Caption)
    lblCmdStringDPInst.Caption = LocaliseString(StrPathFile, strFormName, "lblCmdStringDPInst", lblCmdStringDPInst.Caption)
    strTableOSHeader1 = LocaliseString(StrPathFile, strFormName, "TableOSHeader1", "Версия")
    strTableOSHeader2 = LocaliseString(StrPathFile, strFormName, "TableOSHeader2", "Наименование")
    strTableOSHeader3 = LocaliseString(StrPathFile, strFormName, "TableOSHeader3", "Пакеты драйверов")
    strTableOSHeader4 = LocaliseString(StrPathFile, strFormName, "TableOSHeader4", "База данных")
    strTableOSHeader5 = LocaliseString(StrPathFile, strFormName, "TableOSHeader5", "x64")
    strTableOSHeader6 = LocaliseString(StrPathFile, strFormName, "TableOSHeader6", "PhysX")
    strTableOSHeader7 = LocaliseString(StrPathFile, strFormName, "TableOSHeader7", "Lang")
    strTableOSHeader8 = LocaliseString(StrPathFile, strFormName, "TableOSHeader8", "ExludeFiles")
    strTableOSHeader9 = LocaliseString(StrPathFile, strFormName, "TableOSHeader9", "ExludeFiles")
    strTableUtilHeader1 = LocaliseString(StrPathFile, strFormName, "TableUtilHeader1", "Наименование")
    strTableUtilHeader2 = LocaliseString(StrPathFile, strFormName, "TableUtilHeader2", "Путь")
    strTableUtilHeader3 = LocaliseString(StrPathFile, strFormName, "TableUtilHeader3", "Параметр")
    strTableUtilHeader4 = LocaliseString(StrPathFile, strFormName, "TableUtilHeader4", "Путь x64")
    frMain2.Caption = LocaliseString(StrPathFile, strFormName, "frMain2", frMain2.Caption)
    lblCompareVersionDRV.Caption = LocaliseString(StrPathFile, strFormName, "lblCompareVersionDRV", lblCompareVersionDRV.Caption)
    optCompareByDate.Caption = LocaliseString(StrPathFile, strFormName, "optCompareByDate", optCompareByDate.Caption)
    optCompareByVersion.Caption = LocaliseString(StrPathFile, strFormName, "optCompareByVersion", optCompareByVersion.Caption)
    txtCompareVersionDRV.Text = LocaliseString(StrPathFile, strFormName, "txtCompareVersionDRV", txtCompareVersionDRV.Text)
    chkSilentDll.Caption = LocaliseString(StrPathFile, strFormName, "chkSilentDll", chkSilentDll.Caption)
    chkDateFormatRus.Caption = LocaliseString(StrPathFile, strFormName, "chkDateFormatRus", chkDateFormatRus.Caption)
    chkSearchOnStart.Caption = LocaliseString(StrPathFile, strFormName, "chkSearchOnStart", chkSearchOnStart.Caption)
    lblPauseAfterSearch.Caption = LocaliseString(StrPathFile, strFormName, "lblPauseAfterSearch", lblPauseAfterSearch.Caption)
    chkCreateRP.Caption = LocaliseString(StrPathFile, strFormName, "chkCreateRP", chkCreateRP.Caption)
    chkCompatiblesHWID.Caption = LocaliseString(StrPathFile, strFormName, "chkCompatiblesHWID", chkCompatiblesHWID.Caption)
    chkLoadUnSupportedOS.Caption = LocaliseString(StrPathFile, strFormName, "chkLoadUnSupportedOS", chkLoadUnSupportedOS.Caption)
    chkDebugLog2AppPath.Caption = LocaliseString(StrPathFile, strFormName, "chkDebugLog2AppPath", chkDebugLog2AppPath.Caption)
    frDebug.Caption = LocaliseString(StrPathFile, strFormName, "frDebug", frDebug.Caption)
    lblMacrosType.Caption = LocaliseString(StrPathFile, strFormName, "lblMacrosType", lblMacrosType.Caption)
    lblMacrosParam.Caption = LocaliseString(StrPathFile, strFormName, "lblMacrosParam", lblMacrosParam.Caption)
    lblMacrosDescription.Caption = LocaliseString(StrPathFile, strFormName, "lblMacrosDescription", lblMacrosDescription.Caption)
    lblMacrosPCName.Caption = LocaliseString(StrPathFile, strFormName, "lblMacrosPCName", lblMacrosPCName.Caption)
    lblMacrosPCModel.Caption = LocaliseString(StrPathFile, strFormName, "lblMacrosPCModel", lblMacrosPCModel.Caption)
    lblMacrosOSVer.Caption = LocaliseString(StrPathFile, strFormName, "lblMacrosOSVer", lblMacrosOSVer.Caption)
    lblMacrosOSBit.Caption = LocaliseString(StrPathFile, strFormName, "lblMacrosOSBit", lblMacrosOSBit.Caption)
    lblMacrosDate.Caption = LocaliseString(StrPathFile, strFormName, "lblMacrosDate", lblMacrosDate.Caption)
    chkDebugTime2File.Caption = LocaliseString(StrPathFile, strFormName, "chkDebugTime2File", chkDebugTime2File.Caption)
    lblDebugLogName.Caption = LocaliseString(StrPathFile, strFormName, "lblDebugLogName", lblDebugLogName.Caption)
    cmdFontColorButton.Caption = LocaliseString(StrPathFile, strFormName, "cmdFontColorButton", cmdFontColorButton.Caption)
    cmdFontColorTabOS.Caption = LocaliseString(StrPathFile, strFormName, "cmdFontColorTabOS", cmdFontColorTabOS.Caption)
    cmdFontColorTabDrivers.Caption = LocaliseString(StrPathFile, strFormName, "cmdFontColorTabDrivers", cmdFontColorTabDrivers.Caption)
    chkButtonDisable.Caption = LocaliseString(StrPathFile, strFormName, "chkButtonDisable", chkButtonDisable.Caption)
    lblTheme.Caption = LocaliseString(StrPathFile, strFormName, "lblTheme", lblTheme.Caption)
    cmdFontColorToolTip.Caption = LocaliseString(StrPathFile, strFormName, "cmdFontColorToolTip", cmdFontColorToolTip.Caption)
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvOptions_ItemChanged
'! Description (Описание)  :   [При выборе опции происходит отображение соответсвующего окна]
'! Parameters  (Переменные):   iIndex (Long)
'!--------------------------------------------------------------------------------
Private Sub lvOptions_ItemChanged(ByVal iIndex As Long)

    Select Case iIndex

        Case 0
        'ItemOptions1=Основные настройки
            frMain.ZOrder 0

        Case 1
        ' ItemOptions8=Основные настройки 2
            frMain2.ZOrder 0

        Case 2
        'ItemOptions2=Поддерживаемые ОС
            frOS.ZOrder 0
            txtExcludeHWID.SetFocus

        Case 3
        'ItemOptions3=Рабочие утилиты
            frMainTools.ZOrder 0
            ucDevCon86Path.SetFocus

        Case 4
        'ItemOptions4=Вспомогательные утилиты
            frOtherTools.ZOrder 0

        Case 5
        'ItemOptions5=Оформление программы
            frDesign.ZOrder 0
            cmbImageMain.SetFocus

        Case 6
        'ItemOptions9=Оформление программы 2
            frDesign2.ZOrder 0

        Case 7
        'ItemOptions6=Параметры запуска DPInst
            frDpInstParam.ZOrder 0
            txtCmdStringDPInst.SetFocus

        Case 8
        'ItemOptions10=Отладочный режим
            frDebug.ZOrder 0
            txtDebugLogName.SetFocus

        Case Else
            frOther.ZOrder 0
    End Select

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvOS_ItemDblClick
'! Description (Описание)  :   [Двойнок клик по элементу списка вызывает форму редактирования]
'! Parameters  (Переменные):   Item (LvwListItem)
'                              Button (Integer)
'!--------------------------------------------------------------------------------
Private Sub lvOS_ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)
    TransferOSData
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvUtils_ItemDblClick
'! Description (Описание)  :   [Двойнок клик по элементу списка вызывает форму редактирования]
'! Parameters  (Переменные):   Item (LvwListItem)
'                              Button (Integer)
'!--------------------------------------------------------------------------------
Private Sub lvUtils_ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)
    TransferUtilsData
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ReadOptions
'! Description (Описание)  :   [Читаем настройки программы и заполняем поля]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ReadOptions()
    ' загрузить список ОС
    LoadList_OS
    ' загрузить список Утилит
    LoadList_Utils
    ' Остальные параметры
    chkUpdate.Value = mbUpdateCheck
    chkUpdateBeta.Value = mbUpdateCheckBeta
    chkSilentDll.Value = mbSilentDLL
    chkRemoveTemp.Value = mbDelTmpAfterClose
    chkDebug.Value = mbDebugEnable
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
    LoadListCombo cmbImageMain, strPathImageMain
    LoadListCombo cmbImageStatus, strPathImageStatusButton
    LoadIconImage2Object imgOK, "BTN_OK", strPathImageStatusButtonWork
    cmbImageMain.Text = strImageMainName
    cmbImageStatus.Text = strImageStatusButtonName
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
    DebugCtlEnable chkDebug.Value
    DebugCtlEnableLog2App Not chkDebugLog2AppPath.Value
    TempCtlEnable chkTempPath.Value
    UpdateCtlEnable chkUpdate.Value
    TabCtlEnable Not chkTabHide.Value
    Tab2CtlEnable chkTabBlock.Value
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
        If strSysIni <> strWorkTempBackSL & "DriversInstaller.ini" Then
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
    IniWriteStrPrivate "Main", "DelTmpAfterClose", CStr(Abs(chkRemoveTemp.Value)), strSysIniTemp
    ' Автообновление
    IniWriteStrPrivate "Main", "UpdateCheck", CStr(Abs(chkUpdate.Value)), strSysIniTemp
    ' Автообновление Beta
    IniWriteStrPrivate "Main", "UpdateCheckBeta", CStr(Abs(chkUpdateBeta.Value)), strSysIniTemp

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

    IniWriteStrPrivate "Main", "StartMode", CStr(miRezim), strSysIniTemp
    IniWriteStrPrivate "Main", "EULAAgree", CStr(Abs(mbEULAAgree)), strSysIniTemp
    IniWriteStrPrivate "Main", "HideOtherProcess", CStr(Abs(chkHideOtherProcess.Value)), strSysIniTemp
    IniWriteStrPrivate "Main", "AlternativeTemp", CStr(Abs(chkTempPath.Value)), strSysIniTemp
    IniWriteStrPrivate "Main", "AlternativeTempPath", ucTempPath.Path, strSysIniTemp
    IniWriteStrPrivate "Main", "IconMainSkin", cmbImageMain.Text, strSysIniTemp
    IniWriteStrPrivate "Main", "SilentDLL", CStr(Abs(chkSilentDll.Value)), strSysIniTemp
    ' поиск новых устройств при старте
    IniWriteStrPrivate "Main", "SearchOnStart", CStr(Abs(chkSearchOnStart.Value)), strSysIniTemp
    IniWriteStrPrivate "Main", "PauseAfterSearch", txtPauseAfterSearch.Value, strSysIniTemp
    ' Создание точки восстановления
    IniWriteStrPrivate "Main", "CreateRestorePoint", CStr(Abs(chkCreateRP.Value)), strSysIniTemp

    If mbLoadIniTmpAfterRestart Then
        IniWriteStrPrivate "Main", "LoadIniTmpAfterRestart", 1, strSysIniTemp
    End If

    ' Секция Debug
    IniWriteStrPrivate "Debug", "DebugEnable", CStr(Abs(chkDebug.Value)), strSysIniTemp
    ' Очистка истории:
    IniWriteStrPrivate "Debug", "CleenHistory", 1, strSysIniTemp
    ' Путь до лог-файла
    IniWriteStrPrivate "Debug", "DebugLog2AppPath", CStr(Abs(chkDebugLog2AppPath.Value)), strSysIniTemp
    IniWriteStrPrivate "Debug", "DebugLogPath", ucDebugLogPath.Path, strSysIniTemp
    strLogNameTemp = "DIA-LOG_%DATE%.txt"

    If LenB(txtDebugLogName.Text) > 0 Then
        If InStr(txtDebugLogName.Text, ".") Then
            strLogNameTemp = txtDebugLogName.Text
        End If
    End If

    IniWriteStrPrivate "Debug", "DebugLogName", strLogNameTemp, strSysIniTemp
    IniWriteStrPrivate "Debug", "Detailmode", CStr(lngDetailMode), strSysIniTemp
    'Секция Devcon
    IniWriteStrPrivate "Devcon", "PathExe", ucDevCon86Path.Path, strSysIniTemp
    IniWriteStrPrivate "Devcon", "PathExe64", ucDevCon64Path.Path, strSysIniTemp
    IniWriteStrPrivate "Devcon", "PathExeW2K", ucDevCon86Pathw2k.Path, strSysIniTemp
    IniWriteStrPrivate "Devcon", "CollectHwidsCmd", ucCmdDevconPath.Path, strSysIniTemp
    'Секция DPInst
    IniWriteStrPrivate "DPInst", "PathExe", ucDPInst86Path.Path, strSysIniTemp
    IniWriteStrPrivate "DPInst", "PathExe64", ucDPInst64Path.Path, strSysIniTemp
    IniWriteStrPrivate "DPInst", "LegacyMode", CStr(Abs(chkLegacyMode.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "PromptIfDriverIsNotBetter", CStr(Abs(chkPromptIfDriverIsNotBetter.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "ForceIfDriverIsNotBetter", CStr(Abs(chkForceIfDriverIsNotBetter.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "SuppressAddRemovePrograms", CStr(Abs(chkSuppressAddRemovePrograms.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "SuppressWizard", CStr(Abs(chkSuppressWizard.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "QuietInstall", CStr(Abs(chkQuietInstall.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "ScanHardware", CStr(Abs(chkScanHardware.Value)), strSysIniTemp
    'Секция Arc
    IniWriteStrPrivate "Arc", "PathExe", ucArchPath.Path, strSysIniTemp
    'Секция OS
    'Число ОС
    lngOSCountNew = lvOS.ListItems.Count
    IniWriteStrPrivate "OS", "OSCount", CStr(lngOSCountNew), strSysIniTemp
    ' Рекурсивный перебор папок
    IniWriteStrPrivate "OS", "Recursion", CStr(Abs(chkRecursion.Value)), strSysIniTemp
    ' кол-во табов на строку
    IniWriteStrPrivate "OS", "OSCountPerRow", txtTabPerRowCount.Value, strSysIniTemp
    ' Блокировать лишние вкладки
    IniWriteStrPrivate "OS", "TabBlock", CStr(Abs(chkTabBlock.Value)), strSysIniTemp
    ' Скрывать лишние вкладки
    IniWriteStrPrivate "OS", "TabHide", CStr(Abs(chkTabHide.Value)), strSysIniTemp
    ' Обрабатывать файлы Finish
    IniWriteStrPrivate "OS", "LoadFinishFile", CStr(Abs(chkLoadFinishFile.Value)), strSysIniTemp
    ' Считывать класс пакета драйверов из Finish
    'IniWriteStrPrivate "OS", "ReadClasses", CStr(Abs(chkReadClasses.value)), strSysIniTemp
    ' Считывать класс пакета драйверов из Finish
    IniWriteStrPrivate "OS", "ReadDPName", CStr(Abs(chkReadDPName.Value)), strSysIniTemp
    ' Считывать класс пакета драйверов из Finish
    IniWriteStrPrivate "OS", "ConvertDPName", CStr(Abs(chkConvertDPName.Value)), strSysIniTemp
    IniWriteStrPrivate "OS", "ExcludeHWID", txtExcludeHWID.Text, strSysIniTemp
    ' Сравнение версий драйверов
    IniWriteStrPrivate "OS", "CompareDrvVerByDate", CStr(Abs(optCompareByDate.Value)), strSysIniTemp
    IniWriteStrPrivate "OS", "DateFormatRus", CStr(Abs(chkDateFormatRus.Value)), strSysIniTemp
    ' Необрабатывать вкладки
    IniWriteStrPrivate "OS", "LoadUnSupportedOS", CStr(Abs(chkLoadUnSupportedOS.Value)), strSysIniTemp
    ' Обрабатывать совместимые HWID
    IniWriteStrPrivate "OS", "CompatiblesHWID", CStr(Abs(chkCompatiblesHWID.Value)), strSysIniTemp

    'Заполяем в цикле подсекции ОС
    For cnt = 1 To lngOSCountNew

        'Секция OS_N
        With lvOS.ListItems(cnt)
            IniWriteStrPrivate "OS_" & cnt, "Ver", .Text, strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "Name", .SubItems(1), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "drpFolder", .SubItems(2), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "devIDFolder", .SubItems(3), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "is64bit", .SubItems(4), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "PathPhysX", .SubItems(5), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "PathLanguages", .SubItems(6), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "PathRuntimes", .SubItems(7), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "ExcludeFileName", .SubItems(8), strSysIniTemp
        End With

    Next

    'Секция Utils
    'Число утилит
    lngUtilsCountNew = lvUtils.ListItems.Count
    IniWriteStrPrivate "Utils", "UtilsCount", CStr(lngUtilsCountNew), strSysIniTemp

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
    IniWriteStrPrivate "MainForm", "StartMaximazed", CStr(Abs(chkFormMaximaze.Value)), strSysIniTemp
    mbSaveSizeOnExit = chkFormSizeSave.Value
    IniWriteStrPrivate "MainForm", "SaveSizeOnExit", CStr(Abs(chkFormSizeSave.Value)), strSysIniTemp
    IniWriteStrPrivate "MainForm", "HighlightColor", CStr(glHighlightColor), strSysIniTemp
    'Секция Buttons
    IniWriteStrPrivate "Button", "Width", txtButtonWidth.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Height", txtButtonHeight.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Left", txtButtonLeft.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Top", txtButtonTop.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Btn2BtnLeft", txtButton2BtnL.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Btn2BtnTop", txtButton2BtnT.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "TextUpCase", CStr(Abs(chkButtonTextUpCase.Value)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontName", strFontBtn_Name, strSysIniTemp
    IniWriteStrPrivate "Button", "FontSize", CStr(miFontBtn_Size), strSysIniTemp
    IniWriteStrPrivate "Button", "FontUnderline", CStr(Abs(mbFontBtn_Underline)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontStrikethru", CStr(Abs(mbFontBtn_Strikethru)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontItalic", CStr(Abs(mbFontBtn_Italic)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontBold", CStr(Abs(mbFontBtn_Bold)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontColor", CStr(cmdFutureButton.TextColor), strSysIniTemp
    IniWriteStrPrivate "Button", "IconStatusSkin", cmbImageStatus.Text, strSysIniTemp
    'Секция Tab
    IniWriteStrPrivate "Tab", "FontName", strFontTab_Name, strSysIniTemp
    IniWriteStrPrivate "Tab", "FontSize", CStr(miFontTab_Size), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontUnderline", CStr(Abs(mbFontTab_Underline)), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontStrikethru", CStr(Abs(mbFontTab_Strikethru)), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontItalic", CStr(Abs(mbFontTab_Italic)), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontBold", CStr(Abs(mbFontTab_Bold)), strSysIniTemp
    'Секция Tab2
    IniWriteStrPrivate "Tab2", "FontName", strFontTab2_Name, strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontSize", CStr(miFontTab2_Size), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontUnderline", CStr(Abs(mbFontTab2_Underline)), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontStrikethru", CStr(Abs(mbFontTab2_Strikethru)), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontItalic", CStr(Abs(mbFontTab2_Italic)), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontBold", CStr(Abs(mbFontTab2_Bold)), strSysIniTemp
    'Секция ToolTip
    IniWriteStrPrivate "ToolTip", "FontName", strFontTT_Name, strSysIniTemp
    IniWriteStrPrivate "ToolTip", "FontSize", CStr(miFontTT_Size), strSysIniTemp
    IniWriteStrPrivate "ToolTip", "FontUnderline", CStr(Abs(mbFontTT_Underline)), strSysIniTemp
    IniWriteStrPrivate "ToolTip", "FontStrikethru", CStr(Abs(mbFontTT_Strikethru)), strSysIniTemp
    IniWriteStrPrivate "ToolTip", "FontItalic", CStr(Abs(mbFontTT_Italic)), strSysIniTemp
    IniWriteStrPrivate "ToolTip", "FontBold", CStr(Abs(mbFontTT_Bold)), strSysIniTemp
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

        If i = -1 Then

            Exit Sub

        End If

        frmOSEdit.txtOSVer.Text = .ListItems.Item(i).Text
        frmOSEdit.txtOSName.Text = .ListItems.Item(i).SubItems(1)
        frmOSEdit.ucPathDRP.Path = .ListItems.Item(i).SubItems(2)
        frmOSEdit.ucPathDB.Path = .ListItems.Item(i).SubItems(3)
        frmOSEdit.chk64bit.Value = CBool(.ListItems.Item(i).SubItems(4))

        Select Case .ListItems.Item(i).SubItems(4)

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

        frmOSEdit.ucPhysXPath.Path = .ListItems.Item(i).SubItems(5)
        frmOSEdit.ucLangPath.Path = .ListItems.Item(i).SubItems(6)
        frmOSEdit.ucRuntimesPath.Path = .ListItems.Item(i).SubItems(7)
        frmOSEdit.txtExcludeFileName.Text = .ListItems.Item(i).SubItems(8)
    End With

    'LVOS
    frmOSEdit.Show vbModal, Me
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

        If i = -1 Then

            Exit Sub

        End If

        frmUtilsEdit.txtUtilName.Text = .ListItems.Item(i).Text
        frmUtilsEdit.ucPathUtil.Path = .ListItems.Item(i).SubItems(1)
        frmUtilsEdit.ucPathUtil64.Path = .ListItems.Item(i).SubItems(2)
        frmUtilsEdit.txtParamUtil.Text = .ListItems.Item(i).SubItems(3)
    End With

    frmUtilsEdit.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub tvOptionsLoad
'! Description (Описание)  :   [Построение дерева настроек]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub tvOptionsLoad()
    Set lvOptions = New cListView

    With lvOptions
        .Create frOptions.hWnd, LVS_LIST Or LVS_SINGLESEL Or LVS_SHOWSELALWAYS, 10, 29, 180, 195, , WS_EX_STATICEDGE
        .SetStyleEx LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_MAIN", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_MAIN2", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_OSLIST", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_TOOLS_MAIN", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_TOOLS_OTHER", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_DESIGN", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_DESIGN2", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_DPINST", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_DEVPARSER", strPathImageMainWork)
        .AddItem strItemOptions1, , 0
        .AddItem strItemOptions8, , 1
        .AddItem strItemOptions2, , 2
        .AddItem strItemOptions3, , 3
        .AddItem strItemOptions4, , 4
        .AddItem strItemOptions5, , 5
        .AddItem strItemOptions9, , 6
        .AddItem strItemOptions6, , 7
        .AddItem strItemOptions10, , 8
    End With

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

    If ucArchPath.FileCount > 0 Then
        strTempPath = ucArchPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
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

    If ucCmdDevconPath.FileCount > 0 Then
        strTempPath = ucCmdDevconPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
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

    If ucDebugLogPath.FileCount > 0 Then
        strTempPath = ucDebugLogPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
        ucDebugLogPath.Path = strTempPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDevCon64Path_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon64Path_Click()

    Dim strTempPath As String

    If ucDevCon64Path.FileCount > 0 Then
        strTempPath = ucDevCon64Path.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
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

    If ucDevCon86Path.FileCount > 0 Then
        strTempPath = ucDevCon86Path.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
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

    If ucDevCon86Pathw2k.FileCount > 0 Then
        strTempPath = ucDevCon86Pathw2k.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
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

    If ucDPInst64Path.FileCount > 0 Then
        strTempPath = ucDPInst64Path.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
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

    If ucDPInst86Path.FileCount > 0 Then
        strTempPath = ucDPInst86Path.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
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

    If ucTempPath.FileCount > 0 Then
        strTempPath = ucTempPath.Path

        If InStr(1, strTempPath, strAppPath, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
        ucTempPath.Path = strTempPath
    End If

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
'! Procedure   (Функция)   :   Sub LoadToolTip
'! Description (Описание)  :   [Инициализация подсказки для "будущей" кнопки]
'! Parameters  (Переменные):   ColumnHeader (LvwColumnHeader)
'!--------------------------------------------------------------------------------
Private Sub LoadToolTip()
Dim strTTText As String
Dim strTTipTextTitle As String

    strTTipTextTitle = LocaliseString(strPCLangCurrentPath, frmMain.Name, "ToolTipTextTitle", "Файл пакета драйверов:")
    strTTText = "d:\DriversInstaller\driverpacks.net\All\" & vbNewLine & vbNewLine & _
                "dp_chipset_wnt5_x86-32_1209.7z" & vbNewLine & _
                "File Size: 4,33 МБ" & vbNewLine & _
                "Class of the Drivers: System" & vbNewLine & vbNewLine & _
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
