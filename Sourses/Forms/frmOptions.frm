VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������� ���������"
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
      Caption         =   "���������"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ctlJCbutton cmdOK 
         Height          =   750
         Left            =   75
         TabIndex        =   0
         Top             =   3500
         Width           =   2850
         _ExtentX        =   5027
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
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "��������� ��������� � �����"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdExit 
         Height          =   735
         Left            =   75
         TabIndex        =   18
         Top             =   4400
         Width           =   2850
         _ExtentX        =   5027
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
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "����� ��� ����������"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
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
      Caption         =   "�������� ��������� ���������"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
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
         Caption         =   "frmOptions.frx":058A
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
         Caption         =   "frmOptions.frx":0602
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
         Caption         =   "frmOptions.frx":065E
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
         Caption         =   "frmOptions.frx":06D2
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
         Caption         =   "frmOptions.frx":07A2
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
         Caption         =   "frmOptions.frx":0808
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
         Caption         =   "frmOptions.frx":0858
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
         Caption         =   "frmOptions.frx":08CE
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
         Caption         =   "frmOptions.frx":096A
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
         TabIndex        =   59
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
         Caption         =   "��������� (����������� ��������)"
         Mode            =   2
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjDIADBS.ctlJCbutton optRezim_Upd 
         Height          =   510
         Left            =   5700
         TabIndex        =   60
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
         Caption         =   "�������� ��� ���������� ���� ���������"
         Mode            =   2
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjDIADBS.ctlJCbutton optRezim_Ust 
         Height          =   510
         Left            =   3060
         TabIndex        =   61
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
         Caption         =   "��������� (������ - ���� �����)"
         Mode            =   2
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
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
         Caption         =   "����� ����� ������: "
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
         Caption         =   "������ � ���������� �������"
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
         Caption         =   "�������� ��� ������� ���������"
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
         Caption         =   "����� ������ ��� ������ ���������"
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
      Caption         =   "�������� ��������� ��������� 2"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
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
         Caption         =   "frmOptions.frx":09E8
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
         Caption         =   "frmOptions.frx":0A6A
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
         Text            =   "frmOptions.frx":0B18
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         CueBanner       =   "frmOptions.frx":0D0C
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
         Caption         =   "frmOptions.frx":0D2C
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
         Caption         =   "frmOptions.frx":0DA6
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
         Caption         =   "frmOptions.frx":0E2E
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
         Caption         =   "��������� ������ ���������"
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
      Caption         =   "������������ �������� ������ (Tools)"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ctlUcPickBox ucDevCon86Path 
         Height          =   315
         Left            =   2520
         TabIndex        =   27
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
         TabIndex        =   29
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   8
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
         TabIndex        =   64
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
         Caption         =   "�������� ��������� ������������ ������"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
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
      Caption         =   "��������������� ������� (������������ � ���� ""�������"")"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
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
         Icons           =   "frmOptions.frx":0EA8
         SmallIcons      =   "frmOptions.frx":0ED4
         ColumnHeaderIcons=   "frmOptions.frx":0F00
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
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "��������"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdEditUtil 
         Height          =   750
         Left            =   2160
         TabIndex        =   69
         Top             =   4455
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
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "��������"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdDelUtil 
         Height          =   750
         Left            =   4200
         TabIndex        =   70
         Top             =   4455
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
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "�������"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
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
      Caption         =   "�������������� ��"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
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
         Icons           =   "frmOptions.frx":0F2C
         SmallIcons      =   "frmOptions.frx":0F58
         ColumnHeaderIcons=   "frmOptions.frx":0F84
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
         Text            =   "frmOptions.frx":0FB0
         MultiLine       =   -1  'True
         ScrollBars      =   2
         CueBanner       =   "frmOptions.frx":0FD0
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
         Caption         =   "frmOptions.frx":0FF0
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
         Caption         =   "frmOptions.frx":10BC
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdAddOS 
         Height          =   750
         Left            =   120
         TabIndex        =   65
         Top             =   4440
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
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "��������"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdEditOS 
         Height          =   750
         Left            =   2160
         TabIndex        =   66
         Top             =   4455
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
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "��������"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdDelOS 
         Height          =   750
         Left            =   4200
         TabIndex        =   67
         Top             =   4455
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
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "�������"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
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
         Caption         =   "��������� ��������� HWID (����������� ����� "";"") �� ��������� (�������������� ����� ""*""):"
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
      Caption         =   "����������"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
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
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2700
         Visible         =   0   'False
         Width           =   495
      End
      Begin prjDIADBS.CheckBoxW chkFutureButton 
         Height          =   210
         Left            =   4680
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   2940
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
         Caption         =   "frmOptions.frx":1154
         Transparent     =   -1  'True
      End
      Begin VB.ComboBox cmbImageMain 
         Height          =   315
         ItemData        =   "frmOptions.frx":1174
         Left            =   615
         List            =   "frmOptions.frx":1181
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   4605
         Width           =   3000
      End
      Begin VB.ComboBox cmbImageStatus 
         Height          =   315
         ItemData        =   "frmOptions.frx":11B9
         Left            =   3960
         List            =   "frmOptions.frx":11C6
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   4605
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
         Caption         =   "frmOptions.frx":11FE
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
         Caption         =   "frmOptions.frx":1272
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
         Left            =   5685
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
         Left            =   5685
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
         Caption         =   "frmOptions.frx":12D8
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorButton 
         Height          =   795
         Left            =   420
         TabIndex        =   73
         Top             =   3075
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
         Caption         =   "���������� ���� � ����� ������ ������"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDIADBS.CheckBoxW chkButtonDisable 
         Height          =   270
         Left            =   3960
         TabIndex        =   74
         Top             =   3480
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":133C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlXpButton cmdFutureButton 
         Height          =   615
         Left            =   4620
         TabIndex        =   75
         Top             =   2760
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
         Caption         =   "������ ������ ���������"
         ButtonStyle     =   3
         PictureWidth    =   32
         PictureHeight   =   32
         ShowFocusRect   =   0   'False
         XPColor_Pressed =   15116940
         XPColor_Hover   =   4692449
         TextColor       =   0
         MenuExist       =   -1  'True
      End
      Begin prjDIADBS.LabelW lblTheme 
         Height          =   225
         Left            =   360
         TabIndex        =   104
         Top             =   4020
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
         Caption         =   "����� ���������� ��������� (��������� �������� ������, � ������ ������� ������)"
      End
      Begin prjDIADBS.LabelW lblImageStatus 
         Height          =   255
         Left            =   3960
         TabIndex        =   105
         Top             =   4305
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
         Caption         =   "������ ��� ������ �������"
      End
      Begin prjDIADBS.LabelW lblImageMain 
         Height          =   255
         Left            =   615
         TabIndex        =   106
         Top             =   4305
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
         Caption         =   "�������� ��������"
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
         Caption         =   "������:"
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
         Caption         =   "������:"
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
         Caption         =   "������:"
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
         Caption         =   "�����:"
         AutoSize        =   -1  'True
      End
      Begin prjDIADBS.LabelW lblButton2BtnT 
         Height          =   195
         Left            =   3525
         TabIndex        =   111
         Top             =   1965
         Width           =   1845
         _ExtentX        =   3413
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
         Caption         =   "�������� �� ���������:"
         AutoSize        =   -1  'True
      End
      Begin prjDIADBS.LabelW lblButton2BtnL 
         Height          =   195
         Left            =   3525
         TabIndex        =   112
         Top             =   1605
         Width           =   2010
         _ExtentX        =   3678
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
         Caption         =   "�������� �� �����������:"
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
         Caption         =   "�������� ������"
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
         Caption         =   "������:"
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
         Caption         =   "������:"
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
         Caption         =   "������� ��������� ����"
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
      Caption         =   "���������� 2"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
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
         Caption         =   "frmOptions.frx":13B2
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
         Caption         =   "frmOptions.frx":1474
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
         Caption         =   "frmOptions.frx":1522
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorTabOS 
         Height          =   795
         Left            =   390
         TabIndex        =   71
         Top             =   2070
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
         Caption         =   "���������� ���� � ����� ������ ��������"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorTabDrivers 
         Height          =   795
         Left            =   390
         TabIndex        =   72
         Top             =   3360
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
         Caption         =   "���������� ���� � ����� ������ ��������"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
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
         Caption         =   "���-�� ������� �� �� ���� ������: "
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
         Caption         =   "TabControl - �������������� ��"
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
         Caption         =   "TabControl 2 - ������ ���������"
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
      Caption         =   "��������� ������� DPInst"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
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
         Text            =   "frmOptions.frx":159E
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":15BE
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
         Caption         =   "frmOptions.frx":15DE
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
         Caption         =   "frmOptions.frx":1612
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
         Caption         =   "frmOptions.frx":1664
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
         Caption         =   "frmOptions.frx":16B4
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
         Caption         =   "frmOptions.frx":1706
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
         Caption         =   "frmOptions.frx":1742
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
         Caption         =   "frmOptions.frx":177A
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
         Caption         =   "�������� ��������� ������� "
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
         Caption         =   "��������  ���������"
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
         Caption         =   "��������"
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
         Caption         =   $"frmOptions.frx":17B2
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
      Caption         =   "���������� �����"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
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
         Text            =   "frmOptions.frx":18B0
         CueBanner       =   "frmOptions.frx":18F4
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
         Text            =   "frmOptions.frx":1914
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1940
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
         Text            =   "frmOptions.frx":1960
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":198E
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
         Text            =   "frmOptions.frx":19AE
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":19DC
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
         Text            =   "frmOptions.frx":19FC
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1A2E
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
         Text            =   "frmOptions.frx":1A4E
         Locked          =   -1  'True
         CueBanner       =   "frmOptions.frx":1A7E
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
         Caption         =   "frmOptions.frx":1A9E
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlUcPickBox ucDebugLogPath 
         Height          =   315
         Left            =   480
         TabIndex        =   84
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
         Caption         =   "frmOptions.frx":1AEE
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
         Caption         =   "frmOptions.frx":1B6E
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
         Caption         =   "���� � ����� �������� ���-�����"
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
         Caption         =   "����������� ������������ �������, � ���� x32[64]"
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
         Caption         =   "������ ������������ ������� � ���� wnt5[6]"
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
         Caption         =   "������ ����������/����������� �����"
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
         Caption         =   "��������"
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
         Caption         =   "��������  ���������"
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
         Caption         =   "������� ��� ����������, ��� ��������� ��������"
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
         Caption         =   "��������� ���������������� ��� ����� ���-�����:"
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
         Caption         =   "������� ��� �������� log-������:"
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
         Caption         =   "��������� ����������� ������"
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
         Caption         =   "������� ��� �������� log-������:"
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

Public WithEvents lvOptions As cListView
Attribute lvOptions.VB_VarHelpID = -1

'ItemOptions1=�������� ���������
'ItemOptions2=�������������� ��
'ItemOptions3=������� �������
'ItemOptions4=��������������� �������
'ItemOptions5=���������� ���������
'ItemOptions6=��������� ������� DPInst
'ItemOptions8=�������� ��������� 2
'ItemOptions9=���������� ��������� 2
'ItemOptions10=���������� �����
Private strItemOptions1     As String
Private strItemOptions2     As String
Private strItemOptions3     As String
Private strItemOptions4     As String
Private strItemOptions5     As String
Private strItemOptions6     As String
Private strItemOptions8     As String
Private strItemOptions9     As String
Private strItemOptions10    As String
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
'! Procedure   (�������)   :   Sub FontCharsetChange
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub FontCharsetChange()

    ' ���������� �����
    With Me.Font
        .Name = strOtherForm_FontName
        .Size = lngOtherForm_FontSize
        .Charset = lngDialog_Charset
    End With

    frDebug.Font.Charset = lngDialog_Charset
    frDesign.Font.Charset = lngDialog_Charset
    frDesign2.Font.Charset = lngDialog_Charset
    frDpInstParam.Font.Charset = lngDialog_Charset
    frMain.Font.Charset = lngDialog_Charset
    frMain2.Font.Charset = lngDialog_Charset
    frMainTools.Font.Charset = lngDialog_Charset
    frOptions.Font.Charset = lngDialog_Charset
    frOS.Font.Charset = lngDialog_Charset
    frOther.Font.Charset = lngDialog_Charset
    frOtherTools.Font.Charset = lngDialog_Charset
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ChangeButtonProperties
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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

    SetButtonProperties cmdFutureButton
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkButtonDisable_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkButtonDisable_Click()
    cmdFutureButton.EnabledCtrl = chkButtonDisable.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkButtonTextUpCase_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkButtonTextUpCase_Click()
    ChangeButtonProperties
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkDebug_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub chkDebugLog2AppPath_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkDebugLog2AppPath_Click()
    DebugCtlEnableLog2App Not chkDebugLog2AppPath.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkForceIfDriverIsNotBetter_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkForceIfDriverIsNotBetter_Click()
    mbDpInstForceIfDriverIsNotBetter = chkForceIfDriverIsNotBetter.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkFormMaximaze_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkFormMaximaze_Click()

    If chkFormMaximaze.Value Then
        chkFormSizeSave.Value = False
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkFormSizeSave_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkFormSizeSave_Click()

    If chkFormSizeSave.Value Then
        chkFormMaximaze.Value = False
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkLegacyMode_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkLegacyMode_Click()
    mbDpInstLegacyMode = chkLegacyMode.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkPromptIfDriverIsNotBetter_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkPromptIfDriverIsNotBetter_Click()
    mbDpInstPromptIfDriverIsNotBetter = chkPromptIfDriverIsNotBetter.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkQuietInstall_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkQuietInstall_Click()
    mbDpInstQuietInstall = chkQuietInstall.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkScanHardware_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkScanHardware_Click()
    mbDpInstScanHardware = chkScanHardware.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkSuppressAddRemovePrograms_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkSuppressAddRemovePrograms_Click()
    mbDpInstSuppressAddRemovePrograms = chkSuppressAddRemovePrograms.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkSuppressWizard_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkSuppressWizard_Click()
    mbDpInstSuppressWizard = chkSuppressWizard.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkTabBlock_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkTabBlock_Click()
    Tab2CtlEnable chkTabBlock.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkTabHide_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkTabHide_Click()
    TabCtlEnable Not chkTabHide.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkTempPath_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkTempPath_Click()
    TempCtlEnable chkTempPath.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkUpdate_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkUpdate_Click()
    UpdateCtlEnable chkUpdate.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmbImageMain_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmbImageMain_Click()

    If PathExists(strPathImageMain & cmbImageMain.Text) = False Then
        cmbImageMain.BackColor = vbRed
    Else
        cmbImageMain.BackColor = &H80000005
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmbImageMain_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmbImageMain_GotFocus()
    HighlightActiveControl Me, cmbImageMain, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmbImageMain_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub cmbImageStatus_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub cmbImageStatus_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmbImageStatus_GotFocus()
    HighlightActiveControl Me, cmbImageStatus, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmbImageStatus_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmbImageStatus_LostFocus()

    If PathExists(strPathImageStatusButton & cmbImageStatus.Text) = False Then
        cmbImageStatus.BackColor = vbRed
    Else
        cmbImageStatus.BackColor = &H80000005
    End If

    HighlightActiveControl Me, cmbImageStatus, False
End Sub

'! -----------------------------------------------------------
'!  �������     :  cmdAddOS_Click
'!  ����������  :
'!  ��������    :  ������ ���������� ��
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdAddOS_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdAddOS_Click()
    mbAddInList = True
    frmOSEdit.Show vbModal, Me
End Sub

'! -----------------------------------------------------------
'!  �������     :  cmdAddUtil_Click
'!  ����������  :
'!  ��������    :  ������ ���������� �������
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdAddUtil_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdAddUtil_Click()
    mbAddInList = True
    frmUtilsEdit.Show vbModal, Me
End Sub

'! -----------------------------------------------------------
'!  �������     :  cmdDelOS_Click
'!  ����������  :
'!  ��������    :  ������ �������� ��
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdDelOS_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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

'! -----------------------------------------------------------
'!  �������     :  cmdDelUtil_Click
'!  ����������  :
'!  ��������    :  ������ �������� �������
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdDelUtil_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub cmdDriverVer_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdDriverVer_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = Kavichki & "http://msdn.microsoft.com/en-us/library/ff547394%28VS.85%29.aspx?ppud=4" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

'! -----------------------------------------------------------
'!  �������     :  cmdEditOS_Click
'!  ����������  :
'!  ��������    :  ������ �������������� ��
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdEditOS_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdEditOS_Click()
    TransferOSData
End Sub

'! -----------------------------------------------------------
'!  �������     :  cmdEditUtil_Click
'!  ����������  :
'!  ��������    :  ������ �������������� �������
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdEditUtil_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdEditUtil_Click()
    TransferUtilsData
End Sub

'! -----------------------------------------------------------
'!  �������     :  cmdExit_Click
'!  ����������  :
'!  ��������    : ������� ������ �����. ����� ��� ����������
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdExit_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    lngShowMessageResult = vbNo
    Me.Hide
    ChangeStatusTextAndDebug cmdExit.Caption
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdFontColorButton_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdFontColorButton_Click()

    With frmFontDialog
        .opt3.Value = True
        .txtFont.Font.Name = strDialog_FontName
        .txtFont.Font.Size = miDialog_FontSize
        .txtFont.Font.Bold = mbDialog_Bold
        .txtFont.Font.Italic = mbDialog_Italic
        .txtFont.Font.Underline = mbDialog_Underline
        .txtFont.Font.Charset = lngDialog_Charset
        .txtFont.ForeColor = lngDialog_Color
        .Show vbModal, Me
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdFontColorTabDrivers_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdFontColorTabDrivers_Click()

    With frmFontDialog
        .opt2.Value = True
        .txtFont.Font.Name = strDialogTab2_FontName
        .txtFont.Font.Size = miDialogTab2_FontSize
        .txtFont.Font.Bold = mbDialogTab2_Bold
        .txtFont.Font.Italic = mbDialogTab2_Italic
        .txtFont.Font.Underline = mbDialogTab2_Underline
        .txtFont.Font.Charset = lngDialog_Charset
        .ForeColor = lngDialogTab2_Color
        .Show vbModal, Me
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdFontColorTabOS_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdFontColorTabOS_Click()

    With frmFontDialog
        .opt1.Value = True
        .txtFont.Font.Name = strDialogTab_FontName
        .txtFont.Font.Size = miDialogTab_FontSize
        .txtFont.Font.Bold = mbDialogTab_Bold
        .txtFont.Font.Italic = mbDialogTab_Italic
        .txtFont.Font.Underline = mbDialogTab_Underline
        .txtFont.Font.Charset = lngDialog_Charset
        .txtFont.ForeColor = lngDialogTab_Color
        .Show vbModal, Me
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdForceIfDriverIsNotBetter_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub cmdLegacyMode_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdLegacyMode_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = Kavichki & "http://msdn.microsoft.com/en-us/library/ms794322.aspx" & Kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

'! -----------------------------------------------------------
'!  �������     :  cmdOK_Click
'!  ����������  :
'!  ��������    :  ������� ������ ��. ���������� ��������
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdOK_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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

    'Unload Me
    Me.Hide
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdPathDefault_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdPathDefault_Click()
    ucDevCon86Path.Path = "Tools\Devcon\devcon.exe"
    ucDevCon64Path.Path = "Tools\Devcon\devcon64.exe"
    ucDevCon86Pathw2k.Path = "Tools\Devcon\devconw2k.exe"
    '������ DPInst
    ucDPInst86Path.Path = "Tools\DPInst\DPInst.exe"
    ucDPInst64Path.Path = "Tools\DPInst\DPInst64.exe"
    '������ Arc
    ucArchPath.Path = "Tools\Arc\7za.exe"
    ucCmdDevconPath.Path = "Tools\Devcon\devcon_c.cmd"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdPromptIfDriverIsNotBetter_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub cmdQuietInstall_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub cmdScanHardware_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub cmdSuppressAddRemovePrograms_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub cmdSuppressWizard_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub DebugCtlEnable
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub DebugCtlEnable(ByVal mbEnable As Boolean)
    chkDebugTime2File.Enabled = mbEnable
    txtDebugLogName.Enabled = mbEnable
    ucDebugLogPath.Enabled = mbEnable
    chkDebugLog2AppPath.Enabled = mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub DebugCtlEnableLog2App
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub DebugCtlEnableLog2App(ByVal mbEnable As Boolean)
    ucDebugLogPath.Enabled = mbEnable
End Sub

'! -----------------------------------------------------------
'!  �������     :  Form_KeyDown
'!  ����������  :  KeyCode As Integer, Shift As Integer
'!  ��������    :  ��������� ������� ������ ���������� ������� �� �����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_KeyDown
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        If MsgBox(strMessages(37), vbQuestion + vbYesNo, strProductName) = vbYes Then
            cmdExit_Click
        End If
    End If

End Sub

'! -----------------------------------------------------------
'!  �������     :  Form_Load
'!  ����������  :
'!  ��������    :  �������� �����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_Load
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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

    'Top
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
    'Left
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
    ' ������������� ����������� ��������
    txtTabPerRowCount.Min = 2
    txtFormHeight.Min = MainFormHeightMin
    txtFormWidth.Min = MainFormWidthMin
    txtButtonHeight.Min = ButtonHeightMin
    txtButtonWidth.Min = ButtonWidthMin
    ' ������������� �������� ������ � ������� �������� ������
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
    FormLoadAction
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub FormLoadAction
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub FormLoadAction()

    ' ����������z ����������
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' ���������� �����
        FontCharsetChange
    End If

    ' ��������� ������ �����
    tvOptionsLoad
    ' ��������� �����
    ReadOptions
    ' ���������� ����� ������ � �����
    Set cmdFutureButton.Picture = imgOK.Picture
    SetButtonProperties cmdFutureButton
    DoEvents
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_QueryUnload
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   Cancel (Integer)
'                              UnloadMode (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' ��������� �� ������ ����� � ������ ����������
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
        ChangeStatusTextAndDebug cmdExit.Caption
    Else
        Set lvOptions = Nothing
        Set frmOptions = Nothing
    End If

End Sub

'! -----------------------------------------------------------
'!  �������     :  Form_Resize
'!  ����������  :
'!  ��������    :  ��������� �������� �����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_Resize
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub InitializeObjectProperties
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub InitializeObjectProperties()

    With cmdFutureButton
        .CheckExist = True
    End With

    chkFutureButton.ZOrder 0
    ' ��������� ������ � ������
    ChangeButtonProperties
End Sub

'! -----------------------------------------------------------
'!  �������     :  LoadList_OS
'!  ����������  :
'!  ��������    :  ���������� ���c�� ��
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadList_OS
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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

    'LVOS
    LastIdOS = lngOSCount
End Sub

'! -----------------------------------------------------------
'!  �������     :  LoadList_Utils
'!  ����������  :
'!  ��������    :  ���������� ����� ������
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadList_Utils
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub LoadListCombo
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   cmbName (ComboBox)
'                              strImagePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadListCombo(cmbName As ComboBox, strImagePath As String)

    Dim strListFolderTemp() As String
    Dim i                   As Integer

    strListFolderTemp = GetAllFolderInFolder(strImagePath)

    With cmbName
        .Clear

        For i = LBound(strListFolderTemp) To UBound(strListFolderTemp)
            .AddItem strListFolderTemp(i), i
        Next

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Localise
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal StrPathFile As String)

    Dim strFormNameMain As String

    strFormNameMain = frmMain.Name
    ' ���������� ����� ��������� (��������� ������ �� �� ��� ������� �� �������������� ������)
    FontCharsetChange
    ' �������� �����
    Me.Caption = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
    frOptions.Caption = LocaliseString(StrPathFile, strFormName, "frOptions", frOptions.Caption)
    ' �������� �������
    optRezim_Intellect.Caption = LocaliseString(StrPathFile, strFormNameMain, "RezimIntellect", optRezim_Intellect.Caption)
    optRezim_Ust.Caption = LocaliseString(StrPathFile, strFormNameMain, "RezimUst", optRezim_Ust.Caption)
    optRezim_Upd.Caption = LocaliseString(StrPathFile, strFormNameMain, "RezimUpd", optRezim_Upd.Caption)
    optRezim_Intellect.ToolTipText = LocaliseString(StrPathFile, strFormNameMain, "RezimIntellectTip", optRezim_Intellect.ToolTipText)
    optRezim_Ust.ToolTipText = LocaliseString(StrPathFile, strFormNameMain, "RezimUstTip", optRezim_Ust.ToolTipText)
    optRezim_Upd.ToolTipText = LocaliseString(StrPathFile, strFormNameMain, "RezimUpdTip", optRezim_Upd.ToolTipText)
    strItemOptions1 = LocaliseString(StrPathFile, strFormName, "ItemOptions1", "�������� ���������")
    strItemOptions2 = LocaliseString(StrPathFile, strFormName, "ItemOptions2", "�������������� ��")
    strItemOptions3 = LocaliseString(StrPathFile, strFormName, "ItemOptions3", "������� �������")
    strItemOptions4 = LocaliseString(StrPathFile, strFormName, "ItemOptions4", "��������������� �������")
    strItemOptions5 = LocaliseString(StrPathFile, strFormName, "ItemOptions5", "���������� ���������")
    strItemOptions6 = LocaliseString(StrPathFile, strFormName, "ItemOptions6", "��������� ������� DPInst")
    strItemOptions8 = LocaliseString(StrPathFile, strFormName, "ItemOptions8", "�������� ��������� 2")
    strItemOptions9 = LocaliseString(StrPathFile, strFormName, "ItemOptions9", "���������� ��������� 2")
    strItemOptions10 = LocaliseString(StrPathFile, strFormName, "ItemOptions10", "���������� �����")
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
    strTableOSHeader1 = LocaliseString(StrPathFile, strFormName, "TableOSHeader1", "������")
    strTableOSHeader2 = LocaliseString(StrPathFile, strFormName, "TableOSHeader2", "������������")
    strTableOSHeader3 = LocaliseString(StrPathFile, strFormName, "TableOSHeader3", "������ ���������")
    strTableOSHeader4 = LocaliseString(StrPathFile, strFormName, "TableOSHeader4", "���� ������")
    strTableOSHeader5 = LocaliseString(StrPathFile, strFormName, "TableOSHeader5", "x64")
    strTableOSHeader6 = LocaliseString(StrPathFile, strFormName, "TableOSHeader6", "PhysX")
    strTableOSHeader7 = LocaliseString(StrPathFile, strFormName, "TableOSHeader7", "Lang")
    strTableOSHeader8 = LocaliseString(StrPathFile, strFormName, "TableOSHeader8", "ExludeFiles")
    strTableOSHeader9 = LocaliseString(StrPathFile, strFormName, "TableOSHeader9", "ExludeFiles")
    strTableUtilHeader1 = LocaliseString(StrPathFile, strFormName, "TableUtilHeader1", "������������")
    strTableUtilHeader2 = LocaliseString(StrPathFile, strFormName, "TableUtilHeader2", "����")
    strTableUtilHeader3 = LocaliseString(StrPathFile, strFormName, "TableUtilHeader3", "��������")
    strTableUtilHeader4 = LocaliseString(StrPathFile, strFormName, "TableUtilHeader4", "���� x64")
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
End Sub

'! -----------------------------------------------------------
'!  �������     :  lvOptions_ItemChanged
'!  ����������  :
'!  ��������    :  ��� ������ ����� ���������� ����������� ��������������� ����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub lvOptions_ItemChanged
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   iIndex (Long)
'!--------------------------------------------------------------------------------
Private Sub lvOptions_ItemChanged(ByVal iIndex As Long)

    'ItemOptions1=�������� ���������
    'ItemOptions8=�������� ��������� 2
    'ItemOptions2=�������������� ��
    'ItemOptions3=������� �������
    'ItemOptions4=��������������� �������
    'ItemOptions5=���������� ���������
    'ItemOptions9=���������� ��������� 2
    'ItemOptions6=��������� ������� DPInst
    'ItemOptions10=���������� �����
    Select Case iIndex

        Case 0 'strItemOptions1
            frMain.ZOrder 0

        Case 1 ' strItemOptions8
            frMain2.ZOrder 0

        Case 2 'strItemOptions2
            frOS.ZOrder 0
            txtExcludeHWID.SetFocus

        Case 3 'strItemOptions3
            frMainTools.ZOrder 0
            ucDevCon86Path.SetFocus

        Case 4 ' strItemOptions4
            frOtherTools.ZOrder 0

        Case 5 'strItemOptions5
            frDesign.ZOrder 0
            cmbImageMain.SetFocus

        Case 6 ' strItemOptions9
            frDesign2.ZOrder 0

        Case 7 'strItemOptions6
            frDpInstParam.ZOrder 0
            txtCmdStringDPInst.SetFocus

        Case 8 ' strItemOptions10
            frDebug.ZOrder 0
            txtDebugLogName.SetFocus

        Case Else
            frOther.ZOrder 0
    End Select

End Sub

'! -----------------------------------------------------------
'!  �������     :  lvOS_ItemDblClick
'!  ����������  :
'!  ��������    :  ������� ���� �� �������� ������ �������� ����� ��������������
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub lvOS_ItemDblClick
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   Item (LvwListItem)
'                              Button (Integer)
'!--------------------------------------------------------------------------------
Private Sub lvOS_ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)
    TransferOSData
End Sub

'! -----------------------------------------------------------
'!  �������     :  lvUtils_ItemDblClick
'!  ����������  :
'!  ��������    :  ������� ���� �� �������� ������ �������� ����� ��������������
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub lvUtils_ItemDblClick
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   Item (LvwListItem)
'                              Button (Integer)
'!--------------------------------------------------------------------------------
Private Sub lvUtils_ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)
    TransferUtilsData
End Sub

'! -----------------------------------------------------------
'!  �������     :  ReadOptions
'!  ����������  :
'!  ��������    :  ������ ��������� ��������� � ��������� ����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ReadOptions
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ReadOptions()
    ' ��������� ������ ��
    LoadList_OS
    ' ��������� ������ ������
    LoadList_Utils
    ' ��������� ���������
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

    ' ����� ��� ������
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
    txtFormHeight.Value = MainFormHeight
    txtFormWidth.Value = MainFormWidth
    txtExcludeHWID.Text = strExcludeHWID
    'Buttons
    txtButtonWidth.Value = miButtonWidth
    txtButtonHeight.Value = miButtonHeight
    txtButtonLeft.Value = miButtonLeft
    txtButtonTop.Value = miButtonTop
    txtButton2BtnL.Value = miBtn2BtnLeft
    txtButton2BtnT.Value = miBtn2BtnTop
    chkButtonTextUpCase.Value = mbButtonTextUpCase
    txtTabPerRowCount.Value = lngOSCountPerRow

    '���� � ����������
    If mbPatnAbs Then
        '������ Devcon
        ucDevCon86Path.Path = strDevConExePath
        ucDevCon64Path.Path = strDevConExePath64
        ucDevCon86Pathw2k.Path = strDevConExePathW2k
        '������ DPInst
        ucDPInst86Path.Path = strDPInstExePath86
        ucDPInst64Path.Path = strDPInstExePath64
        '������ Arc
        ucArchPath.Path = strArh7zExePATH
        ucCmdDevconPath.Path = strDevconCmdPath
    Else
        '������ Devcon
        ucDevCon86Path.Path = Replace$(strDevConExePath, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucDevCon64Path.Path = Replace$(strDevConExePath64, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucDevCon86Pathw2k.Path = Replace$(strDevConExePathW2k, strAppPathBackSL, vbNullString, , , vbTextCompare)
        '������ DPInst
        ucDPInst86Path.Path = Replace$(strDPInstExePath86, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucDPInst64Path.Path = Replace$(strDPInstExePath64, strAppPathBackSL, vbNullString, , , vbTextCompare)
        '������ Arc
        ucArchPath.Path = Replace$(strArh7zExePATH, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucCmdDevconPath.Path = Replace$(strDevconCmdPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
    End If

    ' ��������� DpInst
    chkLegacyMode.Value = mbDpInstLegacyMode
    chkPromptIfDriverIsNotBetter.Value = mbDpInstPromptIfDriverIsNotBetter
    chkForceIfDriverIsNotBetter.Value = mbDpInstForceIfDriverIsNotBetter
    chkSuppressAddRemovePrograms.Value = mbDpInstSuppressAddRemovePrograms
    chkSuppressWizard.Value = mbDpInstSuppressWizard
    chkQuietInstall.Value = mbDpInstQuietInstall
    chkScanHardware.Value = mbDpInstScanHardware
    ' ������ ���������
    txtCmdStringDPInst = CollectCmdString
    chkLoadFinishFile.Value = mbLoadFinishFile
    'chkReadClasses.value = mbReadClasses
    chkReadDPName.Value = mbReadDPName
    chkConvertDPName.Value = mbConvertDPName
    ' �������� ������ ������
    LoadListCombo cmbImageMain, strPathImageMain
    LoadListCombo cmbImageStatus, strPathImageStatusButton
    LoadIconImage2Object imgOK, "BTN_OK", strPathImageStatusButtonWork
    cmbImageMain.Text = strImageMainName
    cmbImageStatus.Text = strImageStatusButtonName
    ' ��������� ������ ���������
    optCompareByDate.Value = mbCompareDrvVerByDate
    optCompareByVersion.Value = Not mbCompareDrvVerByDate
    ' ���������� ���� � ������� dd/mm/yyyy
    chkDateFormatRus.Value = mbDateFormatRus
    '����� ����� ��������� ��� ������
    chkSearchOnStart.Value = mbSearchOnStart

    With txtPauseAfterSearch
        .Min = 0
        .Increment = 1
        .Value = lngPauseAfterSearch
    End With

    ' ��������� ���������� ���������
    DebugCtlEnable chkDebug.Value
    DebugCtlEnableLog2App Not chkDebugLog2AppPath.Value
    TempCtlEnable chkTempPath.Value
    UpdateCtlEnable chkUpdate.Value
    TabCtlEnable Not chkTabHide.Value
    Tab2CtlEnable chkTabBlock.Value
    ' ������������� ���������� ��� ��������� ������ � �����
    InitializeObjectProperties
End Sub

'! -----------------------------------------------------------
'!  �������     :  SaveOptions
'!  ����������  :
'!  ��������    :  ���������� �������� � ���-����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub SaveOptions
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
    '***************** ������ �������� ****************
    '**************************************************
    ' ������ MAIN
    '�������� TEMP ��� ������
    IniWriteStrPrivate "Main", "DelTmpAfterClose", CStr(Abs(chkRemoveTemp.Value)), strSysIniTemp
    ' ��������������
    IniWriteStrPrivate "Main", "UpdateCheck", CStr(Abs(chkUpdate.Value)), strSysIniTemp
    ' �������������� Beta
    IniWriteStrPrivate "Main", "UpdateCheckBeta", CStr(Abs(chkUpdateBeta.Value)), strSysIniTemp

    ' ����� �������
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
    ' ����� ����� ��������� ��� ������
    IniWriteStrPrivate "Main", "SearchOnStart", CStr(Abs(chkSearchOnStart.Value)), strSysIniTemp
    IniWriteStrPrivate "Main", "PauseAfterSearch", txtPauseAfterSearch.Value, strSysIniTemp
    ' �������� ����� ��������������
    IniWriteStrPrivate "Main", "CreateRestorePoint", CStr(Abs(chkCreateRP.Value)), strSysIniTemp

    If mbLoadIniTmpAfterRestart Then
        IniWriteStrPrivate "Main", "LoadIniTmpAfterRestart", 1, strSysIniTemp
    End If

    ' ������ Debug
    IniWriteStrPrivate "Debug", "DebugEnable", CStr(Abs(chkDebug.Value)), strSysIniTemp
    ' ������� �������:
    IniWriteStrPrivate "Debug", "CleenHistory", 1, strSysIniTemp
    ' ���� �� ���-�����
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
    '������ Devcon
    IniWriteStrPrivate "Devcon", "PathExe", ucDevCon86Path.Path, strSysIniTemp
    IniWriteStrPrivate "Devcon", "PathExe64", ucDevCon64Path.Path, strSysIniTemp
    IniWriteStrPrivate "Devcon", "PathExeW2K", ucDevCon86Pathw2k.Path, strSysIniTemp
    IniWriteStrPrivate "Devcon", "CollectHwidsCmd", ucCmdDevconPath.Path, strSysIniTemp
    '������ DPInst
    IniWriteStrPrivate "DPInst", "PathExe", ucDPInst86Path.Path, strSysIniTemp
    IniWriteStrPrivate "DPInst", "PathExe64", ucDPInst64Path.Path, strSysIniTemp
    IniWriteStrPrivate "DPInst", "LegacyMode", CStr(Abs(chkLegacyMode.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "PromptIfDriverIsNotBetter", CStr(Abs(chkPromptIfDriverIsNotBetter.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "ForceIfDriverIsNotBetter", CStr(Abs(chkForceIfDriverIsNotBetter.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "SuppressAddRemovePrograms", CStr(Abs(chkSuppressAddRemovePrograms.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "SuppressWizard", CStr(Abs(chkSuppressWizard.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "QuietInstall", CStr(Abs(chkQuietInstall.Value)), strSysIniTemp
    IniWriteStrPrivate "DPInst", "ScanHardware", CStr(Abs(chkScanHardware.Value)), strSysIniTemp
    '������ Arc
    IniWriteStrPrivate "Arc", "PathExe", ucArchPath.Path, strSysIniTemp
    '������ OS
    '����� ��
    lngOSCountNew = lvOS.ListItems.Count
    IniWriteStrPrivate "OS", "OSCount", CStr(lngOSCountNew), strSysIniTemp
    ' ����������� ������� �����
    IniWriteStrPrivate "OS", "Recursion", CStr(Abs(chkRecursion.Value)), strSysIniTemp
    ' ���-�� ����� �� ������
    IniWriteStrPrivate "OS", "OSCountPerRow", txtTabPerRowCount.Value, strSysIniTemp
    ' ����������� ������ �������
    IniWriteStrPrivate "OS", "TabBlock", CStr(Abs(chkTabBlock.Value)), strSysIniTemp
    ' �������� ������ �������
    IniWriteStrPrivate "OS", "TabHide", CStr(Abs(chkTabHide.Value)), strSysIniTemp
    ' ������������ ����� Finish
    IniWriteStrPrivate "OS", "LoadFinishFile", CStr(Abs(chkLoadFinishFile.Value)), strSysIniTemp
    ' ��������� ����� ������ ��������� �� Finish
    'IniWriteStrPrivate "OS", "ReadClasses", CStr(Abs(chkReadClasses.value)), strSysIniTemp
    ' ��������� ����� ������ ��������� �� Finish
    IniWriteStrPrivate "OS", "ReadDPName", CStr(Abs(chkReadDPName.Value)), strSysIniTemp
    ' ��������� ����� ������ ��������� �� Finish
    IniWriteStrPrivate "OS", "ConvertDPName", CStr(Abs(chkConvertDPName.Value)), strSysIniTemp
    IniWriteStrPrivate "OS", "ExcludeHWID", txtExcludeHWID.Text, strSysIniTemp
    ' ��������� ������ ���������
    IniWriteStrPrivate "OS", "CompareDrvVerByDate", CStr(Abs(optCompareByDate.Value)), strSysIniTemp
    IniWriteStrPrivate "OS", "DateFormatRus", CStr(Abs(chkDateFormatRus.Value)), strSysIniTemp
    ' �������������� �������
    IniWriteStrPrivate "OS", "LoadUnSupportedOS", CStr(Abs(chkLoadUnSupportedOS.Value)), strSysIniTemp
    ' ������������ ����������� HWID
    IniWriteStrPrivate "OS", "CompatiblesHWID", CStr(Abs(chkCompatiblesHWID.Value)), strSysIniTemp

    '�������� � ����� ��������� ��
    For cnt = 1 To lngOSCountNew

        '������ OS_N
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

    '������ Utils
    '����� ������
    lngUtilsCountNew = lvUtils.ListItems.Count
    IniWriteStrPrivate "Utils", "UtilsCount", CStr(lngUtilsCountNew), strSysIniTemp

    '�������� � ����� ��������� �������
    For cnt = 1 To lngUtilsCountNew

        '������ Utils_N
        With lvUtils.ListItems(cnt)
            IniWriteStrPrivate "Utils_" & cnt, "Name", .Text, strSysIniTemp
            IniWriteStrPrivate "Utils_" & cnt, "Path", .SubItems(1), strSysIniTemp
            IniWriteStrPrivate "Utils_" & cnt, "Path64", .SubItems(2), strSysIniTemp
            IniWriteStrPrivate "Utils_" & cnt, "Params", .SubItems(3), strSysIniTemp
        End With

    Next

    '������ MainForm
    IniWriteStrPrivate "MainForm", "Width", txtFormWidth.Value, strSysIniTemp
    IniWriteStrPrivate "MainForm", "Height", txtFormHeight.Value, strSysIniTemp
    IniWriteStrPrivate "MainForm", "StartMaximazed", CStr(Abs(chkFormMaximaze.Value)), strSysIniTemp
    mbSaveSizeOnExit = chkFormSizeSave.Value
    IniWriteStrPrivate "MainForm", "SaveSizeOnExit", CStr(Abs(chkFormSizeSave.Value)), strSysIniTemp
    IniWriteStrPrivate "MainForm", "HighlightColor", CStr(glHighlightColor), strSysIniTemp
    '������ Buttons
    IniWriteStrPrivate "Button", "Width", txtButtonWidth.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Height", txtButtonHeight.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Left", txtButtonLeft.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Top", txtButtonTop.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Btn2BtnLeft", txtButton2BtnL.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "Btn2BtnTop", txtButton2BtnT.Value, strSysIniTemp
    IniWriteStrPrivate "Button", "TextUpCase", CStr(Abs(chkButtonTextUpCase.Value)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontName", strDialog_FontName, strSysIniTemp
    IniWriteStrPrivate "Button", "FontSize", CStr(miDialog_FontSize), strSysIniTemp
    IniWriteStrPrivate "Button", "FontUnderline", CStr(Abs(mbDialog_Underline)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontStrikethru", CStr(Abs(mbDialog_Strikethru)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontItalic", CStr(Abs(mbDialog_Italic)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontBold", CStr(Abs(mbDialog_Bold)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontColor", CStr(cmdFutureButton.TextColor), strSysIniTemp
    IniWriteStrPrivate "Button", "IconStatusSkin", cmbImageStatus.Text, strSysIniTemp
    '������ Tab
    IniWriteStrPrivate "Tab", "FontName", strDialogTab_FontName, strSysIniTemp
    IniWriteStrPrivate "Tab", "FontSize", CStr(miDialogTab_FontSize), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontUnderline", CStr(Abs(mbDialogTab_Underline)), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontStrikethru", CStr(Abs(mbDialogTab_Strikethru)), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontItalic", CStr(Abs(mbDialogTab_Italic)), strSysIniTemp
    IniWriteStrPrivate "Tab", "FontBold", CStr(Abs(mbDialogTab_Bold)), strSysIniTemp
    '������ Tab2
    IniWriteStrPrivate "Tab2", "FontName", strDialogTab2_FontName, strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontSize", CStr(miDialogTab2_FontSize), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontUnderline", CStr(Abs(mbDialogTab2_Underline)), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontStrikethru", CStr(Abs(mbDialogTab2_Strikethru)), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontItalic", CStr(Abs(mbDialogTab2_Italic)), strSysIniTemp
    IniWriteStrPrivate "Tab2", "FontBold", CStr(Abs(mbDialogTab2_Bold)), strSysIniTemp
    '������ "NotebookVendor"
    IniWriteStrPrivate "NotebookVendor", "FilterCount", UBound(arrNotebookFilterList), strSysIniTemp

    For cnt = 0 To UBound(arrNotebookFilterList) - 1
        IniWriteStrPrivate "NotebookVendor", "Filter_" & cnt + 1, arrNotebookFilterList(cnt), strSysIniTemp
    Next

    ' �������� Ini ���� � ������������ ����
    NormIniFile strSysIniTemp
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Tab2CtlEnable
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub Tab2CtlEnable(ByVal mbEnable As Boolean)
    chkTabHide.Enabled = mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub TabCtlEnable
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub TabCtlEnable(ByVal mbEnable As Boolean)
    chkTabBlock.Enabled = mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub TempCtlEnable
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub TempCtlEnable(ByVal mbEnable As Boolean)
    ucTempPath.Enabled = mbEnable
End Sub

'! -----------------------------------------------------------
'!  �������     :  TransferOSData
'!  ����������  :
'!  ��������    :  �������� ���������� �� �� ����� � ����� ��������������
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub TransferOSData
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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

'! -----------------------------------------------------------
'!  �������     :  TransferUtilsData
'!  ����������  :
'!  ��������    :  �������� ���������� ������ �� ����� � ����� ��������������
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub TransferUtilsData
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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

'! -----------------------------------------------------------
'!  �������     :  tvOptionsLoad
'!  ����������  :
'!  ��������    :  ���������� ������ ��������
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub tvOptionsLoad
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub txtButtonHeight_Change
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtButtonHeight_Change()
    ChangeButtonProperties
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtButtonWidth_Change
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtButtonWidth_Change()
    ChangeButtonProperties
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtExcludeHWID_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtExcludeHWID_GotFocus()
    HighlightActiveControl Me, txtExcludeHWID, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtExcludeHWID_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtExcludeHWID_LostFocus()
    HighlightActiveControl Me, txtExcludeHWID, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucArchPath_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub ucArchPath_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucArchPath_GotFocus()
    HighlightActiveControl Me, ucArchPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucArchPath_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucArchPath_LostFocus()
    HighlightActiveControl Me, ucArchPath, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucCmdDevconPath_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub ucCmdDevconPath_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucCmdDevconPath_GotFocus()
    HighlightActiveControl Me, ucCmdDevconPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucCmdDevconPath_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucCmdDevconPath_LostFocus()
    HighlightActiveControl Me, ucCmdDevconPath, False
End Sub

'! -----------------------------------------------------------
'!  �������     :  ucDebugLogPath_Click
'!  ����������  :
'!  ��������    :  ����� �������� ��� �����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDebugLogPath_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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

'! -----------------------------------------------------------
'!  �������     :  ucDevCon64Path_Click
'!  ����������  :
'!  ��������    :  ����� �������� ��� �����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDevCon64Path_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub ucDevCon64Path_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon64Path_GotFocus()
    HighlightActiveControl Me, ucDevCon64Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDevCon64Path_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon64Path_LostFocus()
    HighlightActiveControl Me, ucDevCon64Path, False
End Sub

'! -----------------------------------------------------------
'!  �������     :  ucDevCon86Path_Click
'!  ����������  :
'!  ��������    :  ����� �������� ��� �����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDevCon86Path_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub ucDevCon86Path_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon86Path_GotFocus()
    HighlightActiveControl Me, ucDevCon86Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDevCon86Path_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon86Path_LostFocus()
    HighlightActiveControl Me, ucDevCon86Path, False
End Sub

'! -----------------------------------------------------------
'!  �������     :  ucDevCon86Pathw2k_Click
'!  ����������  :
'!  ��������    :  ����� �������� ��� �����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDevCon86Pathw2k_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub ucDevCon86Pathw2k_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon86Pathw2k_GotFocus()
    HighlightActiveControl Me, ucDevCon86Pathw2k, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDevCon86Pathw2k_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDevCon86Pathw2k_LostFocus()
    HighlightActiveControl Me, ucDevCon86Pathw2k, False
End Sub

'! -----------------------------------------------------------
'!  �������     :  ucDPInst64Path_Click
'!  ����������  :
'!  ��������    :  ����� �������� ��� �����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDPInst64Path_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub ucDPInst64Path_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst64Path_GotFocus()
    HighlightActiveControl Me, ucDPInst64Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDPInst64Path_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst64Path_LostFocus()
    HighlightActiveControl Me, ucDPInst64Path, False
End Sub

'! -----------------------------------------------------------
'!  �������     :  ucDPInst86Path_Click
'!  ����������  :
'!  ��������    :  ����� �������� ��� �����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDPInst86Path_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub ucDPInst86Path_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst86Path_GotFocus()
    HighlightActiveControl Me, ucDPInst86Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDPInst86Path_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst86Path_LostFocus()
    HighlightActiveControl Me, ucDPInst86Path, False
End Sub

'! -----------------------------------------------------------
'!  �������     :  ucTempPath_Click
'!  ����������  :
'!  ��������    :  ����� �������� ��� �����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucTempPath_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Sub UpdateCtlEnable
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub UpdateCtlEnable(ByVal mbEnable As Boolean)
    chkUpdateBeta.Enabled = mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtDebugLogName_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtDebugLogName_GotFocus()
    HighlightActiveControl Me, txtDebugLogName, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtDebugLogName_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtDebugLogName_LostFocus()
    HighlightActiveControl Me, txtDebugLogName, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtCmdStringDPInst_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtCmdStringDPInst_GotFocus()
    HighlightActiveControl Me, txtCmdStringDPInst, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtCmdStringDPInst_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtCmdStringDPInst_LostFocus()
    HighlightActiveControl Me, txtCmdStringDPInst, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDebugLogPath_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDebugLogPath_GotFocus()
    HighlightActiveControl Me, ucDebugLogPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucDebugLogPath_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucDebugLogPath_LostFocus()
    HighlightActiveControl Me, ucDebugLogPath, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucTempPath_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucTempPath_GotFocus()
    HighlightActiveControl Me, ucTempPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucTempPath_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucTempPath_LostFocus()
    HighlightActiveControl Me, ucTempPath, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub lvUtils_ColumnClick
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   ColumnHeader (LvwColumnHeader)
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
'! Procedure   (�������)   :   Sub lvOS_ColumnClick
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   ColumnHeader (LvwColumnHeader)
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
