VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Помощник установки драйверов (Drivers Installer Assistant)"
   ClientHeight    =   10575
   ClientLeft      =   3540
   ClientTop       =   4215
   ClientWidth     =   11265
   DrawStyle       =   6  'Inside Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   11265
   Begin prjDIADBS.ctlUcStatusBar ctlUcStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   29
      Top             =   9870
      Width           =   11265
      _ExtentX        =   19870
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
      Theme           =   2
   End
   Begin VB.PictureBox pbProgressBar 
      Align           =   2  'Align Bottom
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   11265
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   9345
      Visible         =   0   'False
      Width           =   11265
      Begin prjDIADBS.ctlJCbutton cmdBreakUpdateDB 
         Height          =   385
         Left            =   4200
         TabIndex        =   28
         Top             =   75
         Visible         =   0   'False
         Width           =   3500
         _ExtentX        =   6165
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   12244692
         Caption         =   "Прервать выполнения задания"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         ColorScheme     =   3
      End
      Begin prjDIADBS.ProgressBar ctlProgressBar1 
         Height          =   375
         Left            =   120
         Top             =   60
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   661
         Max             =   1000
         Value           =   100
         Step            =   10
      End
   End
   Begin prjDIADBS.ctlJCFrames frMainPanel 
      Height          =   8835
      Left            =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   15584
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13160660
      FillColor       =   14215660
      Style           =   8
      RoundedCorner   =   0   'False
      Caption         =   ""
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ctlJCFrames frInfo 
         Height          =   1175
         Left            =   75
         Top             =   45
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2064
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14339020
         FillColor       =   14339020
         MoverForm       =   -1  'True
         MoverControle   =   -1  'True
         Collapsar       =   -1  'True
         Collapsado      =   -1  'True
         Style           =   4
         RoundedCorner   =   0   'False
         Caption         =   "Сведения об операционой системе и компьютере..."
         TextBoxHeight   =   20
         Alignment       =   0
         ThemeColor      =   1
         Begin prjDIADBS.LabelW lblPCInfo 
            Height          =   255
            Left            =   75
            TabIndex        =   1
            Top             =   850
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            Caption         =   "Модель PC:"
         End
         Begin prjDIADBS.LabelW lblOsInfo 
            Height          =   255
            Left            =   75
            TabIndex        =   0
            ToolTipText     =   "Starting ""System Information Viewer"""
            Top             =   480
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   15783104
            MousePointer    =   4
            BackStyle       =   0
            Caption         =   "Операционная система:"
         End
      End
      Begin prjDIADBS.ctlJCFrames frTabPanel 
         Height          =   4875
         Left            =   75
         Top             =   3885
         Visible         =   0   'False
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8599
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
         FillColor       =   13160660
         Style           =   8
         Caption         =   "Обнаруженные пакеты драйверов"
         ThemeColor      =   2
         Begin prjDIADBS.CheckBoxW chkPackFiles 
            Height          =   210
            Index           =   0
            Left            =   180
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   4380
            Visible         =   0   'False
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
            Caption         =   "frmMain.frx":000C
            Transparent     =   -1  'True
         End
         Begin prjDIADBS.ctlJCbutton acmdPackFiles 
            Height          =   555
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   4200
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
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
            ShowFocusRect   =   -1  'True
            BackColor       =   14933984
            Caption         =   "Кнопка пакета драйверов"
            CaptionEffects  =   0
            Mode            =   2
            PictureAlign    =   0
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            PicturePushOnHover=   -1  'True
            ColorScheme     =   2
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   4155
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   11160
            _ExtentX        =   19685
            _ExtentY        =   7329
            _Version        =   393216
            Tabs            =   4
            TabHeight       =   520
            WordWrap        =   0   'False
            ShowFocusRect   =   0   'False
            Enabled         =   0   'False
            ForeColor       =   -2147483630
            MouseIcon       =   "frmMain.frx":002C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "OSName_1"
            TabPicture(0)   =   "frmMain.frx":0048
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblNoDPInProgram"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSTab2(0)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "OSName_2"
            TabPicture(1)   =   "frmMain.frx":0064
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
            TabCaption(2)   =   "OSName_3"
            TabPicture(2)   =   "frmMain.frx":0080
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            TabCaption(3)   =   "OSName_4"
            TabPicture(3)   =   "frmMain.frx":009C
            Tab(3).ControlEnabled=   0   'False
            Tab(3).ControlCount=   0
            Begin TabDlg.SSTab SSTab2 
               Height          =   2895
               Index           =   0
               Left            =   0
               TabIndex        =   19
               Top             =   660
               Width           =   10980
               _ExtentX        =   19368
               _ExtentY        =   5106
               _Version        =   393216
               Tabs            =   5
               TabsPerRow      =   5
               TabHeight       =   520
               WordWrap        =   0   'False
               ForeColor       =   -2147483635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Все драйверпаки"
               TabPicture(0)   =   "frmMain.frx":00B8
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "lblNoDP4Mode"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "ctlScrollControl1(0)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).ControlCount=   2
               TabCaption(1)   =   "Доступно обновление"
               TabPicture(1)   =   "frmMain.frx":00D4
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "ctlScrollControlTab1(0)"
               Tab(1).ControlCount=   1
               TabCaption(2)   =   "Неустановленные"
               TabPicture(2)   =   "frmMain.frx":00F0
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "ctlScrollControlTab2(0)"
               Tab(2).ControlCount=   1
               TabCaption(3)   =   "Установленные"
               TabPicture(3)   =   "frmMain.frx":010C
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "ctlScrollControlTab3(0)"
               Tab(3).ControlCount=   1
               TabCaption(4)   =   "БД не создана"
               TabPicture(4)   =   "frmMain.frx":0128
               Tab(4).ControlEnabled=   0   'False
               Tab(4).Control(0)=   "ctlScrollControlTab4(0)"
               Tab(4).ControlCount=   1
               Begin prjDIADBS.ctlScrollControl ctlScrollControl1 
                  Height          =   1575
                  Index           =   0
                  Left            =   25
                  TabIndex        =   24
                  TabStop         =   0   'False
                  Top             =   350
                  Width           =   4155
                  _ExtentX        =   7329
                  _ExtentY        =   2778
                  AutoScrollToFocus=   0   'False
               End
               Begin prjDIADBS.ctlScrollControl ctlScrollControlTab1 
                  Height          =   1575
                  Index           =   0
                  Left            =   -74975
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   350
                  Width           =   4095
                  _ExtentX        =   7223
                  _ExtentY        =   2778
                  AutoScrollToFocus=   0   'False
               End
               Begin prjDIADBS.ctlScrollControl ctlScrollControlTab2 
                  Height          =   1575
                  Index           =   0
                  Left            =   -74975
                  TabIndex        =   20
                  TabStop         =   0   'False
                  Top             =   350
                  Width           =   4095
                  _ExtentX        =   7223
                  _ExtentY        =   2778
                  AutoScrollToFocus=   0   'False
               End
               Begin prjDIADBS.ctlScrollControl ctlScrollControlTab3 
                  Height          =   1575
                  Index           =   0
                  Left            =   -74975
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   350
                  Width           =   4095
                  _ExtentX        =   7223
                  _ExtentY        =   2778
                  AutoScrollToFocus=   0   'False
               End
               Begin prjDIADBS.ctlScrollControl ctlScrollControlTab4 
                  Height          =   1575
                  Index           =   0
                  Left            =   -74975
                  TabIndex        =   23
                  TabStop         =   0   'False
                  Top             =   350
                  Width           =   4095
                  _ExtentX        =   7223
                  _ExtentY        =   2778
                  AutoScrollToFocus=   0   'False
               End
               Begin prjDIADBS.LabelW lblNoDP4Mode 
                  Height          =   285
                  Left            =   105
                  TabIndex        =   25
                  Top             =   2460
                  Visible         =   0   'False
                  Width           =   10590
                  _ExtentX        =   18680
                  _ExtentY        =   503
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Alignment       =   2
                  BackStyle       =   0
                  Caption         =   "Нет пакетов драйверов для данного режима работы"
               End
            End
            Begin prjDIADBS.LabelW lblNoDPInProgram 
               Height          =   285
               Left            =   120
               TabIndex        =   26
               Top             =   3600
               Visible         =   0   'False
               Width           =   10920
               _ExtentX        =   19262
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               BackStyle       =   0
               Caption         =   "Программа не обнаружила пакетов драйверов, или не верно настроены пути"
               AutoSize        =   -1  'True
            End
         End
      End
      Begin prjDIADBS.ctlJCFrames frRunChecked 
         Height          =   2550
         Left            =   7920
         Top             =   1250
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   4498
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
         Caption         =   "Выполнить ... "
         Alignment       =   0
         Begin prjDIADBS.ctlJCFrames frCheck 
            Height          =   1350
            Left            =   0
            Top             =   1200
            Width           =   3324
            _ExtentX        =   5874
            _ExtentY        =   2381
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
            Caption         =   "Выделение пакетов драйверов:"
            Alignment       =   0
            GradientHeaderStyle=   1
            Begin prjDIADBS.ComboBoxW cmbCheckButton 
               Height          =   315
               Left            =   120
               TabIndex        =   16
               Top             =   405
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
               Style           =   2
               Text            =   "frmMain.frx":0144
               CueBanner       =   "frmMain.frx":0180
            End
            Begin prjDIADBS.ctlJCbutton cmdCheck 
               Height          =   430
               Left            =   120
               TabIndex        =   17
               Top             =   800
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   767
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
               BackColor       =   12244692
               Caption         =   "Выделить"
               CaptionEffects  =   0
               PictureAlign    =   0
               PicturePushOnHover=   -1  'True
               PictureShadow   =   -1  'True
               ColorScheme     =   3
            End
         End
         Begin prjDIADBS.ctlJCbutton cmdRunTask 
            Default         =   -1  'True
            Height          =   675
            Left            =   120
            TabIndex        =   15
            Top             =   420
            Width           =   3120
            _ExtentX        =   4524
            _ExtentY        =   1191
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
            BackColor       =   12244692
            Caption         =   "Выполнить задание для выбранных пакетов драйверов на вкладке"
            CaptionEffects  =   0
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            DropDownSymbol  =   6
            DropDownSeparator=   -1  'True
            DropDownEnable  =   -1  'True
            ColorScheme     =   3
         End
      End
      Begin prjDIADBS.ctlJCFrames frRezim 
         Height          =   1550
         Left            =   75
         Top             =   1250
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   2725
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
         Caption         =   "Режим работы программы с пакетами драйверов"
         TextBoxHeight   =   20
         Alignment       =   0
         Begin prjDIADBS.ctlJCbutton cmdViewAllDevice 
            Height          =   510
            Left            =   120
            TabIndex        =   5
            Top             =   930
            Width           =   7575
            _ExtentX        =   13361
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
            BackColor       =   12244692
            Caption         =   "Список всех устройств вашего компьютера + Поиск драйвера в интернете по HWID"
            CaptionEffects  =   0
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            ColorScheme     =   3
         End
         Begin prjDIADBS.ctlJCbutton optRezim_Intellect 
            Height          =   510
            Left            =   120
            TabIndex        =   2
            Top             =   350
            Width           =   2415
            _ExtentX        =   4260
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
            BackColor       =   12244692
            Caption         =   "Установка (Совместимые драйвера)"
            CaptionEffects  =   0
            Mode            =   2
            Value           =   -1  'True
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            ColorScheme     =   3
         End
         Begin prjDIADBS.ctlJCbutton optRezim_Upd 
            Height          =   510
            Left            =   5280
            TabIndex        =   4
            Top             =   350
            Width           =   2415
            _ExtentX        =   4260
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
            BackColor       =   12244692
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
            Left            =   2640
            TabIndex        =   3
            Top             =   350
            Width           =   2535
            _ExtentX        =   4471
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
            BackColor       =   12244692
            Caption         =   "Установка (Полная - весь пакет)"
            CaptionEffects  =   0
            Mode            =   2
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            ColorScheme     =   3
         End
      End
      Begin prjDIADBS.ctlJCFrames frDescriptionIco 
         Height          =   950
         Left            =   75
         Top             =   2830
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   1667
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
         Collapsado      =   -1  'True
         Style           =   3
         RoundedCorner   =   0   'False
         Caption         =   "Обозначения кнопок (наведите курсор на картинку для просмотра описания)"
         TextBoxHeight   =   20
         Alignment       =   0
         GradientHeaderStyle=   1
         Begin VB.PictureBox imgOkAttentionOld 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            DrawStyle       =   5  'Transparent
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgUpdBD 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            DrawStyle       =   5  'Transparent
            ForeColor       =   &H80000008&
            Height          =   510
            Left            =   6840
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   350
            Width           =   510
         End
         Begin VB.PictureBox imgNoDB 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   6000
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgNo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   5160
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgOkOld 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   4320
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgOkNew 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3480
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgOkAttention 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2640
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgOK 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1800
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgOkAttentionNew 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            DrawStyle       =   5  'Transparent
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   960
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
      End
   End
   Begin prjDIADBS.ToolTip TTOtherControl 
      Left            =   1500
      Top             =   9000
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
      Title           =   "frmMain.frx":01A0
   End
   Begin prjDIADBS.ToolTip TTStatusIcon 
      Left            =   900
      Top             =   9000
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
      Title           =   "frmMain.frx":01C0
   End
   Begin prjDIADBS.ToolTip TT 
      Left            =   300
      Top             =   9000
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
      Title           =   "frmMain.frx":01E0
   End
   Begin VB.Menu mnuRezim 
      Caption         =   "Обновление баз данных"
      Begin VB.Menu mnuRezimBaseDrvUpdateALL 
         Caption         =   "Обновить базы для ВСЕХ пакетов драйверов"
      End
      Begin VB.Menu mnuRezimBaseDrvUpdateNew 
         Caption         =   "Обновить базы только для НОВЫХ пакетов драйверов"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRezimBaseDrvClean 
         Caption         =   "Удалить файлы баз данных отсутствующих пакетов драйверов"
      End
      Begin VB.Menu mnuDelDuplicateOldDP 
         Caption         =   "Удалить устаревшие версии пакетов драйверов"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadOtherPC 
         Caption         =   "Загрузить информацию другого ПК (Эмуляция работы)"
      End
      Begin VB.Menu mnuSaveInfoPC 
         Caption         =   "Сохранить информацию об устройствах для эмуляции на другом ПК"
      End
   End
   Begin VB.Menu mnuService 
      Caption         =   "Сервис"
      Begin VB.Menu mnuShowHwidsTxt 
         Caption         =   "Показать HWIDs устройств компьютера (текстовый файл)"
      End
      Begin VB.Menu mnuShowHwidsXLS 
         Caption         =   "Показать HWIDs устройств компьютера (файл Excel)"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowHwidsAll 
         Caption         =   "Показать ПОЛНЫЙ СПИСОК УСТРОЙСТВ компьютера"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdateStatusAll 
         Caption         =   "Обновить статус всех пакетов драйверов"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuUpdateStatusTab 
         Caption         =   "Обновить статус всех пакетов драйверов (текущая вкладка)"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReCollectHWID 
         Caption         =   "Обновить конфигурацию оборудования"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuReCollectHWIDTab 
         Caption         =   "Обновить конфигурацию оборудования (текущая вкладка)"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuAutoInfoAfterDelDRV 
         Caption         =   "Автообновление конфигурации при удалении драйверов"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRunSilentMode 
         Caption         =   "Запустить тихую автоматическую установку драйверов"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateRestorePoint 
         Caption         =   "Создать точку восстановления системы"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateBackUp 
         Caption         =   "Создать резервную копию драйверов"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDPInstLog 
         Caption         =   "Просмотреть DPinst.log"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Параметры"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuMainUtils 
      Caption         =   "Утилиты"
      Begin VB.Menu mnuUtils_devmgmt 
         Caption         =   "Диспетчер устройств Windows"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuUtils_DevManView 
         Caption         =   "DevManView"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuUtils_DoubleDriver 
         Caption         =   "DoubleDriver"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuUtils_SIV 
         Caption         =   "System Information Viewer"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuUtils_UDI 
         Caption         =   "Unknown Device Identifier"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuUtils_UnknownDevices 
         Caption         =   "Unknown Devices"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtils 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMainAbout 
      Caption         =   "Справка"
      Begin VB.Menu mnuLinks 
         Caption         =   "Ссылки"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "История изменения"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Справка по работе"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHomePage1 
         Caption         =   "Домашная страница программы"
      End
      Begin VB.Menu mnuHomePage 
         Caption         =   "Обсуждение программы на OsZone.net"
      End
      Begin VB.Menu mnuDriverPacks 
         Caption         =   "Посетить сайт driverpacks.net"
      End
      Begin VB.Menu mnuDriverPacksOnMySite 
         Caption         =   "Скачать пакеты драйверов..."
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckUpd 
         Caption         =   "Проверить обновление программы"
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModulesVersion 
         Caption         =   "Модули..."
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDonate 
         Caption         =   "Поблагодарить автора..."
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "Лицензионное соглашение..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе..."
      End
   End
   Begin VB.Menu mnuMainLang 
      Caption         =   "Язык"
      Begin VB.Menu mnuLangStart 
         Caption         =   "Использовать выбранный язык при запуске (отмена автовыбора)"
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLang 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuContextMenu 
      Caption         =   "Контекстное меню"
      Begin VB.Menu mnuContextXLS 
         Caption         =   "Открыть файл базы данных в программе Excel"
      End
      Begin VB.Menu mnuContextTxt 
         Caption         =   "Открыть файл базы данных в текстовом виде"
      End
      Begin VB.Menu mnuContextSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextToolTip 
         Caption         =   "Показать список доступных драйверов для компьютера"
      End
      Begin VB.Menu mnuContextSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextUpdStatus 
         Caption         =   "Обновить статус пакета драйверов"
      End
      Begin VB.Menu mnuContextSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextEditDPName 
         Caption         =   "Изменить отображаемое имя пакета драйверов в программе"
      End
      Begin VB.Menu mnuContextSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextTestDRP 
         Caption         =   "Протестировать данный пакет драйверов программой 7-zip"
      End
      Begin VB.Menu mnuContextSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextDeleteDRP 
         Caption         =   "Удалить пакет драйверов"
      End
      Begin VB.Menu mnuContextSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextDeleteDevIDs 
         Caption         =   "Удалить драйвера устройств:"
         Begin VB.Menu mnuContextDeleteDevIDDesc 
            Caption         =   "Список драйверов доступных для удаления"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuContextSep7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuContextDeleteDevID 
            Caption         =   "Список устройств"
            Index           =   0
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuContextCopyHWIDs 
         Caption         =   "Скопировать HWID в буфер обмена:"
         Begin VB.Menu mnuContextCopyHWIDDesc 
            Caption         =   "Список доступных HWID"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuContextSep8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuContextCopyHWID2Clipboard 
            Caption         =   "Список устройств"
            Index           =   0
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuContextMenu2 
      Caption         =   "Контекстное меню2"
      Begin VB.Menu mnuContextLegendIco 
         Caption         =   "Просмотреть описание всех обозначений"
      End
   End
   Begin VB.Menu mnuContextMenu3 
      Caption         =   "Контекстное меню3"
      Begin VB.Menu mnuContextInstallGroupDP 
         Caption         =   "Обычная установка"
         Index           =   0
      End
      Begin VB.Menu mnuContextInstallGroupDP 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContextInstallGroupDP 
         Caption         =   "Выборочная установка"
         Index           =   2
      End
      Begin VB.Menu mnuContextInstallGroupDP 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextInstallGroupDP 
         Caption         =   "Распаковать в каталог - Все подобранные драйвера"
         Index           =   4
      End
      Begin VB.Menu mnuContextInstallGroupDP 
         Caption         =   "Распаковать в каталог - Выбрать драйвера..."
         Index           =   5
      End
   End
   Begin VB.Menu mnuContextMenu4 
      Caption         =   "Контекстное меню4"
      Begin VB.Menu mnuContextInstallSingleDP 
         Caption         =   "Обычная установка"
         Index           =   0
      End
      Begin VB.Menu mnuContextInstallSingleDP 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContextInstallSingleDP 
         Caption         =   "Выборочная установка"
         Index           =   2
      End
      Begin VB.Menu mnuContextInstallSingleDP 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextInstallSingleDP 
         Caption         =   "Распаковать в каталог - Все подобранные драйвера"
         Index           =   4
      End
      Begin VB.Menu mnuContextInstallSingleDP 
         Caption         =   "Распаковать в каталог - Выбрать драйвера..."
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngCntBtn                   As Long
Private lngSSTabCurrentOS           As Long
Private lngFirstActiveTabIndex      As Long
Private lngNotFinedDriversInDP      As Long
Private lngFrameTime                As Long
Private lngFrameCount               As Long
Private lngOffSideCount             As Long         ' Кол-во переходов строк при построении кнопок

Private mbNextTab                   As Boolean
Private mbStatusHwid                As Boolean
Private mbStatusNewer               As Boolean
Private mbStatusOlder               As Boolean
Private mbUnpackAdditionalFile      As Boolean
Private mbNoSupportedOS             As Boolean
Private mbNotSupportedDevDB         As Boolean
Private mbLoadAppEnd                As Boolean
Private mbSet2UpdateFromTab4        As Boolean
Private mbOffSideButton             As Boolean      ' Флаг, указывающий что надо переходить на следующую строку при построении кнопок
Private mbDevParserRun              As Boolean      ' Флаг, указывающий что начата обработка пакета, защита от двойного нажатия
Private mbBreakUpdateDBAll          As Boolean      ' Флаг, указывающий что нажата кнопка прерывания процесса групповой обработки пакетов
Private mbIgnorStatusHwid           As Boolean
Private mbDRVNotInstall             As Boolean

Private strFormName                 As String
Private strCurSelButtonPath         As String
Private strSSTabCurrentOSList       As String
Private strCmbChkBtnListElement1    As String
Private strCmbChkBtnListElement2    As String
Private strCmbChkBtnListElement3    As String
Private strCmbChkBtnListElement4    As String
Private strCmbChkBtnListElement5    As String
Private strCmbChkBtnListElement6    As String
Private strTTipTextTitle            As String
Private strTTipTextFileSize         As String
Private strTTipTextClassDRV         As String
Private strTTipTextDrv2Install      As String
Private strTTipTextDrv4UnsupOS      As String
Private strTTipTextTitleStatus      As String
Private strSSTabTypeDPTab1          As String
Private strSSTabTypeDPTab2          As String
Private strSSTabTypeDPTab3          As String
Private strSSTabTypeDPTab4          As String
Private strSSTabTypeDPTab5          As String
Private strTTipTextHeaders          As String       ' Заголовок для Подсказки пакета драйверов

Private objHashOutput               As Scripting.Dictionary
Private objHashOutput2              As Scripting.Dictionary
Private objHashOutput3              As Scripting.Dictionary

Private objRegExpCheck              As RegExp
Private objRegExpCompat             As RegExp

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
'! Procedure   (Функция)   :   Sub BaseUpdateOrRunTask
'! Description (Описание)  :   [Обновление всех баз или только новых поочередно]
'! Parameters  (Переменные):   mbOnlyNew (Boolean = False)
'                              mbTasks (Boolean = False)
'!--------------------------------------------------------------------------------
Private Sub BaseUpdateOrRunTask(Optional ByVal mbOnlyNew As Boolean = False, Optional ByVal mbTasks As Boolean = False)

    Dim ButtIndex             As Long
    Dim ButtCount             As Long
    Dim i                     As Integer
    Dim TimeScriptRun         As Long
    Dim TimeScriptFinish      As Long
    Dim AllTimeScriptRun      As String
    Dim miPbInterval          As Long
    Dim miPbNext              As Long
    Dim strTextNew            As String
    Dim mbDpNoDBExist         As Boolean
    Dim strMsg                As String
    Dim lngFindCheckCountTemp As Long
    Dim lngSStabStart         As Long
    Dim lngNumButtOnTab       As Long

    If mbDebugStandart Then DebugMode "BaseUpdateOrRunTask-Start"
    
    mbBreakUpdateDBAll = False
    lngSStabStart = SSTab1.Tab
    strTextNew = strSpace
    strMsg = strMessages(24) & strTextNew & strMessages(25)

    If mbTasks Then
        strMsg = strMessages(23)
    End If

    If Not mbTasks Then
        If MsgBox(strMsg, vbQuestion + vbYesNo, strMessages(26)) = vbNo Then
            GoTo TheEnd
        End If
    End If

    BlockControl False
    DoEvents

    If Not mbTasks Then

        ' Переключаем программу в режим обновления
        If Not optRezim_Upd.Value Then
            SelectStartMode 3, False
        End If

        If SSTab2(lngSStabStart).Tab > 0 Then
            SSTab2(lngSStabStart).Tab = 0
        End If
    End If

    SSTab1.Tab = lngFirstActiveTabIndex
    cmdBreakUpdateDB.Visible = True
    TimeScriptRun = 0
    AllTimeScriptRun = vbNullString
    TimeScriptRun = GetTickCount
    ButtIndex = acmdPackFiles.UBound
    ButtCount = acmdPackFiles.Count
    ' Отображаем ProgressBar
    CreateProgressNew

    If ButtIndex Then
        ' В цикле обрабатываем обновление
        miPbInterval = 1000 / ButtCount

        If mbTasks Then
            lngFindCheckCountTemp = FindCheckCount

            If lngFindCheckCountTemp Then
                miPbInterval = 1000 / lngFindCheckCountTemp
            End If
        End If

        miPbNext = 0

        For i = 0 To ButtIndex
            lngNumButtOnTab = arrOSList(SSTab1.Tab).CntBtn

            Do While i >= lngNumButtOnTab
                SSTab1.Tab = SSTab1.Tab + 1
                DoEvents
                lngNumButtOnTab = arrOSList(SSTab1.Tab).CntBtn
            Loop

            ' Прерываем процесс обновления
            If mbBreakUpdateDBAll Then
                MsgBox strMessages(27) & vbNewLine & acmdPackFiles(i).Tag, vbInformation, strProductName

                Exit For

            End If

            If mbOnlyNew Then
                If acmdPackFiles(i).PictureNormal = imgNoDB.Picture Then
                    mbDpNoDBExist = True
                    acmdPackFiles_Click i
                    miPbNext = miPbNext + miPbInterval
                End If

            Else
                mbDpNoDBExist = True

                If mbTasks Then
                    If chkPackFiles(i).Value Then
                        acmdPackFiles_Click i
                        miPbNext = miPbNext + miPbInterval

                        If chkPackFiles(i).Value Then
                            chkPackFiles(i).Value = False
                            FindCheckCount
                        End If
                    End If

                Else
                    acmdPackFiles_Click i
                    miPbNext = miPbNext + miPbInterval
                End If
            End If

            If miPbNext > 1000 Then
                miPbNext = 1000
            End If

            With ctlProgressBar1
                .Value = miPbNext
                .SetTaskBarProgressValue miPbNext, 1000
            End With

            ChangeFrmMainCaption miPbNext
        Next

    Else

        If mbOnlyNew Then
            If acmdPackFiles(0).PictureNormal = imgNoDB.Picture Then
                acmdPackFiles_Click 0
                mbDpNoDBExist = True
            End If

        Else
            acmdPackFiles_Click 0
            mbDpNoDBExist = True
        End If

        If chkPackFiles(0).Value Then
            chkPackFiles(0).Value = False
            FindCheckCount
        End If
    End If

    TimeScriptFinish = GetTickCount
    AllTimeScriptRun = CalculateTime(TimeScriptRun, TimeScriptFinish)

    If mbBreakUpdateDBAll Then
        cmdBreakUpdateDB.Visible = False
        If mbDebugStandart Then DebugMode strMessages(66) & strSpace & AllTimeScriptRun
        ChangeStatusTextAndDebug strMessages(66) & strSpace & AllTimeScriptRun
    Else

        If mbDpNoDBExist Then
            If mbDebugStandart Then DebugMode strMessages(67) & strSpace & AllTimeScriptRun
            ChangeStatusTextAndDebug strMessages(67) & strSpace & AllTimeScriptRun
        Else
            If mbDebugStandart Then DebugMode strMessages(68) & strSpace & AllTimeScriptRun
            ChangeStatusTextAndDebug strMessages(68) & strSpace & AllTimeScriptRun
        End If
    End If

    ChangeFrmMainCaption
    pbProgressBar.Visible = False
    ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
    cmdBreakUpdateDB.Visible = False
    BlockControl True
    
TheEnd:
    If mbTasks Then
        ' Обновить список неизвестных дров и описание для кнопки
        LoadCmdViewAllDeviceCaption
    End If
        
    mbTasks = False
    SSTab1.Tab = lngSStabStart
    DoEvents
    If mbDebugStandart Then DebugMode "BaseUpdateOrRunTask-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub BlockControl
'! Description (Описание)  :   [Блокировка(Разблокировка) некоторых элементов формы при работе сложных функций]
'! Parameters  (Переменные):   mbBlock (Boolean)
'!--------------------------------------------------------------------------------
Public Sub BlockControl(ByVal mbBlock As Boolean)
    mnuRezim.Enabled = mbBlock
    mnuService.Enabled = mbBlock
    mnuMainUtils.Enabled = mbBlock
    mnuMainAbout.Enabled = mbBlock
    mnuMainLang.Enabled = mbBlock
    optRezim_Intellect.Enabled = mbBlock
    optRezim_Ust.Enabled = mbBlock
    optRezim_Upd.Enabled = mbBlock
    cmdViewAllDevice.Enabled = mbBlock
    cmbCheckButton.Enabled = mbBlock
    cmdCheck.Enabled = mbBlock
    cmdRunTask.Enabled = mbBlock
    imgNo.Enabled = mbBlock
    imgNoDB.Enabled = mbBlock
    imgOK.Enabled = mbBlock
    imgOkAttention.Enabled = mbBlock
    imgOkAttentionOLD.Enabled = mbBlock
    imgOkNew.Enabled = mbBlock
    imgOkOld.Enabled = mbBlock
    imgUpdBD.Enabled = mbBlock
    SSTab1.Enabled = mbBlock
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub BlockControlEx
'! Description (Описание)  :   [Блокировка(Разблокировка) элементов]
'! Parameters  (Переменные):   mbBlock (Boolean)
'!--------------------------------------------------------------------------------
Private Sub BlockControlEx(ByVal mbBlock As Boolean)
    mnuRunSilentMode.Enabled = mbBlock
    optRezim_Ust.Enabled = mbBlock
    optRezim_Intellect.Enabled = mbBlock
    optRezim_Upd.Enabled = mbBlock
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub BlockControlInNoDP
'! Description (Описание)  :   [Блокировка(Разблокировка) элементов если нет пакетов драйверов]
'! Parameters  (Переменные):   mbBlock (Boolean)
'!--------------------------------------------------------------------------------
Private Sub BlockControlInNoDP(ByVal mbBlock As Boolean)
    mnuRezimBaseDrvUpdateALL.Enabled = mbBlock
    mnuRezimBaseDrvUpdateNew.Enabled = mbBlock
    mnuRezimBaseDrvClean.Enabled = mbBlock
    mnuDelDuplicateOldDP.Enabled = mbBlock
    mnuRunSilentMode.Enabled = mbBlock
    cmbCheckButton.Enabled = mbBlock
    cmdCheck.Enabled = mbBlock
    cmdRunTask.Enabled = mbBlock
    optRezim_Intellect.Enabled = mbBlock
    optRezim_Ust.Enabled = mbBlock
    optRezim_Upd.Enabled = mbBlock
    mnuUpdateStatusAll.Enabled = mbBlock
    mnuUpdateStatusTab.Enabled = mbBlock
    mnuReCollectHWID.Enabled = mbBlock
    mnuReCollectHWIDTab.Enabled = mbBlock
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CalculateUnknownDrivers
'! Description (Описание)  :   [Подсчитываем количество неизвестных драйверов]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function CalculateUnknownDrivers() As Long

    Dim ii              As Long
    Dim lngCountUnknown As Long

    For ii = 0 To UBound(arrHwidsLocal)

        If LenB(arrHwidsLocal(ii).DPsList) = 0 Then

            ' Если это OEM-драйвер
            If InStr(1, arrHwidsLocal(ii).Provider, "microsoft", vbTextCompare) = 0 Then
                If InStr(1, arrHwidsLocal(ii).Provider, "майкрософт", vbTextCompare) = 0 Then
                    If InStr(1, arrHwidsLocal(ii).Provider, "standard", vbTextCompare) = 0 Then
                        lngCountUnknown = lngCountUnknown + 1
                    End If
                End If
            End If
        End If

    Next

    CalculateUnknownDrivers = lngCountUnknown
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ChangeFrmMainCaption
'! Description (Описание)  :   [Изменение Caption Формы]
'! Parameters  (Переменные):   lngstrPercentage (Long)
'!--------------------------------------------------------------------------------
Private Sub ChangeFrmMainCaption(Optional ByVal lngstrPercentage As Long)

    Dim strProgressValue As String

    Select Case strPCLangCurrentID

        Case "0419"
            strFrmMainCaptionTemp = "Помощник установки драйверов"
            strFrmMainCaptionTempDate = " (Дата релиза: "

        Case Else
            strFrmMainCaptionTemp = "Drivers Installer Assistant"
            strFrmMainCaptionTempDate = " (Date Build: "
    End Select

    If lngstrPercentage Mod 999 Then
        If ctlProgressBar1.Visible Then
            strProgressValue = (lngstrPercentage \ 10) & "% (" & ctlUcStatusBar1.PanelText(1) & ") - "
        End If
    End If

    If LenB(strThisBuildBy) = 0 Then
        Me.CaptionW = strProgressValue & strFrmMainCaptionTemp & " v." & strProductVersion & " @" & App.CompanyName & " - " & strPCLangCurrentLangName
    Else
        Me.CaptionW = strProgressValue & strFrmMainCaptionTemp & " v." & strProductVersion & strSpace & strThisBuildBy & " - " & strPCLangCurrentLangName
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ChangeStatusAndPictureButton
'! Description (Описание)  :   [Присваиваем картинку в соответсвии с наличием БД к файлу]
'! Parameters  (Переменные):   strPathDevDB (String)
'                              strPackFileName (String)
'                              ButtonIndex (Long)
'!--------------------------------------------------------------------------------
Private Function ChangeStatusAndPictureButton(ByVal strPathDevDB As String, ByVal strPackFileName As String, ByVal ButtonIndex As Long) As String

    Dim strTextHwids As String
    Dim mbUnSuppOS   As Boolean

    If mbDebugDetail Then DebugMode str4VbTab & "ChangeStatusAndPictureButton: strPackFileName=" & strPackFileName
              
    ' Небольшое прерывание для большего отклика от приложения
    DoEvents
    ChangeStatusAndPictureButton = vbNullString

    With acmdPackFiles(ButtonIndex)

        If CheckExistDB(strPathDevDB, strPackFileName) Then

            ' Ищем совпадения HWID в DP в новом режиме
            If mbFirstStart Then
                If mbLoadUnSupportedOS Then
                    strTextHwids = FindHwidInBaseNew(strPathDevDB, UCase$(GetFileNameFromPath(strPackFileName)), ButtonIndex)
                Else

                    If InStr(arrOSList(SSTab1.Tab).Ver, strOSCurrentVersion) Then
                        If arrOSList(SSTab1.Tab).is64bit = 2 Then
                            strTextHwids = FindHwidInBaseNew(strPathDevDB, UCase$(GetFileNameFromPath(strPackFileName)), ButtonIndex)
                        ElseIf arrOSList(SSTab1.Tab).is64bit = 3 Then
                            strTextHwids = FindHwidInBaseNew(strPathDevDB, UCase$(GetFileNameFromPath(strPackFileName)), ButtonIndex)
                        Else
                            If mbIsWin64 = CBool(arrOSList(SSTab1.Tab).is64bit) Then
                                strTextHwids = FindHwidInBaseNew(strPathDevDB, UCase$(GetFileNameFromPath(strPackFileName)), ButtonIndex)
                            Else
                                mbUnSuppOS = True
                            End If
                        End If

                    Else
                        mbUnSuppOS = True
                    End If
                End If

            Else
                strTextHwids = FindHwidInBaseNew(strPathDevDB, UCase$(GetFileNameFromPath(strPackFileName)), ButtonIndex)
            End If
            
            If LenB(strTextHwids) Then
                ChangeStatusAndPictureButton = strTextHwids
                If mbDebugStandart Then DebugMode str4VbTab & "ChangeStatusAndPictureButton-Hwids in file for PC: " & str2vbNewLine & strTextHwids & vbNewLine

                If mbStatusHwid Then
                    If mbStatusNewer Then
                        Set .PictureNormal = Nothing
                        Set .PictureNormal = imgOkNew.Picture
                        If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOkNew"
                    ElseIf mbStatusOlder Then
                        Set .PictureNormal = Nothing
                        Set .PictureNormal = imgOkOld.Picture
                        If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOkOld"
                    Else
                        Set .PictureNormal = Nothing
                        Set .PictureNormal = imgOK.Picture
                        If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOK"
                    End If

                Else

                    If mbIgnorStatusHwid Then
                        If mbDRVNotInstall Then
                            If mbStatusNewer Then
                                Set .PictureNormal = Nothing
                                Set .PictureNormal = imgOkAttentionNew.Picture
                                If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOkAttentionNew"
                            ElseIf mbStatusOlder Then
                                Set .PictureNormal = Nothing
                                Set .PictureNormal = imgOkAttentionOLD.Picture
                                If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOkAttentionOld"
                            Else
                                Set .PictureNormal = Nothing
                                Set .PictureNormal = imgOkAttention.Picture
                                If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOkAttention"
                            End If

                        Else

                            If mbStatusNewer Then
                                Set .PictureNormal = Nothing
                                Set .PictureNormal = imgOkNew.Picture
                                If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOkNew"
                            ElseIf mbStatusOlder Then
                                Set .PictureNormal = Nothing
                                Set .PictureNormal = imgOkOld.Picture
                                If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOkOld"
                            Else
                                Set .PictureNormal = Nothing
                                Set .PictureNormal = imgOK.Picture
                                If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOK"
                            End If
                        End If

                    Else

                        If mbStatusNewer Then
                            Set .PictureNormal = Nothing
                            Set .PictureNormal = imgOkAttentionNew.Picture
                            If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOkAttentionNew"
                        ElseIf mbStatusOlder Then
                            Set .PictureNormal = Nothing
                            Set .PictureNormal = imgOkAttentionOLD.Picture
                            If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOkAttentionOld"
                        Else
                            Set .PictureNormal = Nothing
                            Set .PictureNormal = imgOkAttention.Picture
                            If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgOkAttention"
                        End If
                    End If
                End If

                .DropDownEnable = optRezim_Intellect.Value
                .SetPopupMenu mnuContextMenu4
                .SetPopupMenuRBT mnuContextMenu
            Else
                Set .PictureNormal = Nothing
                Set .PictureNormal = imgNo.Picture
                If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgNo"
                .SetPopupMenuRBT mnuContextMenu
                .DropDownEnable = False

                If mbUnSuppOS Then
                    ChangeStatusAndPictureButton = "unsupported"
                End If
            End If

        Else
            Set .PictureNormal = Nothing
            Set .PictureNormal = imgNoDB.Picture
            If mbDebugDetail Then DebugMode str3VbTab & "ChangeStatusAndPictureButton-ImageForButton: imgNoDB"
            .DropDownEnable = False
        End If

    End With

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CheckAllButton
'! Description (Описание)  :   [Выделение всех кнопок]
'! Parameters  (Переменные):   mbCheckAll (Boolean)
'!--------------------------------------------------------------------------------
Private Sub CheckAllButton(ByVal mbCheckAll As Boolean)

    Dim i As Long

    With acmdPackFiles
        For i = .LBound To .UBound
    
            If Not (.item(i).PictureNormal Is Nothing) Then
                If .item(i).Visible Then
                    chkPackFiles(i).Value = mbCheckAll
                End If
            End If
    
        Next
    End With
    
    cmdRunTask.Enabled = FindCheckCount
End Sub

' проверяем совместимость драйвера по вендору ноутбука
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CheckDRVbyNotebookVendor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strInfPath (String)
'!--------------------------------------------------------------------------------
Private Function CheckDRVbyNotebookVendor(ByVal strInfPath As String) As Boolean

    Dim i                 As Long
    Dim ii                As Long
    Dim strFilterList     As String
    Dim strFilterList_x() As String
    Dim mbFind            As Boolean

    For i = 0 To UBound(arrNotebookFilterList)
        strFilterList = arrNotebookFilterList(i)
        strFilterList_x() = Split(strFilterList, ";")

        For ii = 0 To UBound(strFilterList_x)

            'strCompModel = "dPackard_Bell_123"
            If MatchSpec(strCompModel, strFilterList_x(ii)) Then
                If InStr(strInfPath, strFilterList_x(0) & "_NB\") Then
                    mbFind = True

                    Exit For

                End If
            End If

        Next

        If mbFind Then

            Exit For

        End If

    Next

    CheckDRVbyNotebookVendor = mbFind
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CheckExistbyRegExp
'! Description (Описание)  :   [Функция проверяет есть ли искомый текст в источнике посредством RegEXP]
'! Parameters  (Переменные):   strSourceText (String)
'                              strSearchText (String)
'                              mbGetText (Boolean)
'                              strFindText (String)
'!--------------------------------------------------------------------------------
Private Function CheckExistByRegExp(ByVal strSourceText As String, ByVal strSearchText As String, Optional ByVal mbGetText As Boolean, Optional ByRef strFindText As String) As Boolean

    Dim objMatchesCheck As MatchCollection

    Set objRegExpCheck = New RegExp

    With objRegExpCheck
        .Pattern = strSearchText
        .IgnoreCase = True
        Set objMatchesCheck = .Execute(strSourceText)
    End With

    CheckExistByRegExp = objMatchesCheck.Count

    If mbGetText Then
        If CheckExistByRegExp Then
            strFindText = Trim$(objMatchesCheck.item(0).Value)
        End If
    End If

    ' Очистка переменных
    Set objRegExpCheck = Nothing
    Set objMatchesCheck = Nothing
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CheckExistDB
'! Description (Описание)  :   [Проверяет есть ли txt/hwid файл для данного архива]
'! Parameters  (Переменные):   strDevDBPath (String)
'                              strPackFileName (String)
'!--------------------------------------------------------------------------------
Private Function CheckExistDB(ByVal strDevDBPath As String, ByVal strPackFileName As String) As Boolean

    Dim strFileNameDevDB         As String
    Dim strPathFileNameDevDB     As String
    Dim strPathFileNameDevDBHwid As String
    Dim lngFileDBSize            As Long

    strFileNameDevDB = Replace$(strPackFileName, ".7Z", ".TXT", , , vbTextCompare)

    If InStr(1, strPackFileName, ".zip", vbTextCompare) Then
        strFileNameDevDB = Replace$(strPackFileName, ".ZIP", ".TXT", , , vbTextCompare)
    End If

    strDevDBPath = BackslashAdd2Path(strDevDBPath)

    If Not mbDP_Is_aFolder Then
        strPathFileNameDevDB = PathCombine(strDevDBPath, GetFileNameFromPath(strFileNameDevDB))
        strPathFileNameDevDBHwid = Replace$(strPathFileNameDevDB, ".TXT", ".HWID")
    Else
        strPathFileNameDevDB = PathCombine(strDevDBPath, GetFileNameFromPath(strPackFileName) & ".TXT")
        strPathFileNameDevDBHwid = Replace$(strPathFileNameDevDB, ".TXT", ".HWID")
    End If

    strCurSelButtonPath = strPathFileNameDevDB

    If PathExists(strPathFileNameDevDBHwid) Then
        lngFileDBSize = GetFileSizeByPath(strPathFileNameDevDBHwid)
        If mbDebugDetail Then DebugMode str5VbTab & "CheckExistDB: Find file=" & strPathFileNameDevDBHwid & " (FileSize: " & lngFileDBSize & " bytes)"

        If lngFileDBSize Then
            If PathExists(strPathFileNameDevDB) Then
                lngFileDBSize = GetFileSizeByPath(strPathFileNameDevDB)
                If mbDebugDetail Then DebugMode str5VbTab & "CheckExistDB: Find file=" & strPathFileNameDevDB & " (FileSize: " & lngFileDBSize & " bytes)"

                If lngFileDBSize Then
                    CheckExistDB = CompareDevDBVersion(strPathFileNameDevDB)
                    mbNotSupportedDevDB = Not CheckExistDB

                Else
                    If mbDebugDetail Then DebugMode str5VbTab & "CheckExistDB: File is zero = 0 bytes"
                End If

            Else
                If mbDebugDetail Then DebugMode str5VbTab & "CheckExistDB: NOT FIND DB FILE=" & strPathFileNameDevDB
            End If

        Else
            If mbDebugDetail Then DebugMode str5VbTab & "CheckExistDB: File is zero = 0 bytes"
        End If

    Else
        If mbDebugDetail Then DebugMode str5VbTab & "CheckExistDB: NOT FIND DB FILE=" & strPathFileNameDevDBHwid
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CheckMenuUtilsPath
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub CheckMenuUtilsPath()

    If mbIsWin64 Then
        If PathExists(PathCollect(strDevManView_Path64)) = False Then
            mnuUtils_DevManView.Enabled = False
        End If

        If PathExists(PathCollect(strSIV_Path64)) = False Then
            mnuUtils_SIV.Enabled = False
            lblOSInfo.MousePointer = 0
            lblOSInfo.ToolTipText = vbNullString
        End If

    Else

        If PathExists(PathCollect(strDevManView_Path)) = False Then
            mnuUtils_DevManView.Enabled = False
        End If

        If PathExists(PathCollect(strSIV_Path)) = False Then
            mnuUtils_SIV.Enabled = False
            lblOSInfo.MousePointer = 0
            lblOSInfo.ToolTipText = vbNullString
        End If
    End If

    If PathExists(PathCollect(strDoubleDriver_Path)) = False Then
        mnuUtils_DoubleDriver.Enabled = False
    End If

    If PathExists(PathCollect(strUDI_Path)) = False Then
        mnuUtils_UDI.Enabled = False
    End If

    If PathExists(PathCollect(strUnknownDevices_Path)) = False Then
        mnuUtils_UnknownDevices.Enabled = False
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CompatibleDriver4OS
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strSection (String)
'                              strDPFileName (String)
'                              strDPInfPath (String)
'                              strSectionUnsupported (String)
'!--------------------------------------------------------------------------------
Private Function CompatibleDriver4OS(ByVal strSection As String, ByVal strDPFileName As String, ByVal strDPInfPath As String, ByVal strSectionUnsupported As String) As Boolean

    Dim mbOSx64                   As Boolean
    Dim strOsVer                  As String
    Dim strDRVx64                 As String
    Dim lngDRVx64                 As Long
    Dim strDRVOSVer               As String
    Dim objMatch                  As Match
    Dim objMatches                As MatchCollection
    Dim mbCompatibleByArch        As Boolean
    Dim mbCompatibleByVer         As Boolean
    Dim mbVerFromSection          As Boolean
    Dim mbArchFromSection         As Boolean
    Dim mbVerFromMarkers          As Boolean
    Dim mbArchFromMarkers         As Boolean
    Dim mbVerFromDPName           As Boolean
    Dim mbArchFromDPName          As Boolean
    Dim strRegExpMarkerPattern    As String
    Dim mbMarkerCheckExist        As Boolean
    Dim mbMarkerSTRICTCheckExist  As Boolean
    Dim strSection_x()            As String
    Dim strSectionMain            As String
    Dim strSectionUnsupportedTemp As String
    Dim mbMarkerFORCEDCheckExist  As Boolean
    Dim strDRVOSVerUnsupported    As String

    mbOSx64 = mbIsWin64

    If Not mbSearchCompatibleDriverOtherOS Then
        strOsVer = arrOSList(SSTab1.Tab).Ver
    Else
        strOsVer = strOSCurrentVersion
    End If

    strDPInfPath = UCase$(strDPInfPath)
    ' На случай проверка работы программы
    'mbOSx64 = True
    'mbOSx64 = False
    'strOsVer = "6.1"
    'strSection = "ATHEROS.NTAMD64.6.1"
    'strDPFileName = "DP_WLAN_1300.7z"
    'strDPInfPath = "NTx86\220\"
    'strSectionUnsupported = "Atheros,Atheros.NT.6.0,Atheros.NTamd64.6.0"
    ' проверяем есть ли вообще маркеры в пути
    'strDPInfPath = "5x86\M\N\"
    mbMarkerCheckExist = CheckExistByRegExp(strDPInfPath, strVer_All_Known_Ver)
    ' проверяем есть ли вообще маркер FORCED в пути
    'strDPInfPath = "5x86\FORCED\M\N\"
    mbMarkerFORCEDCheckExist = CheckExistByRegExp(strDPInfPath, strVerFORCED & vbBackslashDouble)
    ' проверяем есть ли вообще маркер STRICT в пути
    'strDPInfPath = "5x86\STRICT\M\N\"
    mbMarkerSTRICTCheckExist = CheckExistByRegExp(strDPInfPath, strVerSTRICT & vbBackslashDouble)

    ' Если нет маркера FORCED, то получаем данные из секции
    'Debug.Print strDPInfPath
    If Not mbMarkerFORCEDCheckExist Then
        Set objRegExpCompat = New RegExp

        With objRegExpCompat
            .Pattern = "\.NT(X86|AMD64|IA64|)(?:\.(\d(?:.\d)))?"
            .IgnoreCase = True
            'strSection = "AMD.NTAMD64.5.1.1"
            Set objMatches = .Execute(strSection)
        End With

        'Получаем значения версии ОС драйвера и архитектуры
        'данные берем из секции Manufactured
        If objMatches.Count Then
            Set objMatch = objMatches.item(0)
            strDRVx64 = UCase$(Trim$(objMatch.SubMatches(0)))
            strDRVOSVer = UCase$(Trim$(objMatch.SubMatches(1)))
            lngDRVx64 = InStr(strDRVx64, "64")
        End If

    Else
        If mbDebugDetail Then DebugMode str6VbTab & "CompatibleDriver4OS: Check Inf-Section: " & strSection & " Inf-Path: " & strDPInfPath & " contained FORCED marker, section [Manufactured] not analyzing"
    End If

    ' Если секция не имеет окончания вида .NTX86.6.0 - т.е .Count=0, то тогда мы не можем определить точно подходит или нет.
    ' Считаем раз лежит в данной папке, то подходит.
    ' Если в секции manufactured не указано ля каких систем драйвер подходит, то анализируем имя файла
    ' !!! После сделаю поддержку маркеров
    ' Если версия не определена, определяем версию по маркерам или по имени DP
    If LenB(strDRVOSVer) = 0 Then
CheckVerByMarkers:

        Select Case strOSCurrentVersion

            Case "5.1"

                If mbOSx64 Then
                    strRegExpMarkerPattern = strVer_51x64
                Else
                    strRegExpMarkerPattern = strVer_51x86
                End If

                strRegExpMarkerPattern = strRegExpMarkerPattern & "|" & strVer_51xXX

            Case "6.1"

                If mbOSx64 Then
                    strRegExpMarkerPattern = strVer_61x64
                Else
                    strRegExpMarkerPattern = strVer_61x86
                End If

                strRegExpMarkerPattern = strRegExpMarkerPattern & "|" & strVer_61xXX

            Case "6.2"

                If mbOSx64 Then
                    strRegExpMarkerPattern = strVer_62x64
                Else
                    strRegExpMarkerPattern = strVer_62x86
                End If

                strRegExpMarkerPattern = strRegExpMarkerPattern & "|" & strVer_62xXX

            Case "6.3"

                If mbOSx64 Then
                    strRegExpMarkerPattern = strVer_63x64
                Else
                    strRegExpMarkerPattern = strVer_63x86
                End If

                strRegExpMarkerPattern = strRegExpMarkerPattern & "|" & strVer_63xXX

            Case "5.0"
                strRegExpMarkerPattern = vbNullString

            Case "5.2"

                If mbOSx64 Then
                    strRegExpMarkerPattern = strVer_51x64
                Else
                    strRegExpMarkerPattern = strVer_51x86
                End If

            Case "6.0"

                If mbOSx64 Then
                    strRegExpMarkerPattern = strVer_60x64
                Else
                    strRegExpMarkerPattern = strVer_60x86
                End If

                strRegExpMarkerPattern = strRegExpMarkerPattern & "|" & strVer_60xXX

            Case Else

                If mbOSx64 Then
                    strRegExpMarkerPattern = strVer_Any64
                Else
                    strRegExpMarkerPattern = strVer_Any86
                End If

        End Select

        ' Добавляем к поиску драйвера для всех ОС или для конкретной разрядности
        If mbOSx64 Then
            strRegExpMarkerPattern = strRegExpMarkerPattern & "|" & strVer_XXx64 & "|" & strVer_XXxXX
        Else
            strRegExpMarkerPattern = strRegExpMarkerPattern & "|" & strVer_XXx86 & "|" & strVer_XXxXX
        End If

        ' Поиск подходящей версии ОС в маркерах
        mbVerFromMarkers = CheckExistByRegExp(strDPInfPath, strRegExpMarkerPattern)

        If mbVerFromMarkers Then
            strDRVOSVer = strOSCurrentVersion
        Else

            ' Если по маркерам определить нельзя, определяем версию по имени DP
            If mbMatchHWIDbyDPName Then
                If Not mbMarkerCheckExist Then
                    If InStr(strDPFileName, "WXP") Then
                        strDRVOSVer = "5.0;5.1;5.2"
                    ElseIf InStr(strDPFileName, "WNT5") Then
                        strDRVOSVer = "5.0;5.1;5.2"
                    ElseIf InStr(strDPFileName, "WNT6") Then
                        strDRVOSVer = "6.0;6.1;6.2;6.3"
                    Else

                        If mbOSx64 Then
                            If InStr(strDRVx64, "AMD64") Then
                                strDRVOSVer = strOSCurrentVersion
                            End If

                        Else

                            If InStr(strDRVx64, "X86") Then
                                strDRVOSVer = strOSCurrentVersion
                            End If
                        End If
                    End If
                End If
            End If
        End If

    Else

        If mbMarkerSTRICTCheckExist Then
            If mbDebugDetail Then DebugMode str6VbTab & "CompatibleDriver4OS: Check Inf-Section: " & strSection & " Inf-Path: " & strDPInfPath & " contained STRICT marker, section [Manufactured] not analyzing by Version"
            GoTo CheckVerByMarkers
        Else
            mbVerFromSection = True
        End If
    End If

    ' Если архитектура не определена, определяем версию по маркерам или по имени DP
    If LenB(strDRVx64) = 0 Then
CheckVerByMarkersArch:

        If mbOSx64 Then
            strRegExpMarkerPattern = strVer_Any64 + "|" + strVer_XXxXX
        Else
            strRegExpMarkerPattern = strVer_Any86 + "|" + strVer_XXxXX
        End If

        ' Поиск подходящей версии ОС в маркерах
        mbArchFromMarkers = CheckExistByRegExp(strDPInfPath, strRegExpMarkerPattern, True, strDRVx64)

        If mbArchFromMarkers Then
            lngDRVx64 = InStr(strDRVx64, "X64")
        Else
            ' Если по маркерам определить нельзя, определяем версию по имени DP
            mbArchFromMarkers = False

            If mbMatchHWIDbyDPName And Not mbMarkerCheckExist Then
                lngDRVx64 = InStr(strDPFileName, "X64")
                mbArchFromDPName = True
            Else

                If mbMarkerCheckExist Then
                    lngDRVx64 = InStr(strDPInfPath, "X64")
                End If
            End If
        End If

    Else
        mbArchFromSection = True
    End If

    ' проверка на архитектуру
    If CBool(lngDRVx64) = mbOSx64 Then
        mbCompatibleByArch = True
    End If

    ' проверка на версию ОС
    If LenB(strDRVOSVer) Then
        If mbVerFromSection Then
            If InStr(strOsVer, strDRVOSVer) Then
                mbCompatibleByVer = True
            Else

                If CompareByVersion(strOsVer, strDRVOSVer) = ">" Then
                    mbCompatibleByVer = True
                End If
            End If

        Else

            If mbVerFromMarkers Then
                If InStr(strOsVer, strDRVOSVer) Then
                    mbCompatibleByVer = True
                End If

            Else

                If InStr(strDRVOSVer, strOsVer) Then
                    mbCompatibleByVer = True
                    mbVerFromDPName = True
                End If
            End If
        End If

    Else
        mbCompatibleByVer = False
    End If

    ' Проверка на несовместимые секции
    If mbCompatibleByVer Then
        If InStr(strSectionUnsupported, "-") = 0 Then
            strSectionUnsupportedTemp = strSectionUnsupported & ","
            strSection_x = Split(strSection, ".")
            strSectionMain = strSection_x(0)

            If StrComp(strOSCurrentVersion, "5.0") <> 0 Then
                Set objRegExpCompat = New RegExp

                With objRegExpCompat

                    If mbOSx64 Then
                        .Pattern = strSectionMain & "\.NT[AMD64|IA64]*(?:\.(\d(?:.\d)*)*)*,"
                    Else
                        .Pattern = strSectionMain & "\.NT[X86]*(?:\.(\d(?:.\d)*)*)*,"
                    End If

                    'ATHEROS\.NT[AMD64|IA64]*(?:\.(\d(?:.\d)*)*)*,
                    .IgnoreCase = True
                    'strSection = "AMD.NTAMD64.5.1.1"
                    '.Pattern = "Atheros,Atheros.NT.6.0,Atheros.NTamd64.6.0"
                    Set objMatches = .Execute(strSectionUnsupportedTemp)
                End With

                'Если несовместиые секции найдены
                'данные берем из секции Manufactured
                If objMatches.Count Then
                    Set objMatch = objMatches.item(0)
                    strDRVOSVerUnsupported = Trim$(objMatch.SubMatches(0))

                    If LenB(strDRVOSVerUnsupported) Then

                        ' Если в inf неподдерживаемые секции с версией например 6.0, то неподдериваются ос 6.0 и выше
                        ' т.е если текущая ОС больше чем найденная в inf пустая секция, т.е драйвер не поддерживается
                        If CompareByVersion(strOsVer, strDRVOSVerUnsupported) = ">" Then
                            ' Если в inf неподдерживаемые секции с версией например 6.0, а драйвер найден в секции 6.1, то драйвер найден правильно, иначе
                            If CompareByVersion(strDRVOSVerUnsupported, strDRVOSVer) = ">" Then
                                If mbDebugDetail Then DebugMode str6VbTab & "CompatibleDriver4OS: Check Inf-Section: " & strSection & " Result: " & CompatibleDriver4OS & " by SectionUnsupported:" & strSectionUnsupported
                                mbCompatibleByArch = False
                                mbCompatibleByVer = False
                            End If
                        ElseIf CompareByVersion(strOsVer, strDRVOSVerUnsupported) = "=" Then
                            ' Если в inf неподдерживаемые секции с версией например 6.0, а драйвер найден в секции 6.1, то драйвер найден правильно, иначе
                            If CompareByVersion(strDRVOSVerUnsupported, strDRVOSVer) = ">" Then
                                If mbDebugDetail Then DebugMode str6VbTab & "CompatibleDriver4OS: Check Inf-Section: " & strSection & " Result: " & CompatibleDriver4OS & " by SectionUnsupported:" & strSectionUnsupported
                                mbCompatibleByArch = False
                                mbCompatibleByVer = False
                            End If
                        End If
                    End If
                End If

            Else

                If UBound(strSection_x) < 1 Then
                    If mbDebugStandart Then DebugMode str6VbTab & "CompatibleDriver4OS: verOS=" & strOSCurrentVersion & " Check Inf-Section: " & strSection & " Result: " & CompatibleDriver4OS & " by SectionUnsupported:" & strSectionUnsupported
                    mbCompatibleByArch = False
                    mbCompatibleByVer = False
                End If
            End If
        End If
    End If

    'Итог
    If mbCompatibleByArch Then
        CompatibleDriver4OS = mbCompatibleByVer
    Else
        CompatibleDriver4OS = False
    End If

    ' Если это ноутбук, и пакет дров от тачпада, то проверяем совместимость драйвера по вендору ноутбука
    If CompatibleDriver4OS Then
        If InStr(strDPFileName, "TOUCHPAD") Then
            If mbIsNotebok Then
                If InStr(strDPInfPath, "_NB\") Then
                    CompatibleDriver4OS = CheckDRVbyNotebookVendor(strDPInfPath)
                End If

            Else
                CompatibleDriver4OS = False
            End If
        End If
    End If

    ' Чистка переменных
    Set objRegExpCompat = Nothing
    Set objMatch = Nothing
    Set objMatches = Nothing
    If mbDebugDetail Then DebugMode str6VbTab & "CompatibleDriver4OS: Check Inf-Section: " & strSection & " Result: " & CompatibleDriver4OS & " (by Version-" & mbCompatibleByVer & "; by Architecture-" & mbCompatibleByArch & "; by ManufacturedSection:Ver/Arch-" & _
                                mbVerFromSection & "/" & mbArchFromSection & "; by Markers:Ver/Arch-" & mbVerFromMarkers & "/" & mbArchFromMarkers & ")"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ConvertDPName
'! Description (Описание)  :   [Изменяем имя пакета драйверов (удаление лишних Символов)]
'! Parameters  (Переменные):   strButtonName (String)
'!--------------------------------------------------------------------------------
Private Function ConvertDPName(ByVal strButtonName As String) As String

    Dim strButtonNameTemp As String

    strButtonNameTemp = LCase$(strButtonName)

    If mbConvertDPName Then
        If InStr(strButtonNameTemp, ".7z") Then strButtonName = Replace$(strButtonName, ".7z", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, ".zip") Then strButtonName = Replace$(strButtonName, ".zip", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, "dp") Then strButtonName = Replace$(strButtonName, "dp", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, "driverpack") Then strButtonName = Replace$(strButtonName, "driverpack", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, "wnt5") Then strButtonName = Replace$(strButtonName, "wnt5", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, "wnt6") Then strButtonName = Replace$(strButtonName, "wnt6", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, "wxp") Then strButtonName = Replace$(strButtonName, "wxp", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, "x86") Then strButtonName = Replace$(strButtonName, "x86", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, "w7x64") Then strButtonName = Replace$(strButtonName, "w7x64", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, "_32") Then strButtonName = Replace$(strButtonName, "_32", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, "-32") Then strButtonName = Replace$(strButtonName, "-32", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, "x64") Then strButtonName = Replace$(strButtonName, "x64", vbNullString, , , vbTextCompare)
        If InStr(strButtonNameTemp, "_") Then strButtonName = Replace$(strButtonName, "_", strSpace)
        If InStr(strButtonNameTemp, "-") Then strButtonName = Replace$(strButtonName, "-", strSpace)
        If InStr(strButtonName, str3Space) Then strButtonName = Replace$(strButtonName, str3Space, strSpace)
        If InStr(strButtonName, str2Space) Then strButtonName = Replace$(strButtonName, str2Space, strSpace)
        strButtonName = Trim$(strButtonName)
    End If

    ' все в верхний регистр
    If mbButtonTextUpCase Then
        strButtonName = UCase$(strButtonName)
    End If

    ConvertDPName = strButtonName
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CreateButtonsonSSTab
'! Description (Описание)  :   [Создание кнопок на выбранной вкладке табконтрола]
'! Parameters  (Переменные):   strDrpPath (String)
'                              strDevDBPath (String)
'                              miTabIndex (Long)
'                              lngProgressDelta (Long)
'!--------------------------------------------------------------------------------
Private Sub CreateButtonsOnSSTab(ByVal strDrpPath As String, ByVal strDevDBPath As String, ByVal miTabIndex As Long, ByVal lngProgressDelta As Long)

    Dim strButtonName        As String
    Dim strPackFileName      As String
    Dim StartPositionLeft    As Long
    Dim StartPositionTop     As Long
    Dim NextPositionLeft     As Long
    Dim NextPositionTop      As Long
    Dim MaxLeftPosition      As Long
    Dim DeltaPositionLeft    As Long
    Dim DeltaPositionTop     As Long
    Dim mbStep               As Boolean
    Dim tabN                 As Long
    Dim TabHeight            As Long
    Dim ii                   As Long
    Dim strFileList_x()      As FindListStruct
    Dim lngOffSideCountTemp  As Long
    Dim strPhysXPath         As String
    Dim strLangPath          As String
    Dim strRuntimes          As String
    Dim lngFileCount         As Long
    Dim lngProgressDeltaTemp As Single

    On Error Resume Next

    If mbDebugStandart Then DebugMode vbTab & "CreateButtonsOnSSTab-Start" & vbNewLine & _
              str2VbTab & "CreateButtonsonSSTab: miTabIndex=" & miTabIndex

    If PathExists(strDrpPath) Then
        tabN = miTabIndex
        TabHeight = SSTab1.Height
        DoEvents
        SSTab1.Tab = tabN
        StartPositionLeft = lngButtonLeft
        StartPositionTop = lngButtonTop

        If tabN Then
            Load SSTab2(tabN)
            Set SSTab2(tabN).Container = SSTab1
            Load ctlScrollControl1(tabN)
            Set ctlScrollControl1(tabN).Container = SSTab2(tabN)
        End If

        With SSTab2(tabN)
            .Height = TabHeight - SSTab1.TabHeight - 50
            .Top = SSTab1.TabHeight + 50
            .Left = 0
            .Visible = True
            .Width = SSTab1.Width - 180
        End With

        With ctlScrollControl1(tabN)
            .Visible = True
            .Height = SSTab2(tabN).Height - SSTab2(tabN).TabHeight - 120
            .Width = SSTab2(tabN).Width - 100
        End With

        If lngOSCount > lngOSCountPerRow Then
            StartPositionTop = StartPositionTop
        End If

        mbStep = False
        If mbDebugStandart Then DebugMode str2VbTab & "CreateButtonsonSSTab: Recursion: " & mbRecursion & vbNewLine & _
                  str2VbTab & "CreateButtonsonSSTab: Get ListFile in folder: " & strDrpPath

        'Строим список файлов 7z
        If Not mbDP_Is_aFolder Then
            strFileList_x = SearchFilesInRoot(strDrpPath, "DP*.7z;DP*.zip", mbRecursion, False, False)
            'Иначе это каталоги, а не 7z
        Else

            If FolderContainsSubfolders(strDrpPath) Then
                strFileList_x = SearchFoldersInRoot(strDrpPath, "DP*")
            End If
        End If

        If mbDebugStandart Then DebugMode str2VbTab & "CreateButtonsonSSTab: FileCount: " & UBound(strFileList_x)

        If UBound(strFileList_x) = 0 Then
            If LenB(strFileList_x(0).FullPath) = 0 Then
                SSTab1.TabEnabled(tabN) = False

                If mbTabHide Then
                    SSTab1.TabVisible(tabN) = False
                End If

                With ctlProgressBar1
                    .Value = .Value + lngProgressDelta
                    .SetTaskBarProgressValue .Value, 1000
                    ChangeFrmMainCaption .Value
                End With

                Exit Sub

            End If
        End If

        strPhysXPath = GetFileNameFromPath(arrOSList(tabN).PathPhysX)
        strLangPath = GetFileNameFromPath(arrOSList(tabN).PathLanguages)
        strRuntimes = GetFileNameFromPath(arrOSList(tabN).PathRuntimes)
        strExcludeFileName = arrOSList(tabN).ExcludeFileName
        lngFileCount = UBound(strFileList_x) + 1
        pbProgressBar.Refresh

        For ii = 0 To UBound(strFileList_x)
            strPackFileName = Replace$(strFileList_x(ii).FullPath, BackslashAdd2Path(strDrpPath), vbNullString, , , vbTextCompare)
            If mbDebugStandart Then DebugMode "====================================================================================================" & vbNewLine & _
                      str2VbTab & "CreateButtonsOnSSTab-Work with File: " & strPackFileName
            ChangeStatusTextAndDebug strMessages(69) & strSpace & strDrpPath & strSpace & vbNewLine & strMessages(70) & "(" & (ii + 1) & strSpace & strMessages(124) & strSpace & lngFileCount & "): " & strPackFileName
            mbStatusHwid = True

            If Not mbDP_Is_aFolder Then
                strButtonName = strFileList_x(ii).Name
            Else
                strButtonName = strPackFileName
            End If

            ' проверяем что файл не является дополнительным для обработки вместе с архивом или исключаемым
            If LenB(strLangPath) Then
                If MatchSpec(strButtonName, strLangPath) Then
                    GoTo NextFiles
                End If
            End If
            
            If LenB(strRuntimes) Then
                If MatchSpec(strButtonName, strRuntimes) Then
                    GoTo NextFiles
                End If
            End If
            
            If LenB(strPhysXPath) Then
                If MatchSpec(strButtonName, strPhysXPath) Then
                    GoTo NextFiles
                End If
            End If
            
            If LenB(strExcludeFileName) Then
                If MatchSpec(strButtonName, strExcludeFileName) Then
                    GoTo NextFiles
                End If
            End If

            ' Изменяем имя пакета драйверов (удаление лишних Символов)
            strButtonName = ConvertDPName(strButtonName)

            If lngCntBtn = 0 Then
                NextPositionLeft = StartPositionLeft
                NextPositionTop = StartPositionTop
            Else

                If mbNextTab Then
                    ' Если переход на сл. вкладку, то
                    NextPositionLeft = StartPositionLeft
                    NextPositionTop = StartPositionTop
                    mbNextTab = False
                Else
                    DeltaPositionLeft = acmdPackFiles(lngCntBtn - 1).Left + lngButtonWidth + lngBtn2BtnLeft - StartPositionLeft
                    NextPositionLeft = StartPositionLeft + DeltaPositionLeft

                    ' Если Кол-во ОС больше кол-ва вкладок на строку
                    If lngOSCount > lngOSCountPerRow Then
                        MaxLeftPosition = NextPositionLeft + lngButtonWidth + 100 * (Abs(lngOSCount / lngOSCountPerRow) - 1)
                    Else
                        MaxLeftPosition = NextPositionLeft + lngButtonWidth + 25
                    End If

                    If MaxLeftPosition > ctlScrollControl1(tabN).Width Then
                        ' Если по горизонтали кнопка не входит, то перешагиваем
                        mbStep = True
                    Else
                        NextPositionTop = StartPositionTop + DeltaPositionTop
                    End If

                    ' Перешагиваем, если кнопки на одну строку не входят
                    If mbStep Then
                        DeltaPositionLeft = 0
                        DeltaPositionTop = DeltaPositionTop + lngButtonHeight + lngBtn2BtnTop
                        NextPositionLeft = StartPositionLeft
                        NextPositionTop = StartPositionTop + DeltaPositionTop

                        If NextPositionTop > TabHeight Then
                            mbOffSideButton = True
                            lngOffSideCountTemp = lngOffSideCountTemp + 1
                        End If

                        mbStep = False
                    End If
                End If
            End If

            ' Загружаем кнопку и чекбокс
            If lngCntBtn Then
                Load acmdPackFiles(lngCntBtn)
                Load chkPackFiles(lngCntBtn)
            Else
                mbNextTab = False
            End If

            ' Присваиваем кнопку табконтролу
            Set acmdPackFiles(lngCntBtn).Container = ctlScrollControl1(tabN)
            Set chkPackFiles(lngCntBtn).Container = ctlScrollControl1(tabN)

            ' Устанавливаем свойства кнопки
            With acmdPackFiles(lngCntBtn)
                .Left = NextPositionLeft
                .Top = NextPositionTop
                .Visible = True
                .Caption = strButtonName
                .Tag = strPackFileName
            End With

            ' Устанавливаем свойства чекбокса
            With chkPackFiles(lngCntBtn)
                .Visible = True
                .Left = NextPositionLeft + 50
                .Top = NextPositionTop + (lngButtonHeight - .Height) / 2
                .ZOrder 0
                .Tag = tabN
            End With

            'Считываем подменяемое имя пакета из файла
            EditOrReadDPName lngCntBtn, True

            ' массив HWID для будущего использования для каждой кнопки
            ReDim Preserve arrDevIDs(acmdPackFiles.UBound)

            ' Создаем подсказку к файлу, а также меняем иконку и т.д.
            ReadOrSaveToolTip strDevDBPath, strDrpPath, strPackFileName, lngCntBtn, True
            lngCntBtn = lngCntBtn + 1
NextFiles:
            lngProgressDeltaTemp = (lngProgressDelta / lngFileCount)

            With ctlProgressBar1
                lngProgressDeltaTemp = .Value + lngProgressDeltaTemp
                .Value = lngProgressDeltaTemp
                .SetTaskBarProgressValue lngProgressDeltaTemp, 1000
            End With

            ChangeFrmMainCaption lngProgressDeltaTemp
            pbProgressBar.Refresh
        Next

    End If

    If tabN Then
        SSTab2(tabN).Tab = 1
        DoEvents
        Load ctlScrollControlTab1(tabN)
        Set ctlScrollControlTab1(tabN).Container = SSTab2(tabN)
        ctlScrollControlTab1(tabN).Visible = True
        ctlScrollControlTab1(tabN).Left = 25
        SSTab2(tabN).Tab = 2
        DoEvents
        Load ctlScrollControlTab2(tabN)
        Set ctlScrollControlTab2(tabN).Container = SSTab2(tabN)
        ctlScrollControlTab2(tabN).Visible = True
        ctlScrollControlTab2(tabN).Left = 25
        SSTab2(tabN).Tab = 3
        DoEvents
        Load ctlScrollControlTab3(tabN)
        Set ctlScrollControlTab3(tabN).Container = SSTab2(tabN)
        ctlScrollControlTab3(tabN).Visible = True
        ctlScrollControlTab3(tabN).Left = 25
        SSTab2(tabN).Tab = 4
        DoEvents
        Load ctlScrollControlTab4(tabN)
        Set ctlScrollControlTab4(tabN).Container = SSTab2(tabN)
        ctlScrollControlTab4(tabN).Visible = True
        ctlScrollControlTab4(tabN).Left = 25
        'SSTab2(tabN).Tab = 0
        'DoEvents
    End If

ExitSub:
    arrOSList(tabN).CntBtn = lngCntBtn

    If lngOffSideCountTemp > lngOffSideCount Then
        lngOffSideCount = lngOffSideCountTemp
    End If

    On Error GoTo 0
    
    If mbDebugStandart Then DebugMode str2VbTab & "CreateButtonsonSSTab: cntButton=" & lngCntBtn & vbNewLine & _
              vbTab & "CreateButtonsonSSTab-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CreateMenuDevIDIndexCopyMenu
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strDevID (String)
'!--------------------------------------------------------------------------------
Private Sub CreateMenuDevIDIndexCopyMenu(ByVal strDevID As String)

    Dim i         As Long
    Dim ii        As Long
    Dim DevId_x() As String
    Dim strName   As String

    On Error Resume Next

    DevId_x = Split(strDevID, ";")

    ' Если меню уже заполнено, то удаляем его
    If mnuContextCopyHWID2Clipboard.Count > 1 Then

        For ii = mnuContextCopyHWID2Clipboard.LBound To mnuContextCopyHWID2Clipboard.UBound
            mnuContextCopyHWID2Clipboard(ii).Visible = False
            Unload mnuContextCopyHWID2Clipboard(ii)
        Next

        mnuContextCopyHWID2Clipboard(0).Visible = False
    End If

    mnuContextCopyHWID2Clipboard(0).Visible = False

    For ii = UBound(DevId_x) To 0 Step -1
        strName = DevId_x(ii)

        If Not mnuContextCopyHWID2Clipboard(0).Visible Then
            'если меню еще не создано
            mnuContextCopyHWID2Clipboard(0).Visible = True
            mnuContextCopyHWID2Clipboard(0).Caption = strName
        Else
            'NOT NOT...
            Load mnuContextCopyHWID2Clipboard(mnuContextCopyHWID2Clipboard.Count)
            mnuContextCopyHWID2Clipboard(mnuContextCopyHWID2Clipboard.Count - 1).Visible = True

            For i = mnuContextCopyHWID2Clipboard.UBound To mnuContextCopyHWID2Clipboard.LBound Step -1

                If i = mnuContextCopyHWID2Clipboard.LBound Then
                    mnuContextCopyHWID2Clipboard(0).Caption = strName

                    Exit For

                End If

                mnuContextCopyHWID2Clipboard(i).Caption = mnuContextCopyHWID2Clipboard(i - 1).Caption
            Next

        End If

    Next

    On Error GoTo 0

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CreateMenuDevIDIndexDelMenu
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strDevID (String)
'!--------------------------------------------------------------------------------
Private Sub CreateMenuDevIDIndexDelMenu(ByVal strDevID As String)

    Dim i         As Long
    Dim ii        As Long
    Dim DevId_x() As String
    Dim strName   As String

    On Error Resume Next

    DevId_x = Split(strDevID, ";")

    ' Если меню уже заполнено, то удаляем его
    If mnuContextDeleteDevID.Count > 1 Then

        For ii = mnuContextDeleteDevID.LBound To mnuContextDeleteDevID.UBound
            mnuContextDeleteDevID(ii).Visible = False
            Unload mnuContextDeleteDevID(ii)
        Next

        mnuContextDeleteDevID(0).Visible = False
    End If

    mnuContextDeleteDevID(0).Visible = False

    For ii = UBound(DevId_x) To 0 Step -1
        strName = DevId_x(ii)

        If Not mnuContextDeleteDevID(0).Visible Then
            'если меню еще не создано
            mnuContextDeleteDevID(0).Visible = True
            mnuContextDeleteDevID(0).Caption = strName
        Else
            'NOT NOT...
            Load mnuContextDeleteDevID(mnuContextDeleteDevID.Count)
            mnuContextDeleteDevID(mnuContextDeleteDevID.Count - 1).Visible = True

            For i = mnuContextDeleteDevID.UBound To mnuContextDeleteDevID.LBound Step -1

                If i = mnuContextDeleteDevID.LBound Then
                    mnuContextDeleteDevID(0).Caption = strName

                    Exit For

                End If

                mnuContextDeleteDevID(i).Caption = mnuContextDeleteDevID(i - 1).Caption
            Next

        End If

    Next

    On Error GoTo 0

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CreateMenuIndex
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strName (String)
'!--------------------------------------------------------------------------------
Private Sub CreateMenuIndex(ByVal strName As String)

    Dim i As Long

    On Error Resume Next

    If Not mnuUtils(0).Visible Then
        'если меню еще не создано
        mnuUtils(0).Visible = True
        mnuUtils(0).Caption = strName
    Else
        Load mnuUtils(mnuUtils.Count)
        mnuUtils(mnuUtils.Count - 1).Visible = True

        For i = mnuUtils.UBound To mnuUtils.LBound Step -1

            If i = mnuUtils.LBound Then
                mnuUtils(0).Caption = strName

                Exit For

            End If

            mnuUtils(i).Caption = mnuUtils(i - 1).Caption
        Next

    End If

    On Error GoTo 0

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CreateMenuLng
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strMenuCaption (String)
'!--------------------------------------------------------------------------------
Private Sub CreateMenuLng()
    Dim i As Long
    
    On Error Resume Next

    If Not mnuLang(0).Visible Then
        'если меню еще не создано
        mnuLang(0).Visible = True
    End If
    
    ' Создаем динамическое меню
    For i = UBound(arrLanguage, 2) - 1 To 1 Step -1
        Load mnuLang(mnuLang.Count)
        mnuLang(mnuLang.Count).Visible = True
        mnuLang(mnuLang.Count).Caption = "Lang " & mnuLang.Count
    Next i
    
    ' Присваиваем свойство Caption для меню
    For i = 0 To UBound(arrLanguage, 2)
        '4  mnuMainLang - "Язык"
        ' 2    mnuLang - "" - Index0 - Visible'False
        SetUniMenu 4, 2 + i, -1, mnuLang(i), arrLanguage(2, i + 1)
    Next i
        
    On Error GoTo 0

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CreateProgressNew
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub CreateProgressNew()

    With ctlProgressBar1
        .Min = 0
        .Max = 1000
        .Value = 0
        .Left = 0
        .Top = 0
        .Width = pbProgressBar.Width
        .Height = pbProgressBar.Height
        .SetTaskBarProgressState PrbTaskBarStateInProgress
        .SetTaskBarProgressValue .Value, 1000
    End With

    pbProgressBar.Visible = True
    pbProgressBar.Refresh
    DoEvents
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DelDuplicateOldDP
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub DelDuplicateOldDP()

    Dim ButtIndex                 As Long
    Dim i                         As Long
    Dim ii                        As Long
    Dim strPackFileName()         As String
    Dim strPackFileNames          As String
    Dim strPackFileName_woVersion As String
    Dim strPackFileNameTemp       As String
    Dim lngVersionPosition        As Long
    Dim strPackFileName_Ext       As String
    Dim objRegExp                 As RegExp
    Dim objMatch                  As Match
    Dim objMatches                As MatchCollection
    Dim strVerDP_1                As String
    Dim strVerDP_2                As String
    Dim strVerDP_1_1              As String
    Dim strVerDP_2_1              As String
    Dim strDPName_1               As String
    Dim strDPName_2               As String
    Dim strVerDP_Main             As String
    Dim strResult                 As String
    Dim strResult1                As String
    Dim strResult2                As String
    Dim strPackFileName2Del       As String
    Dim strPackFileName2DelTemp   As String
    Dim strPackFileName2Del_x()   As String
    Dim lngMsgRet                 As Long
    Dim lngStrLen1                As Long
    Dim lngStrLen2                As Long

    ButtIndex = acmdPackFiles.UBound

    ReDim strPackFileName(ButtIndex, 2)

    If ButtIndex Then

        For i = 0 To ButtIndex
            strPackFileName(i, 0) = acmdPackFiles(i).Tag
            strPackFileName(i, 1) = i

            If LenB(strPackFileNames) Then
                strPackFileNames = strPackFileNames & ";" & acmdPackFiles(i).Tag
            Else
                strPackFileNames = acmdPackFiles(i).Tag
            End If

        Next

    End If

    For i = LBound(strPackFileName, 1) To UBound(strPackFileName, 1)
        strPackFileNameTemp = strPackFileName(i, 0)

        If InStr(strPackFileNameTemp, vbBackslash) Then
            strPackFileNameTemp = GetFileNameFromPath(strPackFileName(i, 0))
        End If

        lngVersionPosition = InStrRev(strPackFileNameTemp, "_", , vbTextCompare)

        If lngVersionPosition Then
            strPackFileName_woVersion = Left$(strPackFileNameTemp, lngVersionPosition)
            strPackFileName_Ext = GetFileNameExtension(strPackFileNameTemp)
            Set objRegExp = New RegExp

            With objRegExp
                .Pattern = "(" & strPackFileName_woVersion & "([\d]+)[a-zA-Z]*([\d]*)." & strPackFileName_Ext & ")"
                .IgnoreCase = True
                .Global = True
                .MultiLine = True
                Set objMatches = .Execute(strPackFileNames)
            End With

            With objMatches

                If .Count > 1 Then
                    strVerDP_Main = vbNullString
                    strResult = vbNullString
                    strVerDP_1 = vbNullString
                    strVerDP_2 = vbNullString
                    strDPName_1 = vbNullString
                    strDPName_2 = vbNullString
                    ii = 0

                    Do While ii + 1 < .Count

                        If LenB(strVerDP_Main) = 0 Then
                            Set objMatch = .item(ii)
                            strVerDP_1 = Trim$(objMatch.SubMatches(1))
                            strDPName_1 = Trim$(objMatch.SubMatches(0))
                            strVerDP_1_1 = Trim$(objMatch.SubMatches(2))
                            Set objMatch = Nothing
                        Else
                            strVerDP_1 = strVerDP_Main
                            strDPName_1 = strDPName_2
                        End If

                        Set objMatch = .item(ii + 1)
                        strVerDP_2 = Trim$(objMatch.SubMatches(1))
                        strDPName_2 = Trim$(objMatch.SubMatches(0))
                        strVerDP_2_1 = Trim$(objMatch.SubMatches(2))
                        lngStrLen1 = Len(strVerDP_1)
                        lngStrLen2 = Len(strVerDP_2)

                        If lngStrLen1 > lngStrLen2 Then
                            strResult1 = CompareByVersion(Left$(strVerDP_1, lngStrLen2), strVerDP_2)
                            strResult = strResult1

                        ElseIf lngStrLen1 < lngStrLen2 Then
                            strResult1 = CompareByVersion(strVerDP_1, Left$(strVerDP_2, lngStrLen1))
                            strResult = strResult1

                        Else
                            strResult = CompareByVersion(strVerDP_1, strVerDP_2)

                            If StrComp(strResult, "=") = 0 Then
                                If LenB(strVerDP_1_1) Then
                                    If LenB(strVerDP_1_1) Then
                                        strResult2 = CompareByVersion(strVerDP_1_1, strVerDP_2_1)
                                    End If
                                End If

                                strResult = strResult2
                            End If
                        End If

                        If StrComp(strResult, ">") = 0 Then
                            strVerDP_Main = strVerDP_1
                            strPackFileName2DelTemp = strDPName_2
                        ElseIf StrComp(strResult, "<") = 0 Then
                            strVerDP_Main = strVerDP_2
                            strPackFileName2DelTemp = strDPName_1
                        End If

                        If LenB(strPackFileName2Del) Then
                            strPackFileName2Del = strPackFileName2Del & vbNewLine & strPackFileName2DelTemp
                        Else
                            strPackFileName2Del = strPackFileName2DelTemp
                        End If

                        ii = ii + 1
                        ' удаляем из списка пакетов, то что ранее уже проверяли
                        strPackFileNames = Replace$(strPackFileNames, strDPName_1, vbNullString, , , vbTextCompare)
                        strPackFileNames = Replace$(strPackFileNames, ";;", ";")
                        strPackFileNames = Replace$(strPackFileNames, strDPName_2, vbNullString, , , vbTextCompare)
                        strPackFileNames = Replace$(strPackFileNames, ";;", ";")
                    Loop

                End If

            End With

        End If

    Next

    ' Собственно удаление устаревших пакетов
    If LenB(strPackFileName2Del) Then
        If ShowMsbBoxForm(strPackFileName2Del, strMessages(139), strMessages(29)) = vbYes Then
            strPackFileName2Del_x = Split(strPackFileName2Del, vbNewLine)

            For i = LBound(strPackFileName2Del_x) To UBound(strPackFileName2Del_x)
                strPackFileName2DelTemp = strPackFileName2Del_x(i)

                For ii = 0 To ButtIndex

                    If StrComp(strPackFileName2DelTemp, acmdPackFiles(ii).Tag, vbTextCompare) = 0 Then
                        lngCurrentBtnIndex = ii
                        mnuContextDeleteDRP_Click
                    End If

                Next
            Next

            lngMsgRet = MsgBox(strMessages(141), vbQuestion + vbApplicationModal + vbYesNo, strProductName)
            mbRestartProgram = lngMsgRet = vbYes
        End If

    Else
        ChangeStatusTextAndDebug strMessages(140)
        MsgBox strMessages(140), vbInformation, strProductName
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DeleteUnUsedBase
'! Description (Описание)  :   [Очистка лишних файлов БД]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub DeleteUnUsedBase()

    Dim TabCount               As Long
    Dim i                      As Integer
    Dim ii                     As Integer
    Dim strPathDRP             As String
    Dim strPathDevDB           As String
    Dim strFileListTXT_x()     As FindListStruct
    Dim strFileListDRP_x()     As FindListStruct
    Dim strFileListTXT()       As String
    Dim strFileListDBExists    As String
    Dim strFileListDBNotExists As String
    Dim strDRPFilename         As String
    Dim strFileNameDB          As String
    Dim strFileNameDBHwid      As String
    Dim strFileNameDBIni       As String
    Dim lngFileDBVerIniSize    As Long
    Dim strFileDBVerIniPath    As String
    Dim strFileName2Del        As String

    If mbDebugStandart Then DebugMode "DeleteUnUsedBase-Start"

    If mbIsDriveCDRoom Then
        MsgBox strMessages(16), vbInformation, strProductName
    Else
        TabCount = SSTab1.Tabs

        ' В цикле обрабатываем все каталоги
        For i = 0 To TabCount - 1
            strPathDRP = arrOSList(i).drpFolderFull
            strPathDevDB = arrOSList(i).devIDFolderFull
            strFileListDBExists = vbNullString

            'Построение списка пакетов драйверов
            If Not mbDP_Is_aFolder Then
                strFileListDRP_x = SearchFilesInRoot(strPathDRP, "DP*.7z;DP*.zip", True, False)
            Else
                strFileListDRP_x = SearchFoldersInRoot(strPathDRP, "DP*")
            End If

            'Построение списка txt и ini файлов в каталоге БД
            strFileListTXT_x = SearchFilesInRoot(strPathDevDB, "*DP*.txt;*DP*.ini;*DP*.hwid;*DevDBVersions*.ini", False, False)

            ' Проверка на существование БД
            For ii = 0 To UBound(strFileListDRP_x)
                strDRPFilename = strFileListDRP_x(ii).Name

                If CheckExistDB(strPathDevDB, strDRPFilename) Then
                    If InStr(1, strDRPFilename, ".zip", vbTextCompare) Then
                        strFileNameDB = strPathDevDB & Replace$(strDRPFilename, ".zip", ".txt", , , vbTextCompare)
                    End If

                    If InStr(1, strDRPFilename, ".7z", vbTextCompare) Then
                        strFileNameDB = strPathDevDB & Replace$(strDRPFilename, ".7z", ".txt", , , vbTextCompare)
                    End If

                    strFileNameDBHwid = Replace$(strFileNameDB, ".txt", ".hwid", , , vbTextCompare)
                    strFileNameDBIni = Replace$(strFileNameDB, ".txt", ".ini", , , vbTextCompare)
                    AppendStr strFileListDBExists, strFileNameDB & vbTab & strFileNameDBHwid, vbTab

                    If PathExists(strFileNameDBIni) Then
                        strFileListDBExists = IIf(LenB(strFileListDBExists), strFileListDBExists & vbTab, vbNullString) & strFileNameDBIni
                    End If
                End If

            Next

            strFileDBVerIniPath = BackslashAdd2Path(strPathDevDB) & "DevDBVersions.ini"
            strFileListDBExists = IIf(LenB(strFileListDBExists), strFileListDBExists & vbTab, vbNullString) & strFileDBVerIniPath

            'Строим список удаляемых файлов для несуществующих пакетов
            For ii = 0 To UBound(strFileListTXT_x)

                If InStr(1, strFileListDBExists, strFileListTXT_x(ii).FullPath, vbTextCompare) = 0 Then
                    If PathExists(strFileListTXT_x(ii).FullPath) Then
                        strFileListDBNotExists = IIf(LenB(strFileListDBNotExists), strFileListDBNotExists & vbNewLine, vbNullString) & Replace$(strFileListTXT_x(ii).FullPath, strAppPath, vbNullString, , , vbTextCompare)
                        'Удаление секции о данном пакете из ini-файла
                        'IniDelAllKeyPrivate GetFileName_woExt(GetFileNameFromPath(strFileListTXT_x(0, ii))), strFileDBVerIniPath
                    End If
                End If

            Next
        Next

        ' Вывод итогового сообщения со списком удаляемых файлов и запросом на удаление
        If LenB(strFileListDBNotExists) Then
            ChangeStatusTextAndDebug strMessages(71)

            If ShowMsbBoxForm(strFileListDBNotExists, strMessages(28), strMessages(29)) = vbYes Then
                strFileListTXT = Split(strFileListDBNotExists, vbNewLine)

                'удаление файлов для несуществующих пакетов
                For ii = 0 To UBound(strFileListTXT)
                    strFileName2Del = PathCollect(strFileListTXT(ii))

                    If PathExists(strFileName2Del) Then
                        DeleteFiles strFileName2Del

                        'Удаление секции о данном пакете из ini-файла
                        For i = 0 To TabCount - 1
                            strPathDevDB = arrOSList(i).devIDFolderFull
                            strFileDBVerIniPath = PathCombine(strPathDevDB, "DevDBVersions.ini")
                            'если файл DevDBVersions.ini нулевого размера, то удаляем и его
                            lngFileDBVerIniSize = GetFileSizeByPath(strFileDBVerIniPath)

                            If lngFileDBVerIniSize Then
                                IniDelAllKeyPrivate GetFileName_woExt(GetFileNameFromPath(strFileListTXT(ii))), strFileDBVerIniPath
                            Else
                                If mbDebugStandart Then DebugMode str2VbTab & "DeleteUnUsedBase: Delete - file is zero = 0 bytes: " & strFileDBVerIniPath
                                DeleteFiles strFileDBVerIniPath
                            End If

                        Next

                    End If

                Next

            End If

        Else
            ChangeStatusTextAndDebug strMessages(30)
            MsgBox strMessages(30), vbInformation, strProductName
        End If
    End If

    If mbDebugStandart Then DebugMode "DeleteUnUsedBase-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub EditOrReadDPName
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   CurButtonIndex (Long)
'                              mbRead (Boolean = False)
'!--------------------------------------------------------------------------------
Private Sub EditOrReadDPName(ByVal CurButtonIndex As Long, Optional ByVal mbRead As Boolean = False)

    Dim strDRPFilename As String
    Dim strDPName      As String
    Dim strDPNameOld   As String
    Dim strDPNameMsg   As String

    If mbDebugDetail Then DebugMode str4VbTab & "EditOrReadDPName: CurButtonIndex=" & CurButtonIndex
    'Считываем текущее имя пакета из файла
    strDPName = vbNullString
    strDRPFilename = GetFileNameFromPath(acmdPackFiles(CurButtonIndex).Tag)
    strDPNameOld = acmdPackFiles(CurButtonIndex).Caption
    strDPName = IniStringPrivate("DPNames", strDRPFilename, strSysIni)

    ' Если такого значения в файле нет, то ничего не добавляем
    If StrComp(strDPName, "no_key") = 0 Then
        strDPName = strDPNameOld
    End If

    If mbRead Then
        If LenB(strDPName) Then
            If mbButtonTextUpCase Then
                acmdPackFiles(CurButtonIndex).Caption = UCase$(strDPName)
            Else
                acmdPackFiles(CurButtonIndex).Caption = strDPName
            End If

            If mbDebugDetail Then DebugMode str5VbTab & "EditOrReadDPName: Change Viewed Name: " & strDRPFilename & " on " & strDPName
        End If

    Else

        If mbIsDriveCDRoom Then
            If Not mbSilentRun Then
                MsgBox strMessages(16), vbInformation, strProductName
            End If

        Else
            ChangeStatusTextAndDebug strMessages(74) & strSpace & strDRPFilename
            strDPName = InputBox(strMessages(75) & strSpace & strDRPFilename, strMessages(76), strDPName)

            If LenB(strDPName) = 0 Then
                strDPName = vbNullString
            End If

            If StrComp(strDPNameOld, strDPName) <> 0 Then
                IniWriteStrPrivate "DPNames", strDRPFilename, strDPName, strSysIni
                ChangeStatusTextAndDebug strMessages(77) & strSpace & strDRPFilename

                If LenB(strDPName) = 0 Then
                    If LenB(strDPName) = 0 Then
                        strDPNameMsg = strDPNameOld
                        strDPName = strDPNameOld
                    Else
                        strDPNameMsg = GetFileNameFromPath(acmdPackFiles(CurButtonIndex).Tag)
                        strDPName = strDPNameMsg
                    End If

                Else
                    strDPNameMsg = strDPName
                End If

                acmdPackFiles(CurButtonIndex).Caption = strDPName
                MsgBox strMessages(32) & str2vbNewLine & "File: " & strDRPFilename & vbNewLine & "New Name: " & strDPNameMsg, vbInformation, strProductName
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub EventOnActivateForm
'! Description (Описание)  :   [Разные сообщения если нет поддерживаемых вкладок, или что-то нет так с пакетами]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub EventOnActivateForm()

    Dim lngMsgRet As Long

    ' Если пакетов нет вообще, то предлагаем пользователю посетить мой сайт
    If StrComp(acmdPackFiles(0).Container.Name, "frTabPanel", vbTextCompare) = 0 Then
        BlockControlInNoDP False

        With lblNoDPInProgram
            Set .Container = SSTab1
            .AutoSize = True
            .Left = 100

            ' Изменяем положение лабел
            Dim cntUnHideTab   As Long
            Dim miValue1       As Long
            Dim sngNum1        As Single
            Dim SSTabTabHeight As Long

            SSTabTabHeight = SSTab1.TabHeight
            cntUnHideTab = FindUnHideTab

            If cntUnHideTab Then
                sngNum1 = (cntUnHideTab + 1) / lngOSCountPerRow
                miValue1 = Round(sngNum1, 0)
            Else
                miValue1 = 1
            End If

            If sngNum1 = miValue1 Then
                .Top = (SSTab1.Height - .Height + (SSTabTabHeight * (miValue1))) / 2
                .Width = SSTab1.Width - 150 * (sngNum1 + 1)
            Else
                .Top = (SSTab1.Height - .Height + (SSTabTabHeight * (miValue1 + 1))) / 2
                .Width = SSTab1.Width - 150 * (sngNum1 + 1)
            End If

            .Visible = True
            .ZOrder 0
        End With

        DoEvents
        Form_Resize
        lngMsgRet = MsgBox(strMessages(108), vbYesNoCancel + vbQuestion, strProductName)

        Select Case lngMsgRet

            Case vbYes
                mnuDriverPacksOnMySite_Click

            Case vbNo
                mnuDriverPacks_Click
        End Select

    Else
        NoSupportOSorNoDevBD
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FindAndInstallPanel
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strArcDRPPath (String)
'                              strIniPath (String)
'                              strSection (String)
'                              lngNumberPanel (Long)
'                              strWorkPath (String)
'!--------------------------------------------------------------------------------
Private Function FindAndInstallPanel(ByVal strArcDRPPath As String, ByVal strIniPath As String, ByVal strSection As String, ByVal lngNumberPanel As Long, ByVal strWorkPath As String) As Boolean

    Dim lngTagFilesCount As Long
    Dim lngCommandsCount As Long
    Dim i                As Long
    Dim strPrefix        As String
    Dim strPrefixTag     As String
    Dim strPrefixCommand As String
    Dim strTemp          As String
    Dim strDPSROOT       As String
    Dim strOtherFile     As String
    Dim cmdString        As String

    If mbDebugStandart Then DebugMode "FindAndInstallPanel-Start" & vbNewLine & _
              "FindAndInstallPanel: strIniPath=" & strIniPath & vbNewLine & _
              "FindAndInstallPanel: strSection=" & strSection & vbNewLine & _
              "FindAndInstallPanel: lngNumberPanel=" & lngNumberPanel
    'exc_1_tagFiles = 3
    'exc_1_tagFile1 = "%SystemDrive%\ATICCC.ins"
    'exc_1_tagFile2 = "%DPSROOT%\D\G\A1\CCC\setup.exe"
    'exc_1_tagFile3 = "%SystemRoot%\system32\atidemgx.dll"
    'exc_1_commands = 2
    'exc_1_command1 = "%DPSROOT%\D\G\A1\CCC\setup.exe /s"
    'exc_1_command2 = "cmd.exe /c DEL /F /S /Q %DPSROOT%\D\G\A1\CCC\setup.exe"
    'exc_1_command1  = "%SystemDrive%\devcon.exe update %DPSROOT%\D\L\NV3\nvnetbus.inf "PCI\VEN_10DE&DEV_00DF&SUBSYS_*""
    strPrefix = "exc_" & lngNumberPanel & "_"
    strPrefixTag = strPrefix & "tagFile"
    strPrefixCommand = strPrefix & "command"
    strDPSROOT = BackslashAdd2Path(strWorkPath)
    ' Получаем число таговых файлов
    lngTagFilesCount = IniLongPrivate(strSection, strPrefixTag & "s", strIniPath)

    'Если число файлов больше 0 то обрабатываем дальше
    If lngTagFilesCount Then
        If lngTagFilesCount <> 9999 Then

            ' получаем список таговых файлов
            For i = 1 To lngTagFilesCount
                strTemp = IniStringPrivate(strSection, strPrefixTag & i, strIniPath)

                If StrComp(strTemp, "no_key") = 0 Then
                    GoTo ExitWithFalse
                End If

                'если в пути %DPSROOT% то заменям рабочим каталогом
                strTemp = Replace$(strTemp, "%DPSROOT%\", strDPSROOT, , , vbTextCompare)

                ' Если в пути есть переменные окружения, то заменяем их на нормальный путь
                If InStr(strTemp, strPercentage) Then
                    strTemp = GetEnviron(strTemp, True)
                End If

                If InStr(1, strTemp, strDPSROOT, vbTextCompare) Then
                    strOtherFile = Replace$(strTemp, strDPSROOT, vbNullString, , , vbTextCompare)
                    UnpackOtherFile strArcDRPPath, strDPSROOT, strOtherFile
                End If

                If PathExists(strTemp) = False Then
                    GoTo ExitWithFalse
                Else

                    If Not PathIsAFolder(strTemp) Then
                        GoTo NextTag
                    Else
                        GoTo ExitWithFalse
                    End If
                End If

NextTag:
            Next

            ' Получаем число таговых файлов
            lngCommandsCount = IniLongPrivate(strSection, strPrefixCommand & "s", strIniPath)

            'Если число комманд больше, 0 то обрабатываем дальше
            If lngCommandsCount Then
                If lngCommandsCount <> 9999 Then

                    ' получаем список Комманд на выполнение
                    For i = 1 To lngCommandsCount
                        strTemp = IniStringPrivate(strSection, strPrefixCommand & i, strIniPath)

                        If StrComp(strTemp, "no_key") = 0 Then
                            GoTo NextCommand
                        End If

                        'если в пути %DPSROOT% то заменям рабочим каталогом
                        strTemp = Replace$(strTemp, "%DPSROOT%\", strDPSROOT, , , vbTextCompare)
                        strTemp = Replace$(strTemp, "%DPSTMP%", strWorkTemp, , , vbTextCompare)
                        '%DPSTMP%
                        strTemp = Replace$(strTemp, "%SystemDrive%\devcon.exe", strDevConExePath, , , vbTextCompare)

                        ' Если в пути есть переменные окружения, то заменяем их на нормальный путь
                        If InStr(strTemp, strPercentage) Then
                            strTemp = GetEnviron(strTemp, True)
                        End If

                        'strCommands(i) = strTemp
                        cmdString = strTemp
                        ChangeStatusTextAndDebug strMessages(78) & " '" & strSection & "': " & cmdString

                        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
                            If Not mbSilentRun Then
                                MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
                            End If

                            ChangeStatusTextAndDebug strMessages(79) & strSpace & strSection
                            If mbDebugStandart Then DebugMode "Error on run : " & cmdString
                        End If

NextCommand:
                    Next

                End If

            Else
                GoTo ExitWithFalse
            End If
        End If
    End If

    ' Успешное завершение функции
    FindAndInstallPanel = True
    If mbDebugStandart Then DebugMode "FindAndInstallPanel-End"

    Exit Function

    ' Аварийный выход
ExitWithFalse:
    FindAndInstallPanel = False
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FindCheckCount
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbMsgStatus (Boolean = True)
'!--------------------------------------------------------------------------------
Private Function FindCheckCount(Optional ByVal mbMsgStatus As Boolean = True) As Long

    Dim i       As Integer
    Dim miCount As Integer

    For i = acmdPackFiles.LBound To acmdPackFiles.UBound

        If chkPackFiles(i).Value Then
            miCount = miCount + 1
        End If

    Next

    With cmdRunTask
    
        If optRezim_Upd.Value Then
            .Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdRunTask", .Caption)
        Else
            .Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdRunTask1", .Caption)
        End If
    
        If mbLoadAppEnd Then
            If optRezim_Upd.Value Then
                ctlUcStatusBar1.PanelText(1) = strMessages(128)
            Else
                If Not mbOnlyUnpackDP Then
                    ctlUcStatusBar1.PanelText(1) = strMessages(129)
                Else
                    ctlUcStatusBar1.PanelText(1) = strMessages(155)
                End If
            End If
    
            If miCount Then
                .Caption = .Caption & " (" & miCount & ")   "
    
                If mbMsgStatus Then
                    ChangeStatusTextAndDebug strMessages(104) & strSpace & miCount, , False
                End If
    
            Else
    
                If mbMsgStatus Then
                    ChangeStatusTextAndDebug strMessages(105), , False
                End If
            End If
        End If
    End With

    FindCheckCount = miCount
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FindHwidInBaseNew
'! Description (Описание)  :   [Поиск вхождения Hwids в БД]
'! Parameters  (Переменные):   strDevDBPath (String)
'                              strPackFileName (String)
'                              lngButtonIndex (Long)
'!--------------------------------------------------------------------------------
Private Function FindHwidInBaseNew(ByVal strDevDBPath As String, ByVal strPackFileName As String, ByVal lngButtonIndex As Long) As String

    Dim i                        As Long
    Dim ii                       As Long
    Dim iii                      As Long
    Dim strFind                  As String
    Dim strFindMachID            As String
    Dim strFindCompatIDTemp      As String
    Dim strFindCompatID_x()      As String
    Dim strFindCompatIDFind      As String
    Dim strFileNameDevDB         As String
    Dim strPathFileNameDevDB     As String
    Dim strPathFileNameDevDBHwid As String
    Dim strLineAll               As String
    Dim strAll                   As String
    Dim strTemp                  As String
    Dim strDevID                 As String
    Dim strDevIDOrig             As String
    Dim strPathInf               As String
    Dim strDevVer                As String
    Dim strDevVerLocal           As String
    Dim strDevStatus             As String
    Dim strDevName               As String
    Dim strSection               As String
    Dim lngMaxLengthRow1         As Long
    Dim lngMaxLengthRow2         As Long
    Dim lngMaxLengthRow4         As Long
    Dim lngMaxLengthRow5         As Long
    Dim lngMaxLengthRow6         As Long
    Dim lngMaxLengthRow9         As Long
    Dim lngMaxLengthRow13        As Long
    Dim lngMaxLengthRowAllLine   As Long
    Dim strTTipLocalArr()        As String
    Dim lngTTipLocalArrCount     As Long
    Dim miMaxCountArr            As Long
    Dim strPriznakSravnenia      As String
    Dim strHwidToDel             As String
    Dim strHwidToDelLine         As String
    Dim lngMatchesCount          As Long
    Dim lngBuffer                As Long
    Dim lngBuffer2               As Long
    Dim lngFileStartFromSymbol   As Long
    Dim strFileFullText          As String
    Dim strFileFullTextHwid      As String
    Dim lngDriverScore           As Long
    Dim lngDriverScorePrev       As Long
    Dim strSectionUnsupported    As String
    Dim strCatFileExists         As String
    Dim TimeScriptRun            As Long
    Dim TimeScriptFinish         As Long
    Dim strFile_x()              As String
    Dim strFileFull_x()          As String
    Dim strResult_x()            As String
    Dim strResultByTab_x()       As String

    If mbDebugStandart Then DebugMode str4VbTab & "FindHwidInBaseNew: strPackFileName=" & strPackFileName
              
    TimeScriptRun = GetTickCount
    mbStatusNewer = False
    mbStatusOlder = False
    mbStatusHwid = True
    mbIgnorStatusHwid = False
    mbDRVNotInstall = False
    strFileNameDevDB = Replace$(strPackFileName, ".7Z", ".TXT")
                    
    If InStr(strFileNameDevDB, ".ZIP") Then
        strFileNameDevDB = Replace$(strFileNameDevDB, ".ZIP", ".TXT")
    End If

    If Not mbDP_Is_aFolder Then
        strPathFileNameDevDB = PathCombine(strDevDBPath, strFileNameDevDB)
        strPathFileNameDevDBHwid = Replace$(strPathFileNameDevDB, ".TXT", ".HWID")
    Else
        strPathFileNameDevDB = PathCombine(strDevDBPath, strPackFileName & ".TXT")
        strPathFileNameDevDBHwid = Replace$(strPathFileNameDevDB, ".TXT", ".HWID")
    End If

    If PathExists(strPathFileNameDevDB) Then
        If Not PathIsAFolder(strPathFileNameDevDB) Then
            ' Считываем содержимое всего файла индекса в буфер
            Erase strFileFull_x
            strFileFullText = FileReadData(strPathFileNameDevDB)
            strFileFull_x = Split(strFileFullText, vbNewLine)
            
            ' Считываем содержимое всего файла HWID в буфер
            Erase strFile_x
            strFileFullTextHwid = FileReadData(strPathFileNameDevDBHwid)
            strFile_x = Split(strFileFullTextHwid, vbNewLine)
            
            miMaxCountArr = 100

            ReDim strTTipLocalArr(11, miMaxCountArr)

            lngMaxLengthRow1 = lngTableHwidHeader1
            lngMaxLengthRow2 = lngTableHwidHeader2
            lngMaxLengthRow4 = lngTableHwidHeader4
            lngMaxLengthRow5 = lngTableHwidHeader5
            lngMaxLengthRow6 = lngTableHwidHeader6
            lngMaxLengthRow9 = lngTableHwidHeader9
            lngMaxLengthRow13 = lngTableHwidHeader13
            maxSizeRowAllLine = 0

            For i = 0 To UBound(arrHwidsLocal)
                strFind = arrHwidsLocal(i).HWIDCutting
                strFindCompatIDTemp = arrHwidsLocal(i).HWIDCompat

                ' Быстрый поиск совпадений в массиве
                lngBuffer = BinarySearch(strFile_x(), strFind)
                
                If mbDebugDetail Then DebugMode str5VbTab & "FindHwidInBaseNew: PreFind by HWID: " & strFind & " =" & lngBuffer
                lngFileStartFromSymbol = lngBuffer

                If lngBuffer < 0 Then
                    ' заносим HWID в индекс, чтобы лишний раз потом его не проверять
                    objHashOutput3.RemoveAll

                    ' подходящие HWID (т.е обозначение HWID просто может быть записано по другому)
                    strFindMachID = arrHwidsLocal(i).HWIDMatches

                    If LenB(strFindMachID) Then
                        If StrComp(strFind, strFindMachID) <> 0 Then
                            If InStr(strFindMachID, "UNKNOWN") = 0 Then
                                If Not MatchSpec(strFindMachID, strExcludeHWID) Then
                                    If InStr(strFindMachID & " | ", strFindCompatIDTemp) = 0 Then
                                        strFindCompatIDTemp = strFindCompatIDTemp & " | " & strFindMachID
                                    End If
                                End If
                            End If
                        End If
                    End If

                    ' Поиск по совместимым HWID
                    If mbCompatiblesHWID Then
                        If InStr(strFindCompatIDTemp, "UNKNOWN") = 0 Then
                            If LenB(strFindCompatIDTemp) Then
                                strFindCompatID_x = Split(strFindCompatIDTemp, " | ")
                            End If

                        Else
                            GoTo NextStrFind
                        End If

                        strFind = vbNullString

                        For iii = 0 To UBound(strFindCompatID_x)

                            'Глубина поиска HWID
                            If iii > lngCompatiblesHWIDCount Then
                                Exit For
                            End If

                            strFindCompatIDFind = strFindCompatID_x(iii)
                            
                            If Not MatchSpec(strFindCompatIDFind, strExcludeHWID) Then

                                If Not objHashOutput3.Exists(strFindCompatIDFind) Then
                                    objHashOutput3.item(strFindCompatIDFind) = "+"
                                    lngBuffer2 = 0
                                    lngBuffer2 = BinarySearch(strFile_x(), strFindCompatIDFind)
                                    If mbDebugDetail Then DebugMode str5VbTab & "FindHwidInBaseNew: ***PreFind by HWID-Compatibles: " & strFindCompatIDFind & " =" & lngBuffer2
                                    lngFileStartFromSymbol = lngBuffer2

                                    If lngBuffer2 >= 0 Then
                                        strFind = strFindCompatIDFind
                                        lngDriverScore = iii
                                        GoTo ExitFromForNext_iii
                                    End If
                                End If
                            End If

                        Next iii

                        If LenB(strFind) = 0 Then
                            GoTo NextStrFind
                        End If

                    Else
                        GoTo NextStrFind
                    End If

                Else
                    lngDriverScore = -1
                End If

ExitFromForNext_iii:

                If lngFileStartFromSymbol < 0 Then
                    If mbDebugStandart Then DebugMode str5VbTab & "FindHwidInBaseNew: !!!ERROR lngFileStartFromSymbol=0 " & (strPackFileName & vbBackslash & strPathInf) & " by HWID=" & strFind
                    GoTo NextStrFind
                End If

                Erase strResult_x
                strResult_x = Filter(strFileFull_x(), strFind & vbTab, True, vbBinaryCompare)
                lngMatchesCount = UBound(strResult_x)

                If lngMatchesCount >= 0 Then
                    If mbDebugStandart Then DebugMode str5VbTab & "FindHwidInBaseNew: !!!Find " & lngMatchesCount & " Match in: " & (strPackFileName & vbBackslash & strPathInf) & " by HWID=" & strFind

                    For ii = 0 To lngMatchesCount
                        strResultByTab_x = Split(strResult_x(ii), vbTab)
                        ' Получаем имя секции файла inf для дальнейшего анализа
                        strPathInf = strResultByTab_x(1)
                        strSection = strResultByTab_x(2)
                        ' получение списка секций несовместимых ОС
                        strSectionUnsupported = strResultByTab_x(4)

                        ' Если драйвер несовместим с текущей ОС (вкладкой), то пропускаем его (анализ имени секции manufactured)
                        If Not CompatibleDriver4OS(strSection, strPackFileName, strPathInf, strSectionUnsupported) Then
                            If mbDebugStandart Then DebugMode str6VbTab & ii & " FindHwidInBaseNew: !!! SKIP. Driver is not compatible for this OS - IniSection: " & strSection & " Inf: " & strPathInf
                            GoTo NextLngMatchesCount
                        End If

                        strDevID = strResultByTab_x(0)
                        
                        If StrComp(strDevID, strFind, vbBinaryCompare) <> 0 Then
                            If mbDebugStandart Then DebugMode str6VbTab & ii & " FindHwidInBaseNew: ***Seeking HWID included in found HWID from database: HWID=" & strDevID
                            GoTo NextLngMatchesCount
                        End If
                        
                        strCatFileExists = strResultByTab_x(5)

                        If mbCalcDriverScore Then
                            ' Проверка и присвоение баллов драйверов
                            ' Если до этого оценок не было, то добавляем в базу
                            If mbDebugStandart Then DebugMode str6VbTab & ii & " FindHwidInBaseNew: ***Driver find in : " & (strPackFileName & vbBackslash & strPathInf) & " Has Score=" & lngDriverScore

                            If arrHwidsLocal(i).DRVScore = 0 Then
                                arrHwidsLocal(i).DRVScore = lngDriverScore
                            Else
                                lngDriverScorePrev = arrHwidsLocal(i).DRVScore

                                If lngDriverScore > lngDriverScorePrev Then
                                    If mbDebugStandart Then DebugMode str7VbTab & ii & " FindHwidInBaseNew: ***Driver is WORSE than found previously: ScoredPrev=" & lngDriverScorePrev
                                    GoTo NextLngMatchesCount
                                Else
                                    arrHwidsLocal(i).DRVScore = lngDriverScore
                                    If mbDebugStandart Then DebugMode str7VbTab & ii & " FindHwidInBaseNew: ***Added! Driver is BETTER OR EQUAL than found previously: ScoredPrev=" & lngDriverScorePrev
                                End If
                            End If
                        End If

                        strDevVer = strResultByTab_x(3)

                        ' если необходимо конвертировать дату в формат dd/mm/yyyy
                        If mbDateFormatRus Then
                            ConvertVerByDate strDevVer
                        End If

                        strDevVerLocal = arrHwidsLocal(i).VerLocal

                        If LenB(strDevVerLocal) = 0 Then
                            strDevVerLocal = "unknown"
                        End If

                        strDevName = strResultByTab_x(6)

                        If arrHwidsLocal(i).Status = 0 Then
                            mbStatusHwid = False

                            If InStr(strDevVerLocal, "unknown") = 0 Then
                                If LenB(strDevVerLocal) Then
                                    mbIgnorStatusHwid = True
                                End If
                            End If
                        End If

                        If LenB(strDevVerLocal) Then
                            strPriznakSravnenia = vbNullString

                            If mbCompareDrvVerByDate Then
                                strPriznakSravnenia = CompareByDate(strDevVer, strDevVerLocal)
                            Else
                                strPriznakSravnenia = CompareByVersion(strDevVer, strDevVerLocal)
                            End If

                            If StrComp(strPriznakSravnenia, ">") = 0 Then
                                ' В БД новее
                                mbStatusNewer = True
                                mbStatusOlder = False
                            ElseIf StrComp(strPriznakSravnenia, "<") = 0 Then
                                ' В БД старее
                                If Not mbStatusOlder Then
                                    If Not mbStatusNewer Then
                                        mbStatusOlder = True
                                    End If
                                End If
                                ' Дрова равны
                            End If

                            arrHwidsLocal(i).PriznakSravnenia = strPriznakSravnenia
                        Else
                            strPriznakSravnenia = "?"

                            If arrHwidsLocal(i).Status = 0 Then
                                mbDRVNotInstall = True
                                strPriznakSravnenia = ">"
                            End If
                        End If

                        strDevStatus = arrHwidsLocal(i).Status
                        arrHwidsLocal(i).InfSection = strSection
                        strDevIDOrig = ParseDoubleHwid(arrHwidsLocal(i).HWIDOrig)
                        ' Для этого драйвера есть совпадение в пакете драйверов. Заносим признак и имя пакета
                        AppendStr arrHwidsLocal(i).DPsList, strPackFileName, " | "

                        ' Если записей в массиве становится больше чем объявлено, то увеличиваем размерность массива
                        If lngTTipLocalArrCount = miMaxCountArr Then
                            miMaxCountArr = 2 * miMaxCountArr

                            ReDim Preserve strTTipLocalArr(11, miMaxCountArr)

                        End If

                        ' Заносим данные во временный массив
                        strTTipLocalArr(0, lngTTipLocalArrCount) = strDevID
                        strTTipLocalArr(1, lngTTipLocalArrCount) = strPathInf
                        strTTipLocalArr(2, lngTTipLocalArrCount) = strDevVer
                        strTTipLocalArr(3, lngTTipLocalArrCount) = strDevVerLocal
                        strTTipLocalArr(4, lngTTipLocalArrCount) = strDevStatus
                        strTTipLocalArr(5, lngTTipLocalArrCount) = strDevName
                        strTTipLocalArr(6, lngTTipLocalArrCount) = strPriznakSravnenia
                        strTTipLocalArr(7, lngTTipLocalArrCount) = strSection
                        strTTipLocalArr(8, lngTTipLocalArrCount) = strDevIDOrig
                        strTTipLocalArr(9, lngTTipLocalArrCount) = lngDriverScore
                        strTTipLocalArr(10, lngTTipLocalArrCount) = strSectionUnsupported
                        strTTipLocalArr(11, lngTTipLocalArrCount) = strCatFileExists
                        lngTTipLocalArrCount = lngTTipLocalArrCount + 1

                        If mbFirstStart Then

                            ' Заносим данные в глобальный массив массив
                            ' Если записей в массиве становится больше чем объявлено, то увеличиваем размерность массива
                            If lngArrDriversIndex = lngArrDriversListCountMax Then
                                lngArrDriversListCountMax = lngArrDriversListCountMax + lngArrDriversListCountMax

                                ReDim Preserve arrDriversList(12, lngArrDriversListCountMax)

                            End If

                            arrDriversList(0, lngArrDriversIndex) = strDevID
                            arrDriversList(1, lngArrDriversIndex) = strPathInf
                            arrDriversList(2, lngArrDriversIndex) = strDevVer
                            arrDriversList(3, lngArrDriversIndex) = strDevVerLocal
                            arrDriversList(4, lngArrDriversIndex) = strDevStatus
                            arrDriversList(5, lngArrDriversIndex) = strDevName
                            arrDriversList(6, lngArrDriversIndex) = strPriznakSravnenia
                            arrDriversList(7, lngArrDriversIndex) = strSection
                            arrDriversList(8, lngArrDriversIndex) = strDevIDOrig
                            arrDriversList(9, lngArrDriversIndex) = lngDriverScore
                            arrDriversList(10, lngArrDriversIndex) = strSectionUnsupported
                            arrDriversList(11, lngArrDriversIndex) = strCatFileExists
                            arrDriversList(12, lngArrDriversIndex) = strPackFileName
                            lngArrDriversIndex = lngArrDriversIndex + 1
                        End If

                        'Устанавливаем ширину колонок в таблице
                        If Len(strDevID) > lngMaxLengthRow1 Then
                            lngMaxLengthRow1 = Len(strDevID)
                        End If

                        If Len(strPathInf) > lngMaxLengthRow2 Then
                            lngMaxLengthRow2 = Len(strPathInf)
                        End If

                        If Len(strDevVer) > lngMaxLengthRow4 Then
                            lngMaxLengthRow4 = Len(strDevVer)
                        End If

                        If Len(strDevVerLocal) > lngMaxLengthRow5 Then
                            lngMaxLengthRow5 = Len(strDevVerLocal)
                        End If

                        If Len(strDevStatus) > lngMaxLengthRow6 Then
                            lngMaxLengthRow6 = Len(strDevStatus)
                        End If

                        If Len(strSection) > lngMaxLengthRow13 Then
                            lngMaxLengthRow13 = Len(strSection)
                        End If

NextLngMatchesCount:

                    Next ii

                Else
                    If mbDebugDetail Then DebugMode str5VbTab & "FindHwidInBaseNew: !!!ERROR Driver NOT find by Regexp in : " & (strPackFileName & vbBackslash & strPathInf) & " by HWID=" & strFind
                End If

NextStrFind:

            Next i

            If lngTTipLocalArrCount Then
                ' Обнуляем индексы. Нужны чтобы убрать повторяющиеся строки в подсказках
                objHashOutput.RemoveAll
                objHashOutput2.RemoveAll

                ReDim Preserve strTTipLocalArr(11, lngTTipLocalArrCount - 1)

                For i = 0 To UBound(strTTipLocalArr, 2)
                    'strDevID
                    strTemp = strTTipLocalArr(0, i)
                    strTTipLocalArr(0, i) = strTemp & Space$(lngMaxLengthRow1 - Len(strTemp) + 1) & "| "
                    'strPathInf
                    strTemp = strTTipLocalArr(1, i)
                    strTTipLocalArr(1, i) = strTemp & Space$(lngMaxLengthRow2 - Len(strTemp) + 1) & "| "
                    'strDevVer
                    strTemp = strTTipLocalArr(2, i)
                    strTTipLocalArr(2, i) = strTemp & Space$(lngMaxLengthRow4 - Len(strTemp) + 1) & "| "
                    'strDevVerLocal
                    strTemp = strTTipLocalArr(3, i)
                    strTTipLocalArr(3, i) = strTemp & Space$(lngMaxLengthRow5 - Len(strTemp) + 1) & "| "
                    ' strPriznakSravnenia
                    strTemp = strTTipLocalArr(6, i)
                    strTTipLocalArr(6, i) = strTemp & Space$(lngMaxLengthRow9 - Len(strTemp) + 1) & "| "
                    'strDevStatus & strDevName
                    strTemp = strTTipLocalArr(4, i)
                    strTTipLocalArr(4, i) = strTemp & Space$(lngMaxLengthRow6 - Len(strTemp) + 1) & "| "
                    ' Секция
                    strTemp = strTTipLocalArr(7, i)
                    strTTipLocalArr(7, i) = strTemp & Space$(lngMaxLengthRow13 - Len(strTemp) + 1) & "|"
                    ' Итоговый
                    strLineAll = strTTipLocalArr(0, i) & strTTipLocalArr(1, i) & strTTipLocalArr(2, i) & strTTipLocalArr(6, i) & strTTipLocalArr(3, i) & strTTipLocalArr(4, i) & strTTipLocalArr(5, i)

                    If Not objHashOutput.Exists(strLineAll) Then
                        objHashOutput.item(strLineAll) = "+"
                        AppendStr strAll, strLineAll, vbNewLine
                    End If

                    ' Заполняем массив для удаления драйверов по HWID
                    strHwidToDelLine = strTTipLocalArr(8, i)

                    If Not objHashOutput2.Exists(strHwidToDelLine) Then
                        objHashOutput2.item(strHwidToDelLine) = "+"
                        AppendStr strHwidToDel, strHwidToDelLine & vbTab & strTTipLocalArr(5, i), ";"
                    End If

                    ' Подсчитываем максимальную длину строки в подсказке
                    If Len(strLineAll) > lngMaxLengthRowAllLine Then
                        lngMaxLengthRowAllLine = Len(strLineAll)
                    End If

                Next i

            End If

            If lngSizeRowDPMax < Len(strPackFileName) Then
                lngSizeRowDPMax = Len(strPackFileName)
            End If

            lngSizeRow1 = lngMaxLengthRow1

            If lngSizeRow1Max < lngSizeRow1 Then
                lngSizeRow1Max = lngSizeRow1
            End If

            lngSizeRow2 = lngMaxLengthRow2

            If lngSizeRow2Max < lngSizeRow2 Then
                lngSizeRow2Max = lngSizeRow2
            End If

            lngSizeRow4 = lngMaxLengthRow4

            If lngSizeRow4Max < lngSizeRow4 Then
                lngSizeRow4Max = lngSizeRow4
            End If

            lngSizeRow5 = lngMaxLengthRow5

            If lngSizeRow5Max < lngSizeRow5 Then
                lngSizeRow5Max = lngSizeRow5
            End If

            lngSizeRow6 = lngMaxLengthRow6

            If lngSizeRow6Max < lngSizeRow6 Then
                lngSizeRow6Max = lngSizeRow6
            End If

            lngSizeRow9 = lngMaxLengthRow9

            If lngSizeRow9Max < lngSizeRow9 Then
                lngSizeRow9Max = lngSizeRow9
            End If

            lngSizeRow13 = lngMaxLengthRow13

            If lngSizeRow13Max < lngSizeRow13 Then
                lngSizeRow13Max = lngSizeRow13
            End If

            maxSizeRowAllLine = lngMaxLengthRowAllLine

            If maxSizeRowAllLineMax < maxSizeRowAllLine Then
                maxSizeRowAllLineMax = maxSizeRowAllLine
            End If
        End If
    End If

    FindHwidInBaseNew = strAll
    arrDevIDs(lngButtonIndex) = strHwidToDel

    ReDim Preserve arrTTipSize(lngButtonIndex + 1)

    arrTTipSize(lngButtonIndex) = maxSizeRowAllLine & (";" & lngSizeRow1 & ";" & lngSizeRow2 & ";" & lngSizeRow4 & ";" & lngSizeRow9 & ";" & lngSizeRow5 & ";" & lngSizeRow6)

ExitFromSub:
   
    TimeScriptFinish = GetTickCount
    If mbDebugStandart Then DebugMode str4VbTab & "FindHwidInBaseNew-Time to Find by HWID - " & strPackFileName & ": " & CalculateTime(TimeScriptRun, TimeScriptFinish, True)

    Exit Function

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FindNoDBCount
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function FindNoDBCount() As Long

    Dim miCount As Integer
    Dim i       As Integer

    With acmdPackFiles
        For i = .LBound To .UBound
    
            If Not (.item(i).PictureNormal Is Nothing) Then
                If .item(i).PictureNormal = imgNoDB.Picture Then
                    miCount = miCount + 1
                End If
            End If
    
        Next
    End With

    FindNoDBCount = miCount
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FindUnHideTab
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function FindUnHideTab() As Integer

    Dim miCount As Integer
    Dim i       As Integer

    miCount = -1

    For i = 0 To SSTab1.Tabs - 1

        If SSTab1.TabVisible(i) Then
            miCount = miCount + 1
        End If

    Next

    FindUnHideTab = miCount
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FontCharsetChange
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub FontCharsetChange()

    ' Выставляем шрифт
    With Me.Font
        .Name = strFontMainForm_Name
        .Size = lngFontMainForm_Size
        .Charset = lngFont_Charset
    End With

    frCheck.Font.Charset = lngFont_Charset
    frDescriptionIco.Font.Charset = lngFont_Charset
    frInfo.Font.Charset = lngFont_Charset
    frRezim.Font.Charset = lngFont_Charset
    frRunChecked.Font.Charset = lngFont_Charset
    frTabPanel.Font.Charset = lngFont_Charset
    ctlUcStatusBar1.Font.Charset = lngFont_Charset
    
    SetBtnFontProperties cmdRunTask
    SetBtnFontProperties cmdBreakUpdateDB
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FRMStateSave
'! Description (Описание)  :   [Запись положения форм в ini-файл]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub FRMStateSave()

    Dim miHeight      As Long
    Dim miWidth       As Long
    Dim miWindowState As Long

    ' Если настройка активна, то выполняем сохранение
    If Me.WindowState = vbMaximized Then
        miWindowState = 1
    Else
        miWindowState = 0
        miHeight = CLng(Me.Height)
        miWidth = vbNullString & CLng(Me.Width) & vbNullString
        IniWriteStrPrivate "MainForm", "Height", CStr(miHeight), strSysIni
        IniWriteStrPrivate "MainForm", "Width", CStr(miWidth), strSysIni
    End If

    IniWriteStrPrivate "MainForm", "StartMaximazed", CStr(miWindowState), strSysIni
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub GroupInstallDP
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub GroupInstallDP()

    Dim ButtIndex             As Long
    Dim miCheckDPCount        As Long
    Dim i                     As Long
    Dim strPackFileName       As String
    Dim strPathDRP            As String
    Dim strPathDevDB          As String
    Dim strPackGetFileName_woExt As String
    Dim ArchTempPath          As String
    Dim DPInstExitCode        As Long
    Dim strDevPathShort       As String
    Dim miCheckDPNumber       As Long
    Dim strPhysXPath          As String
    Dim strLangPath           As String
    Dim strRuntimes           As String
    Dim ReadExitCodeString    As String
    Dim miPbInterval          As Long
    Dim miPbNext              As Long
    Dim lngFindCheckCountTemp As Long
    Dim strTemp_x()           As String
    Dim strTempLine_x()       As String
    Dim i_arr                 As Long

    If mbDebugStandart Then DebugMode "GroupInstallDP-Start"
    ButtIndex = chkPackFiles.UBound
    miCheckDPCount = FindCheckCount
    BlockControl False

    If miCheckDPCount Then

        ReDim arrCheckDP(1, miCheckDPCount - 1)

        If ButtIndex Then
            miCheckDPNumber = 0

            ' Составляем список выделенных пакетов
            For i = 0 To ButtIndex

                ' Если галка стоит на кнопке, то обрабатываем эту кнопку
                If chkPackFiles(i).Value Then
                    If chkPackFiles(i).Left Then
                        ' Индексы отмеченых кнопок
                        arrCheckDP(0, miCheckDPNumber) = i
                        miCheckDPNumber = miCheckDPNumber + 1
                    End If
                End If

            Next

        ElseIf ButtIndex = 0 Then

            If acmdPackFiles(0).Visible Then
                miCheckDPNumber = 0

                ' Если галка стоит на кнопке, то обрабатываем эту кнопку
                If chkPackFiles(i).Value Then
                    If chkPackFiles(i).Left Then
                        ' Индексы отмеченых кнопок
                        arrCheckDP(0, 0) = 0
                        miCheckDPNumber = 1
                    End If
                End If

            Else

                If Not mbSilentRun Then
                    MsgBox strMessages(12), vbInformation, strProductName
                End If

                If mbDebugStandart Then DebugMode "GroupInstallDP-DpPack is not Exist"

                Exit Sub

            End If

        Else

            If Not mbSilentRun Then
                MsgBox strMessages(12), vbInformation, strProductName
            End If

            If mbDebugStandart Then DebugMode "GroupInstallDP-DpPack is not Exist"

            Exit Sub

        End If

        ' Получаем список извлекаемых масок
        ' если выборочная установка, то получаем список каталогов для распаковки
        If mbSelectInstall Then

            ' Если выборочная установка, то показываем форму выбора
            If IsFormLoaded("frmListHwid") = False Then
                frmListHwid.Show vbModal, Me
            Else
                frmListHwid.FormLoadDefaultParam
                frmListHwid.FormLoadAction
                frmListHwid.Show vbModal, Me
            End If

            ' если на форме нажали отмену или закрыли ее, то завершаем обработку
            If Not mbCheckDRVOk Then
                mbDevParserRun = False
                If Not mbOnlyUnpackDP Then
                    ChangeStatusTextAndDebug strMessages(82), strMessages(129)
                Else
                    ChangeStatusTextAndDebug strMessages(132), strMessages(155)
                End If

                Exit Sub

            End If

        Else

            ' иначе список строится сам
            For i = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)

                strPathDRPList = vbNullString
                strTemp_x = Split(arrTTip(arrCheckDP(0, i)), vbNewLine)

                For i_arr = 0 To UBound(strTemp_x)
                    strTempLine_x = Split(strTemp_x(i_arr), " | ")

                    If LenB(Trim$(strTemp_x(i_arr))) Then
                        strDevPathShort = Trim$(strTempLine_x(1))

                        ' Если данного пути нет в списке, то добавляем
                        If InStr(1, strPathDRPList, strDevPathShort, vbTextCompare) = 0 Then
                            AppendStr strPathDRPList, strDevPathShort, strSpace
                        End If
                    End If

                Next i_arr

                strPathDRPList = Trim$(strPathDRPList)

                ' Если по каким либо причинам список папок не получился, то извлекаем все.
                If LenB(strPathDRPList) = 0 Then
                    strPathDRPList = ALL_FILES
                End If

                arrCheckDP(1, i) = strPathDRPList
            Next

        End If

        ChangeStatusTextAndDebug strMessages(83)
        'Имя папки с распакованными драйверами
        strPathDRP = arrOSList(SSTab1.Tab).drpFolderFull
        strPathDevDB = arrOSList(SSTab1.Tab).devIDFolderFull
        ' Групповое извлечение файлов
        ArchTempPath = strWorkTempBackSL & "GroupInstall"

        If PathExists(ArchTempPath) Then
            DelRecursiveFolder (ArchTempPath)
        End If

        If mbOnlyUnpackDP Then

            '# Диалог выбора каталога
            With New CommonDialog
                .InitDir = strAppPathBackSL & "drivers"
                .DialogTitle = strMessages(131)
                .Flags = CdlBIFNewDialogStyle

                If .ShowFolder = True Then
                    ArchTempPath = .FileName
                Else
                    ChangeStatusTextAndDebug strMessages(132), strMessages(155)
                    mbDevParserRun = False
                    '# if user cancel #
                    Exit Sub
                End If

            End With

            If LenB(ArchTempPath) = 0 Then
                ChangeStatusTextAndDebug strMessages(132), strMessages(155)

                '# if user cancel #
                Exit Sub

            End If

            If mbDebugStandart Then DebugMode "StartBackUp: Destination=" & ArchTempPath
        End If

        mbBreakUpdateDBAll = False
        ' Отображаем ProgressBar
        CreateProgressNew
        cmdBreakUpdateDB.Visible = True
        DoEvents
        ' Начальные пременные прогрессбара
        lngFindCheckCountTemp = FindCheckCount

        If lngFindCheckCountTemp Then
            If mbUnpackAdditionalFile Then
                miPbInterval = 700 / lngFindCheckCountTemp
            Else
                miPbInterval = 1000 / lngFindCheckCountTemp
            End If
        End If

        miPbNext = 0

        ' собственно цикл распаковки
        For i = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)

            With acmdPackFiles(arrCheckDP(0, i))
                
                .Value = True
                

                ' Прерываем процесс распаковки
                If mbBreakUpdateDBAll Then
                    MsgBox strMessages(27) & vbNewLine & .Tag, vbInformation, strProductName

                    Exit For

                End If

                strPackFileName = .Tag
                strPackGetFileName_woExt = GetFileName_woExt(strPackFileName)

                If UnPackDPFile(strPathDRP, strPackFileName, arrCheckDP(1, i), ArchTempPath) = False Then
                    If Not mbSilentRun Then
                        MsgBox strMessages(13) & vbNewLine & strPackFileName, vbInformation, strProductName
                    End If
                End If

                If chkPackFiles(arrCheckDP(0, i)).Value Then
                    chkPackFiles(arrCheckDP(0, i)).Value = False
                End If

                
                '.Value = False
                
            End With

            miPbNext = miPbNext + miPbInterval

            If miPbNext > 1000 Then
                miPbNext = 1000
            End If

            With ctlProgressBar1
                .Value = miPbNext
                .SetTaskBarProgressValue miPbNext, 1000
            End With

            ChangeFrmMainCaption miPbNext
        Next

        If mbBreakUpdateDBAll Then
            GoTo BreakUnpack
        End If

        ' Распаковка дополнительных файлов для групповой установки
        If mbUnpackAdditionalFile Then
            If mbBreakUpdateDBAll Then
                MsgBox strMessages(27) & vbNewLine & strPhysXPath, vbInformation, strProductName
                GoTo BreakUnpack
            End If

            ' Распаковка strPhysXPath
            If LenB(arrOSList(SSTab1.Tab).PathPhysX) Then
                strPhysXPath = PathCollect(arrOSList(SSTab1.Tab).PathPhysX)
                UnPackDPFileAdd strPhysXPath, strPathDRP, ArchTempPath
            End If

            miPbNext = miPbNext + 100

            If miPbNext > 1000 Then
                miPbNext = 1000
            End If

            With ctlProgressBar1
                .Value = miPbNext
                .SetTaskBarProgressValue miPbNext, 1000
            End With

            ChangeFrmMainCaption miPbNext

            If mbBreakUpdateDBAll Then
                MsgBox strMessages(27) & vbNewLine & strLangPath, vbInformation, strProductName
                GoTo BreakUnpack
            End If

            ' Распаковка strLangPath
            If LenB(arrOSList(SSTab1.Tab).PathLanguages) Then
                strLangPath = PathCollect(arrOSList(SSTab1.Tab).PathLanguages)
                UnPackDPFileAdd strLangPath, strPathDRP, ArchTempPath
            End If

            miPbNext = miPbNext + 100

            If miPbNext > 1000 Then
                miPbNext = 1000
            End If

            With ctlProgressBar1
                .Value = miPbNext
                .SetTaskBarProgressValue miPbNext, 1000
            End With

            ChangeFrmMainCaption miPbNext

            If mbBreakUpdateDBAll Then
                MsgBox strMessages(27) & vbNewLine & strRuntimes, vbInformation, strProductName
                GoTo BreakUnpack
            End If

            ' Распаковка strRuntimes
            If LenB(arrOSList(SSTab1.Tab).PathRuntimes) Then
                strRuntimes = PathCollect(arrOSList(SSTab1.Tab).PathRuntimes)
                UnPackDPFileAdd strRuntimes, strPathDRP, ArchTempPath
            End If
        End If

        miPbNext = 1000

        With ctlProgressBar1
            .Value = miPbNext
            .SetTaskBarProgressValue miPbNext, 1000
        End With

        ChangeFrmMainCaption
        
BreakUnpack:

        If mbBreakUpdateDBAll Then
            cmdBreakUpdateDB.Visible = False
            ChangeStatusTextAndDebug strMessages(82)
            GoTo EndedSub
        Else
            cmdBreakUpdateDB.Visible = False
        End If

        pbProgressBar.Visible = False
        ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone

        ' Если не выставлена опция только распаковка, то делаем установку
        If Not mbOnlyUnpackDP Then
            ChangeStatusTextAndDebug strMessages(84)
            ' установка всех извлеченных драйверов
            ArchTempPath = strWorkTempBackSL & "GroupInstall"
            DPInstExitCode = RunDPInst(ArchTempPath)
            ReadExitCodeString = ReadExitCode(DPInstExitCode)

            If DPInstExitCode <> 0 Then
                If DPInstExitCode <> -2147483648# Then
                    If InStr(1, ReadExitCodeString, "Cancel or Nothing to Install", vbTextCompare) = 0 Then

                        ' Обрабатываем файл finish
                        If mbLoadFinishFile Then
                            For i = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)
                                strPackFileName = acmdPackFiles(arrCheckDP(0, i)).Tag
                                strPackGetFileName_woExt = GetFileName_woExt(strPackFileName)
                                ArchTempPath = strWorkTempBackSL & strPackGetFileName_woExt
                                WorkWithFinish strPathDRP, strPackFileName, ArchTempPath, arrCheckDP(1, i)
                            Next
                        End If

                        ' Обновление подсказки
                        For i = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)
                            strPackFileName = acmdPackFiles(arrCheckDP(0, i)).Tag
                            ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, CInt(arrCheckDP(0, i)), True
                        Next

                    End If
                End If
            End If

            ChangeStatusTextAndDebug strMessages(85) & strSpace & ReadExitCodeString
            If mbDebugStandart Then DebugMode "Install from : " & strPackFileName & " finish."

            For i = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)

                If chkPackFiles(arrCheckDP(0, i)).Value Then
                    chkPackFiles(arrCheckDP(0, i)).Value = False
                End If

            Next

            If PathExists(ArchTempPath) Then
                DelRecursiveFolder (ArchTempPath)
            End If

        Else
            ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
            ChangeStatusTextAndDebug strMessages(125) & strSpace & ArchTempPath

            If MsgBox(strMessages(125) & str2vbNewLine & strMessages(133), vbYesNo, strProductName) = vbYes Then
                RunUtilsShell ArchTempPath, False
            End If
        End If

        mbUnpackAdditionalFile = False
    Else

        If Not mbSilentRun Then
            MsgBox strMessages(14), vbInformation, strProductName
        End If

        If mbDebugStandart Then DebugMode "GroupInstallDP-DpPack is not Checked"
        ChangeStatusTextAndDebug strMessages(14)
    End If

EndedSub:
    BlockControl True
    If mbDebugStandart Then DebugMode "GroupInstallDP-End"
    FindCheckCount False
    mbBreakUpdateDBAll = False
    ChangeFrmMainCaption
    pbProgressBar.Visible = False
    ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub InitClipboard
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub InitClipboard()

    If GetOpenClipboardWindow() <> NO_CB_OPENED Then
        CloseClipboard
        SetClipboardViewer Me.hWnd
    End If

    strCBError(0) = "Clipboard open error!!!"
    strCBError(1) = "Not Clipboard BITMAP format available!!!"
    strCBError(2) = "Not Clipboard TEXT format available!!!"
    strCBError(3) = "Clipboard already opened by other application!!!"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub InsOrUpdSelectedDP
'! Description (Описание)  :   [Запуск процесс установки(или обновления БД) выделенных пакетов драйверов]
'! Parameters  (Переменные):   mbInstallMode (Boolean)
'!--------------------------------------------------------------------------------
Private Sub InsOrUpdSelectedDP(ByVal mbInstallMode As Boolean)

    If cmdRunTask.Enabled Then
        If mbInstallMode Then
            If optRezim_Upd.Value Then
                SelectStartMode 1, False
            End If

            mbGroupTask = True
            mbSelectInstall = False
            GroupInstallDP
            mbGroupTask = False
        Else

            If Not optRezim_Upd.Value Then
                SelectStartMode 3, False
            End If

            cmdRunTask_Click
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsFormLoaded
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   FormName (String)
'!--------------------------------------------------------------------------------
Private Function IsFormLoaded(FormName As String) As Boolean

    Dim i As Integer

    For i = 0 To Forms.Count - 1

        If Forms(i).Name = FormName Then
            IsFormLoaded = True

            Exit Function

        End If

    Next i

    IsFormLoaded = False
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lblOsInfoChange
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub lblOsInfoChange()

    Dim str64bit         As String
    Dim lblOsInfoCaption As String

    If mbIsWin64 Then
        str64bit = " x64 Edition"
    Else
        str64bit = " x86 Edition"
    End If

    lblOsInfoCaption = LocaliseString(strPCLangCurrentPath, strFormName, "lblOsInfo", lblOSInfo.Caption)
    lblOSInfo.Caption = lblOsInfoCaption & strSpace & OSInfo.Name & strSpace & " (" & OSInfo.VerFullwBuild & strSpace & OSInfo.ServicePack & ")" & str64bit
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadButton
'! Description (Описание)  :   [Загрузка кнопок при старте программы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadButton()

    Dim i                As Long
    Dim cnt              As Long
    Dim pbStart          As Long
    Dim pbDelta          As Long
    Dim strPathFolderDRP As String
    Dim strPathFolderDB  As String

    On Error Resume Next

    If mbDebugStandart Then DebugMode "LoadButton-Start"
    mbNextTab = False
    frTabPanel.Visible = False
    lngCntBtn = 0
    cnt = UBound(arrOSList)

    With ctlProgressBar1
        pbStart = .Value
        .SetTaskBarProgressState PrbTaskBarStateInProgress
        .SetTaskBarProgressValue pbStart, 1000
    End With

    If cnt Then
        pbDelta = (1000 - pbStart) / (cnt + 1)
    Else
        pbDelta = 1000 - pbStart
    End If

    i = 0
    ' массив со списком драйверов для будущего использования для каждой кнопки
    lngArrDriversIndex = 0
    lngArrDriversListCountMax = 100

    ReDim Preserve arrDriversList(12, lngArrDriversListCountMax)

    For i = 0 To cnt
        strPathFolderDRP = arrOSList(i).drpFolderFull
        strPathFolderDB = arrOSList(i).devIDFolderFull
        ChangeStatusTextAndDebug strMessages(69) & strSpace & strPathFolderDRP
        
        If mbDebugStandart Then DebugMode vbTab & "Analize Folder DRP: " & strPathFolderDRP & vbNewLine & _
                  vbTab & "Analize Folder DB: " & strPathFolderDB
                  
        pbProgressBar.Refresh

        If Not arrOSList(i).DPFolderNotExist Then
            ' Запуск процедуры создания кнопок на вкладке
            CreateButtonsOnSSTab strPathFolderDRP, strPathFolderDB, i, pbDelta
        Else
            SSTab1.TabEnabled(i) = False

            If mbTabHide Then
                SSTab1.TabVisible(i) = False
            End If

            ' грузим вкладки , но делаем скрытыми
            If i Then
                Load SSTab2(i)
                Set SSTab2(i).Container = SSTab1
                Load ctlScrollControl1(i)
                Set ctlScrollControl1(i).Container = SSTab2(i)
                SSTab2(i).Visible = False
            End If
        End If

        mbNextTab = True
        pbProgressBar.Refresh
    Next

    With ctlProgressBar1
        .Value = 1000
        .SetTaskBarProgressValue 1000, 1000
        .SetTaskBarProgressState PrbTaskBarStateNone
    End With

    pbProgressBar.Visible = False
    ChangeFrmMainCaption

    If optRezim_Upd.Value Then
        optRezim_Upd_Click
    End If

    If acmdPackFiles(0).Visible Then
        ChangeStatusTextAndDebug strMessages(86)
        If mbDebugStandart Then DebugMode "Create Buttons: True"
    Else

        If acmdPackFiles.Count <= 1 Then
            ChangeStatusTextAndDebug strMessages(87)
            If mbDebugStandart Then DebugMode "Create Buttons: False"
            mnuRezimBaseDrvUpdateALL.Enabled = False
        End If

        SSTab1.Enabled = False
    End If

    If mbDebugStandart Then DebugMode "LoadButton-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadCmdRunTask
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadCmdRunTask()

    With cmdRunTask
        .Enabled = False
        .SetPopupMenu mnuContextMenu3
        .DropDownEnable = True
        .DropDownSeparator = True
        .DropDownSymbol = 6
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadCmdViewAllDeviceCaption
'! Description (Описание)  :   [Изменение описания кнопки cmdViewAllDevice]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub LoadCmdViewAllDeviceCaption()
    lngNotFinedDriversInDP = CalculateUnknownDrivers

    If lngNotFinedDriversInDP Then
        cmdViewAllDevice.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdViewAllDevice", cmdViewAllDevice.Caption) & vbNewLine & strMessages(122) & strSpace & lngNotFinedDriversInDP
        cmdViewAllDevice.ForeColor = vbRed
    Else
        cmdViewAllDevice.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdViewAllDevice", cmdViewAllDevice.Caption)
        cmdViewAllDevice.ForeColor = cmdRunTask.ForeColor
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadIconImage
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadIconImage()
    '--------------------- Статусные Иконки
    LoadIconImage2Object imgNoDB, "BTN_NO_DB", strPathImageStatusButtonWork
    LoadIconImage2Object imgOK, "BTN_OK", strPathImageStatusButtonWork
    LoadIconImage2Object imgOkAttention, "BTN_OK_ATTENTION", strPathImageStatusButtonWork
    LoadIconImage2Object imgOkAttentionNew, "BTN_OK_ATTENTION_NEW", strPathImageStatusButtonWork
    LoadIconImage2Object imgOkAttentionOLD, "BTN_OK_ATTENTION_OLD", strPathImageStatusButtonWork
    LoadIconImage2Object imgOkNew, "BTN_OK_NEW", strPathImageStatusButtonWork
    LoadIconImage2Object imgOkOld, "BTN_OK_OLD", strPathImageStatusButtonWork
    LoadIconImage2Object imgNo, "BTN_NO_DRV", strPathImageStatusButtonWork
    LoadIconImage2Object imgUpdBD, "BTN_UPD_DRV", strPathImageStatusButtonWork
    '--------------------- Остальные Иконки
    LoadIconImage2Object cmdRunTask, "BTN_RUNTASK", strPathImageMainWork
    LoadIconImage2Object cmdBreakUpdateDB, "BTN_BREAK_UPDATE", strPathImageMainWork
    LoadIconImage2Object cmdViewAllDevice, "BTN_VIEW_SEARCH", strPathImageMainWork
    LoadIconImage2Object cmdCheck, "BTN_CHECKMARK", strPathImageMainWork
    '--------------------- Группы
    LoadIconImage2Object frRezim, "FRAME_GROUP", strPathImageMainWork
    If mbDebugStandart Then DebugMode "LoadIconImage-End"
End Sub

'заполнение списка на выделение
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadListChecked
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadListChecked()
    cmbCheckButton.Clear
    ' Режимы выделения
    strCmbChkBtnListElement1 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement1", "Все на текущей вкладке")
    strCmbChkBtnListElement2 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement2", "Сброс отметок")
    strCmbChkBtnListElement3 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement3", "Все")
    strCmbChkBtnListElement4 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement4", "Все новые")
    strCmbChkBtnListElement5 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement5", "Неустановленные")
    strCmbChkBtnListElement6 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement6", "Рекомендуемые")
    
    With cmbCheckButton
        If optRezim_Upd.Value Then
    
            .AddItem strCmbChkBtnListElement1, 0
            .AddItem strCmbChkBtnListElement2, 1
            .AddItem strCmbChkBtnListElement3, 2
            .AddItem strCmbChkBtnListElement4, 3
            .ListIndex = 3
            ' Подсчитываем кол-во пакетов не имеющих БД, и если их нет то ставим "Все новые"
            If FindNoDBCount = 0 Then .ListIndex = 2
    
        ElseIf optRezim_Ust.Value Then
    
            .AddItem strCmbChkBtnListElement2, 0
            .AddItem strCmbChkBtnListElement5, 1
            .AddItem strCmbChkBtnListElement6, 2
            .AddItem strCmbChkBtnListElement1, 3
            .ListIndex = 1
    
        Else
            .AddItem strCmbChkBtnListElement2, 0
            .AddItem strCmbChkBtnListElement5, 1
            .AddItem strCmbChkBtnListElement6, 2
            .AddItem strCmbChkBtnListElement1, 3
            .ListIndex = 1
    
        End If
    End With
        
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadSSTab2Desc
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadSSTab2Desc()

    Dim i As Long

    SetTabPropertiesTabDrivers

    With SSTab2

        For i = .LBound To .UBound
            .item(i).TabCaption(0) = strSSTabTypeDPTab1
            .item(i).TabCaption(1) = strSSTabTypeDPTab2
            .item(i).TabCaption(2) = strSSTabTypeDPTab3
            .item(i).TabCaption(3) = strSSTabTypeDPTab4
            .item(i).TabCaption(4) = strSSTabTypeDPTab5
        Next

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal strPathFile As String)
    ' изменяем шрифт
    FontCharsetChange
    'Frame
    frRezim.Caption = LocaliseString(strPathFile, strFormName, "frRezim", frRezim.Caption)
    frDescriptionIco.Caption = LocaliseString(strPathFile, strFormName, "frDescriptionIco", frDescriptionIco.Caption)
    frRunChecked.Caption = LocaliseString(strPathFile, strFormName, "frRunChecked", frRunChecked.Caption)
    frCheck.Caption = LocaliseString(strPathFile, strFormName, "frCheck", frCheck.Caption)
    frInfo.Caption = LocaliseString(strPathFile, strFormName, "frInfo", frInfo.Caption)
    ' Описание режимов
    optRezim_Intellect.Caption = LocaliseString(strPathFile, strFormName, "RezimIntellect", optRezim_Intellect.Caption)
    optRezim_Ust.Caption = LocaliseString(strPathFile, strFormName, "RezimUst", optRezim_Ust.Caption)
    optRezim_Upd.Caption = LocaliseString(strPathFile, strFormName, "RezimUpd", optRezim_Upd.Caption)
    ' Меню
    '  Вызов основной функции для вывода Caption меню с поддержкой Unicode
    Call LocaliseMenu(strPathFile)
    'Кнопки
    cmdRunTask.Caption = LocaliseString(strPathFile, strFormName, "cmdRunTask", cmdRunTask.Caption)
    cmdCheck.Caption = LocaliseString(strPathFile, strFormName, "cmdCheck", cmdCheck.Caption)
    cmdBreakUpdateDB.Caption = LocaliseString(strPathFile, strFormName, "cmdBreakUpdateDB", cmdBreakUpdateDB.Caption)
    cmdViewAllDevice.Caption = LocaliseString(strPathFile, strFormName, "cmdViewAllDevice", cmdViewAllDevice.Caption)
    ' Лейблы
    lblPCInfo.Caption = LocaliseString(strPathFile, strFormName, "lblPCInfo", lblPCInfo.Caption) & strSpace & strCompModel
    lblNoDP4Mode.Caption = LocaliseString(strPathFile, strFormName, "lblNoDP4Mode", lblNoDP4Mode.Caption)
    lblNoDPInProgram.Caption = LocaliseString(strPathFile, strFormName, "lblNoDPInProgram", lblNoDPInProgram.Caption)
    ' Другие параметры
    strTableHwidHeader1 = LocaliseString(strPathFile, strFormName, "TableHwidHeader1", "-HWID-")
    strTableHwidHeader2 = LocaliseString(strPathFile, strFormName, "TableHwidHeader2", "-Inf-Файл-")
    strTableHwidHeader4 = LocaliseString(strPathFile, strFormName, "TableHwidHeader4", "-Версия(БД)-")
    strTableHwidHeader5 = LocaliseString(strPathFile, strFormName, "TableHwidHeader5", "-Версия(PC)-")
    strTableHwidHeader6 = LocaliseString(strPathFile, strFormName, "TableHwidHeader6", "-Статус-")
    strTableHwidHeader7 = LocaliseString(strPathFile, strFormName, "TableHwidHeader7", "-Наименование устройства-")
    strTableHwidHeader8 = LocaliseString(strPathFile, strFormName, "TableHwidHeader8", "-Пакет драйверов-")
    strTableHwidHeader9 = LocaliseString(strPathFile, strFormName, "TableHwidHeader9", "-!-")
    strTableHwidHeader10 = LocaliseString(strPathFile, strFormName, "TableHwidHeader10", "-Производитель-")
    strTableHwidHeader11 = LocaliseString(strPathFile, strFormName, "TableHwidHeader11", "-Совместимый HWID-")
    strTableHwidHeader12 = LocaliseString(strPathFile, strFormName, "TableHwidHeader12", "-Код устройства-")
    strTableHwidHeader13 = LocaliseString(strPathFile, strFormName, "TableHwidHeader13", "-Секция-")
    strTableHwidHeader14 = LocaliseString(strPathFile, strFormName, "TableHwidHeader14", "Найден в пакете")
    strTTipTextTitle = LocaliseString(strPathFile, strFormName, "ToolTipTextTitle", "Файл пакета драйверов:")
    strTTipTextFileSize = LocaliseString(strPathFile, strFormName, "ToolTipTextFileSize", "Размер файла:")
    strTTipTextClassDRV = LocaliseString(strPathFile, strFormName, "ToolTipTextClassDRV", "Класс драйверов:")
    strTTipTextDrv2Install = LocaliseString(strPathFile, strFormName, "ToolTipTextDrv2Install", "ДРАЙВЕРА ДОСТУПНЫЕ ДЛЯ УСТАНОВКИ:")
    strTTipTextDrv4UnsupOS = LocaliseString(strPathFile, strFormName, "ToolTipTextDrv4UnsupportedOS", "ВНИМАНИЕ! ДРАЙВЕРА ДЛЯ ДРУГОЙ ОС." & vbNewLine & "ОБАБОТКА ВКЛАДКИ ВЫКЛЮЧЕНА В НАСТРОЙКАХ")
    strTTipTextTitleStatus = LocaliseString(strPathFile, strFormName, "ToolTipTextTitleStatus", "Подробное описание:")
    strSSTabTypeDPTab1 = LocaliseString(strPathFile, strFormName, "SSTabTypeDPTab1", "Все драйверпаки")
    strSSTabTypeDPTab2 = LocaliseString(strPathFile, strFormName, "SSTabTypeDPTab2", "Доступно обновление")
    strSSTabTypeDPTab3 = LocaliseString(strPathFile, strFormName, "SSTabTypeDPTab3", "Неустановленные")
    strSSTabTypeDPTab4 = LocaliseString(strPathFile, strFormName, "SSTabTypeDPTab4", "Установленные")
    strSSTabTypeDPTab5 = LocaliseString(strPathFile, strFormName, "SSTabTypeDPTab5", "БД не создана")
    ' Прописываем как константу длину названия колонок
    lngTableHwidHeader1 = Len(strTableHwidHeader1)
    lngTableHwidHeader2 = Len(strTableHwidHeader2)
    lngTableHwidHeader4 = Len(strTableHwidHeader4)
    lngTableHwidHeader5 = Len(strTableHwidHeader5)
    lngTableHwidHeader6 = Len(strTableHwidHeader6)
    lngTableHwidHeader7 = Len(strTableHwidHeader7)
    lngTableHwidHeader8 = Len(strTableHwidHeader8)
    lngTableHwidHeader9 = Len(strTableHwidHeader9)
    lngTableHwidHeader10 = Len(strTableHwidHeader10)
    lngTableHwidHeader11 = Len(strTableHwidHeader11)
    lngTableHwidHeader12 = Len(strTableHwidHeader12)
    lngTableHwidHeader13 = Len(strTableHwidHeader13)
    lngTableHwidHeader14 = Len(strTableHwidHeader14)
    ' Информация о PC/Windows
    lblOsInfoChange
    ' Перегрузка ListChecked
    LoadListChecked
    ' Перегрузка FrmMainCaption
    ChangeFrmMainCaption
    ' Перегрузка ToolTip
    ToolTipStatusLoad
    ToolTipOtherControlReLoad
    ' Изменение SSTab2
    LoadSSTab2Desc
    ' Перегрузка сообщений
    LocaliseMessage strPCLangCurrentPath

    If mbDP_Is_aFolder Then
        frRezim.Caption = frRezim.Caption & strSpace & strMessages(150)
    End If

    ' Установка текста панели
    ctlUcStatusBar1.PanelText(1) = strMessages(127)

    ' Если это не началный запуск программы, то изменяем еще и эти параметры
    If Not mbFirstStart Then
        ' изменение caption кнопки CmdViewAll
        LoadCmdViewAllDeviceCaption
        ' Перезагрузка всплывающих подсказок для кнопок с драйверами
        Me.Font.Name = strFontMainForm_Name
        Me.Font.Size = lngFontMainForm_Size
        ToolTipBtnReLoad
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LocaliseMenu
'! Description (Описание)  :   [Загрузка текста меню с поддеркой Unicode]
'! Parameters  (Переменные):   strPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub LocaliseMenu(ByVal strPathFile As String)
    ' Меню должно быть видимым, так как для невидимого не применяется изменение свойства
    ' Поэтому просто изменяем у них свойство caption, и делаем меню неактивным
    mnuContextMenu.Caption = "Drivers"
    mnuContextMenu2.Caption = "Installer"
    mnuContextMenu3.Caption = "Assistant"
    mnuContextMenu4.Caption = "v." & strProductVersion
    mnuContextMenu.Enabled = False
    mnuContextMenu2.Enabled = False
    mnuContextMenu3.Enabled = False
    mnuContextMenu4.Enabled = False
    
'0  mnuRezim - "Обновление баз данных"
' 0    mnuRezimBaseDrvUpdateALL - "Обновить базы для ВСЕХ пакетов драйверов"
' 1    mnuRezimBaseDrvUpdateNew - "Обновить базы только для НОВЫХ пакетов драйверов"
' 2    mnuSep0 - "-"
' 3    mnuRezimBaseDrvClean - "Удалить файлы баз данных отсутствующих пакетов драйверов"
' 4    mnuDelDuplicateOldDP - "Удалить устаревшие версии пакетов драйверов"
' 5    mnuSep1 - "-"
' 6    mnuLoadOtherPC - "Загрузить информацию другого ПК (Эмуляция работы)"
' 7    mnuSaveInfoPC - "Сохранить информацию об устройствах для эмуляции на другом ПК"
    SetUniMenu -1, 0, -1, mnuRezim, LocaliseString(strPathFile, strFormName, "mnuRezim", mnuRezim.Caption)
    SetUniMenu 0, 0, -1, mnuRezimBaseDrvUpdateALL, LocaliseString(strPathFile, strFormName, "mnuRezimBaseDrvUpdateALL", mnuRezimBaseDrvUpdateALL.Caption)
    SetUniMenu 0, 1, -1, mnuRezimBaseDrvUpdateNew, LocaliseString(strPathFile, strFormName, "mnuRezimBaseDrvUpdateNew", mnuRezimBaseDrvUpdateNew.Caption)
    SetUniMenu 0, 3, -1, mnuRezimBaseDrvClean, LocaliseString(strPathFile, strFormName, "mnuRezimBaseDrvClean", mnuRezimBaseDrvClean.Caption)
    SetUniMenu 0, 4, -1, mnuDelDuplicateOldDP, LocaliseString(strPathFile, strFormName, "mnuDelDuplicateOldDP", mnuDelDuplicateOldDP.Caption)
    SetUniMenu 0, 6, -1, mnuLoadOtherPC, LocaliseString(strPathFile, strFormName, "mnuLoadOtherPC", mnuLoadOtherPC.Caption)
    SetUniMenu 0, 7, -1, mnuSaveInfoPC, LocaliseString(strPathFile, strFormName, "mnuSaveInfoPC", mnuSaveInfoPC.Caption)
    
'1  mnuService - "Сервис"
' 0    mnuShowHwidsTxt - "Показать HWIDs устройств компьютера (текстовый файл)"
' 1    mnuShowHwidsXLS - "Показать HWIDs устройств компьютера (файл Excel)"
' 2    mnuSep2 - "-"
' 3    mnuShowHwidsAll - "Показать ПОЛНЫЙ СПИСОК УСТРОЙСТВ компьютера" - Shortcut{F7}
' 4    mnuSep3 - "-"
' 5    mnuUpdateStatusAll - "Обновить статус всех пакетов драйверов" - Shortcut{F6}
' 6    mnuUpdateStatusTab - "Обновить статус всех пакетов драйверов (текущая вкладка)" - Shortcut+{F6}
' 7    mnuSep4 - "-"
' 8    mnuReCollectHWID - "Обновить конфигурацию оборудования" - Shortcut{F5}
' 9   mnuReCollectHWIDTab - "Обновить конфигурацию оборудования (текущая вкладка)" - Shortcut+{F5}
' 10   mnuAutoInfoAfterDelDRV - "Автообновление конфигурации при удалении драйверов" - Checked -1  'True
' 11   mnuSep5 - "-"
' 12   mnuRunSilentMode - "Запустить тихую автоматическую установку драйверов" - Shortcut{F8}
' 13   mnuSep6 - "-"
' 14   mnuCreateRestorePoint - "Создать точку восстановления системы" - Shortcut{F9}
' 15   mnuSep7 - "-"
' 16   mnuCreateBackUp - "Создать резервную копию драйверов" - Shortcut{F12}
' 17   mnuSep8 - "-"
' 18   mnuViewDPInstLog - "Просмотреть DPinst.log"
' 19   mnuSep9 - "-"
' 20   mnuOptions - "Параметры" - Shortcut^O
    SetUniMenu -1, 1, -1, mnuService, LocaliseString(strPathFile, strFormName, "mnuService", mnuService.Caption)
    SetUniMenu 1, 0, -1, mnuShowHwidsTxt, LocaliseString(strPathFile, strFormName, "mnuShowHwidsTxt", mnuShowHwidsTxt.Caption)
    SetUniMenu 1, 1, -1, mnuShowHwidsXLS, LocaliseString(strPathFile, strFormName, "mnuShowHwidsXLS", mnuShowHwidsXLS.Caption)
    SetUniMenu 1, 3, -1, mnuShowHwidsAll, LocaliseString(strPathFile, strFormName, "mnuShowHwidsAll", mnuShowHwidsAll.Caption), , "F7"
    SetUniMenu 1, 5, -1, mnuUpdateStatusAll, LocaliseString(strPathFile, strFormName, "mnuUpdateStatusAll", mnuUpdateStatusAll.Caption), , "F6"
    SetUniMenu 1, 6, -1, mnuUpdateStatusTab, LocaliseString(strPathFile, strFormName, "mnuUpdateStatusTab", mnuUpdateStatusTab.Caption), , "Shift+F7"
    SetUniMenu 1, 8, -1, mnuReCollectHWID, LocaliseString(strPathFile, strFormName, "mnuReCollectHWID", mnuReCollectHWID.Caption), , "F5"
    SetUniMenu 1, 9, -1, mnuReCollectHWIDTab, LocaliseString(strPathFile, strFormName, "mnuReCollectHWIDTab", mnuReCollectHWIDTab.Caption), , "Shift+F5"
    SetUniMenu 1, 10, -1, mnuAutoInfoAfterDelDRV, LocaliseString(strPathFile, strFormName, "mnuAutoInfoAfterDelDRV", mnuAutoInfoAfterDelDRV.Caption)
    SetUniMenu 1, 12, -1, mnuRunSilentMode, LocaliseString(strPathFile, strFormName, "mnuRunSilentMode", mnuRunSilentMode.Caption), , "F8"
    SetUniMenu 1, 14, -1, mnuCreateRestorePoint, LocaliseString(strPathFile, strFormName, "mnuCreateRestorePoint", mnuCreateRestorePoint.Caption), , "F9"
    SetUniMenu 1, 16, -1, mnuCreateBackUp, LocaliseString(strPathFile, strFormName, "mnuCreateBackUp", mnuCreateBackUp.Caption), , "F12"
    SetUniMenu 1, 18, -1, mnuViewDPInstLog, LocaliseString(strPathFile, strFormName, "mnuViewDPInstLog", mnuViewDPInstLog.Caption)
    SetUniMenu 1, 20, -1, mnuOptions, LocaliseString(strPathFile, strFormName, "mnuOptions", mnuOptions.Caption), , "Ctrl+O"
    
'2  mnuMainUtils - "Утилиты"
' 0    mnuUtils_devmgmt - "Диспетчер устройств Windows" - Shortcut^{F1}
' 1    mnuUtils_DevManView - "DevManView" - Shortcut^{F2}
' 2    mnuUtils_DoubleDriver - "DoubleDriver" - Shortcut^{F3}
' 3    mnuUtils_SIV - "System Information Viewer" - Shortcut^{F4}
' 4    mnuUtils_UDI - "Unknown Device Identifier" - Shortcut^{F5}
' 5    mnuUtils_UnknownDevices - "Unknown Devices" - Shortcut^{F6}
' 6    mnuSep10 - "-"
' 7    mnuUtils - "" - Index   0 - Visible'False
    SetUniMenu -1, 2, -1, mnuMainUtils, LocaliseString(strPathFile, strFormName, "mnuMainUtils", mnuMainUtils.Caption)
    SetUniMenu 2, 0, -1, mnuUtils_devmgmt, LocaliseString(strPathFile, strFormName, "mnuUtils_devmgmt", mnuUtils_devmgmt.Caption), , "Ctrl+F1"
    SetUniMenu 2, 1, -1, mnuUtils_DevManView, LocaliseString(strPathFile, strFormName, "mnuUtils_DevManView", mnuUtils_DevManView.Caption), , "Ctrl+F2"
    SetUniMenu 2, 2, -1, mnuUtils_DoubleDriver, LocaliseString(strPathFile, strFormName, "mnuUtils_DoubleDriver", mnuUtils_DoubleDriver.Caption), , "Ctrl+F3"
    SetUniMenu 2, 3, -1, mnuUtils_SIV, LocaliseString(strPathFile, strFormName, "mnuUtils_SIV", mnuUtils_SIV.Caption), , "Ctrl+F4"
    SetUniMenu 2, 4, -1, mnuUtils_UDI, LocaliseString(strPathFile, strFormName, "mnuUtils_UDI", mnuUtils_UDI.Caption), , "Ctrl+F5"
    SetUniMenu 2, 5, -1, mnuUtils_UnknownDevices, LocaliseString(strPathFile, strFormName, "mnuUtils_UnknownDevices", mnuUtils_UnknownDevices.Caption), , "Ctrl+F6"
    
'3  mnuMainAbout - "Справка"
' 0    mnuLinks - "Ссылки"
' 1    mnuHistory - "История изменения"
' 2    mnuHelp - "Справка по работе" - Shortcut{F1}
' 3    mnuSep11 - "-"
' 4    mnuHomePage1 - "Домашная страница программы"
' 5    mnuHomePage - "Обсуждение программы на OsZone.net"
' 6    mnuDriverPacks - "Посетить сайт driverpacks.net"
' 7    mnuDriverPacksOnMySite - "Скачать пакеты драйверов..."
' 8    mnuSep12 - "-"
' 9    mnuCheckUpd - "Проверить обновление программы"
' 10   mnuSep13 - "-"
' 11   mnuModulesVersion - "Модули..."
' 12   mnuSep14 - "-"
' 13   mnuDonate - "Поблагодарить автора..."
' 14   mnuLicence - "Лицензионное соглашение..."
' 15   mnuAbout - "О программе..."
    SetUniMenu -1, 3, -1, mnuMainAbout, LocaliseString(strPathFile, strFormName, "mnuMainAbout", mnuMainAbout.Caption)
    SetUniMenu 3, 0, -1, mnuLinks, LocaliseString(strPathFile, strFormName, "mnuLinks", mnuLinks.Caption)
    SetUniMenu 3, 1, -1, mnuHistory, LocaliseString(strPathFile, strFormName, "mnuHistory", mnuHistory.Caption)
    SetUniMenu 3, 2, -1, mnuHelp, LocaliseString(strPathFile, strFormName, "mnuHelp", mnuHelp.Caption), , "F1"
    SetUniMenu 3, 4, -1, mnuHomePage1, LocaliseString(strPathFile, strFormName, "mnuHomePage1", mnuHomePage1.Caption)
    SetUniMenu 3, 5, -1, mnuHomePage, LocaliseString(strPathFile, strFormName, "mnuHomePage", mnuHomePage.Caption)
    SetUniMenu 3, 6, -1, mnuDriverPacks, LocaliseString(strPathFile, strFormName, "mnuDriverPacks", mnuDriverPacks.Caption)
    SetUniMenu 3, 7, -1, mnuDriverPacksOnMySite, LocaliseString(strPathFile, strFormName, "mnuDriverPacksOnMySite", mnuDriverPacksOnMySite.Caption)
    SetUniMenu 3, 9, -1, mnuCheckUpd, LocaliseString(strPathFile, strFormName, "mnuCheckUpd", mnuCheckUpd.Caption)
    SetUniMenu 3, 11, -1, mnuModulesVersion, LocaliseString(strPathFile, strFormName, "mnuModulesVersion", mnuModulesVersion.Caption)
    SetUniMenu 3, 13, -1, mnuDonate, LocaliseString(strPathFile, strFormName, "mnuDonate", mnuDonate.Caption)
    SetUniMenu 3, 14, -1, mnuLicence, LocaliseString(strPathFile, strFormName, "mnuLicence", mnuLicence.Caption)
    SetUniMenu 3, 15, -1, mnuAbout, LocaliseString(strPathFile, strFormName, "mnuAbout", mnuAbout.Caption)
    
'4  mnuMainLang - "Язык"
' 0    mnuLangStart - "Использовать выбранный язык при запуске (отмена автовыбора)"
' 1    mnuSep15 - "-"
' 2    mnuLang - "" - Index0 - Visible'False
    SetUniMenu -1, 4, -1, mnuMainLang, LocaliseString(strPathFile, strFormName, "mnuMainLang", mnuMainLang.Caption)
    SetUniMenu 4, 0, -1, mnuLangStart, LocaliseString(strPathFile, strFormName, "mnuLangStart", mnuLangStart.Caption)
    
'5  mnuContextMenu - "Контекстное меню"
' 0    mnuContextXLS - "Открыть файл базы данных в программе Excel"
' 1    mnuContextTxt - "Открыть файл базы данных в текстовом виде"
' 2    mnuContextSep1 - "-"
' 3    mnuContextToolTip - "Показать список доступных драйверов для компьютера"
' 4    mnuContextSep2 - "-"
' 5    mnuContextUpdStatus - "Обновить статус пакета драйверов"
' 6    mnuContextSep3 - "-"
' 7    mnuContextEditDPName - "Изменить отображаемое имя пакета драйверов в программе"
' 8    mnuContextSep4 - "-"
' 9    mnuContextTestDRP - "Протестировать данный пакет драйверов программой 7-zip"
' 10       mnuContextSep5 - "-"
' 11       mnuContextDeleteDRP - "Удалить пакет драйверов"
' 12       mnuContextSep6 - "-"
' 13       mnuContextDeleteDevIDs - "Удалить драйвера устройств:"
'  0          mnuContextDeleteDevIDDesc - "Список драйверов доступных для удаления" -    Enabled'False
'  1          mnuContextSep7 - "-"
'  2          mnuContextDeleteDevID - "Список устройств" -    Index0 - Visible'False
' 14       mnuContextCopyHWIDs - "Скопировать HWID в буфер обмена:"
'  0          mnuContextCopyHWIDDesc - "Список доступных HWID" -    Enabled'False
'  1          mnuContextSep8 - "-"
'  2          mnuContextCopyHWID2Clipboard - "Список устройств" -    Index0 -    Visible'False
    SetUniMenu 5, 0, -1, mnuContextXLS, LocaliseString(strPathFile, strFormName, "mnuContextXLS", mnuContextXLS.Caption)
    SetUniMenu 5, 1, -1, mnuContextTxt, LocaliseString(strPathFile, strFormName, "mnuContextTxt", mnuContextTxt.Caption)
    SetUniMenu 5, 3, -1, mnuContextToolTip, LocaliseString(strPathFile, strFormName, "mnuContextToolTip", mnuContextToolTip.Caption)
    SetUniMenu 5, 5, -1, mnuContextUpdStatus, LocaliseString(strPathFile, strFormName, "mnuContextUpdStatus", mnuContextUpdStatus.Caption)
    SetUniMenu 5, 7, -1, mnuContextEditDPName, LocaliseString(strPathFile, strFormName, "mnuContextEditDPName", mnuContextEditDPName.Caption)
    SetUniMenu 5, 9, -1, mnuContextTestDRP, LocaliseString(strPathFile, strFormName, "mnuContextTestDRP", mnuContextTestDRP.Caption)
    SetUniMenu 5, 11, -1, mnuContextDeleteDRP, LocaliseString(strPathFile, strFormName, "mnuContextDeleteDRP", mnuContextDeleteDRP.Caption)
    SetUniMenu 5, 13, -1, mnuContextDeleteDevIDs, LocaliseString(strPathFile, strFormName, "mnuContextDeleteDevIDs", mnuContextDeleteDevIDs.Caption)
    SetUniMenu 5, 13, 0, mnuContextDeleteDevIDDesc, LocaliseString(strPathFile, strFormName, "mnuContextDeleteDevIDDesc", mnuContextDeleteDevIDDesc.Caption)
    SetUniMenu 5, 14, -1, mnuContextCopyHWIDs, LocaliseString(strPathFile, strFormName, "mnuContextCopyHWIDs", mnuContextCopyHWIDs.Caption)
    SetUniMenu 5, 14, 0, mnuContextCopyHWIDDesc, LocaliseString(strPathFile, strFormName, "mnuContextCopyHWIDDesc", mnuContextCopyHWIDDesc.Caption)

'6  mnuContextMenu2 - "Контекстное меню2"
' 0    mnuContextLegendIco - "Просмотреть описание всех обозначений"
    SetUniMenu 6, 0, -1, mnuContextLegendIco, LocaliseString(strPathFile, strFormName, "mnuContextLegendIco", mnuContextLegendIco.Caption)
    
'7  mnuContextMenu3 - "Контекстное меню3"
' 0    mnuContextInstallGroupDP - "Обычная установка" - Index0
' 1    mnuContextInstallGroupDP - "-" - Index1
' 2    mnuContextInstallGroupDP - "Выборочная установка" - Index2
' 3    mnuContextInstallGroupDP - "-" - Index3
' 4    mnuContextInstallGroupDP - "Распаковать в каталог - Все подобранные драйвера" - Index4
' 5    mnuContextInstallGroupDP - "Распаковать в каталог - Выбрать драйвера..." - Index5
    SetUniMenu 7, 0, -1, mnuContextInstallGroupDP(0), LocaliseString(strPathFile, strFormName, "mnuContextInstall1", mnuContextInstallGroupDP(0).Caption)
    SetUniMenu 7, 2, -1, mnuContextInstallGroupDP(2), LocaliseString(strPathFile, strFormName, "mnuContextInstall2", mnuContextInstallGroupDP(2).Caption)
    SetUniMenu 7, 4, -1, mnuContextInstallGroupDP(4), LocaliseString(strPathFile, strFormName, "mnuContextInstall3", mnuContextInstallGroupDP(4).Caption)
    SetUniMenu 7, 5, -1, mnuContextInstallGroupDP(5), LocaliseString(strPathFile, strFormName, "mnuContextInstall4", mnuContextInstallGroupDP(5).Caption)
'8  mnuContextMenu4 - "Контекстное меню3"
' 0    mnuContextInstallSingleDP - "Обычная установка" - Index0
' 1    mnuContextInstallSingleDP - "-" - Index1
' 2    mnuContextInstallSingleDP - "Выборочная установка" - Index2
' 3    mnuContextInstallSingleDP - "-" - Index3
' 4    mnuContextInstallSingleDP - "Распаковать в каталог - Все подобранные драйвера" - Index4
' 5    mnuContextInstallSingleDP - "Распаковать в каталог - Выбрать драйвера..." - Index5    SetUniMenu 8, 0, -1, mnuContextInstall(0), LocaliseString(StrPathFile, strFormName, "mnuContextInstall1", mnuContextInstall(0).Caption)
    SetUniMenu 8, 0, -1, mnuContextInstallSingleDP(0), LocaliseString(strPathFile, strFormName, "mnuContextInstall1", mnuContextInstallSingleDP(0).Caption)
    SetUniMenu 8, 2, -1, mnuContextInstallSingleDP(2), LocaliseString(strPathFile, strFormName, "mnuContextInstall2", mnuContextInstallSingleDP(2).Caption)
    SetUniMenu 8, 4, -1, mnuContextInstallSingleDP(4), LocaliseString(strPathFile, strFormName, "mnuContextInstall3", mnuContextInstallSingleDP(4).Caption)
    SetUniMenu 8, 5, -1, mnuContextInstallSingleDP(5), LocaliseString(strPathFile, strFormName, "mnuContextInstall4", mnuContextInstallSingleDP(5).Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub NoSupportOSorNoDevBD
'! Description (Описание)  :   [Разные сообщения если нет поддерживаемых вкладок, или что-то нет так с пакетами]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub NoSupportOSorNoDevBD()

    Dim lngCnt As Long

    'Если нет поддерживаемых вкладок для вашей ОС, то
    If mbNoSupportedOS Then
        SelectStartMode 3, False
        BlockControl True
        BlockControlEx False
        ChangeStatusTextAndDebug strMessages(111)
        MsgBox strMessages(111) & vbNewLine & Replace$(optRezim_Upd.Caption, vbNewLine, strSpace), vbInformation, strProductName
        mbSilentRun = False
        mbRunWithParam = False
    End If

    ' Если есть несовместимый(ые) пакеты драйверов, то выводим сообщение
    If mbNotSupportedDevDB Then
        MsgBox strMessages(112), vbInformation & vbApplicationModal, strProductName
    End If

    ' Подсчитываем кол-во пакетов не имеющих БД, и выводим сообщение
    lngCnt = FindNoDBCount

    If lngCnt Then
        If MsgBox(lngCnt & strSpace & strMessages(103), vbYesNo + vbQuestion, strProductName) = vbYes Then
            ' Выставляем вкладку для которых нет БД
            SSTab2(SSTab1.Tab).Tab = 4
            DoEvents
            SelectStartMode 3, False
            ' собственно запуск создания БД
            SilentCheckNoDB
            ' возвращаяем обратно стартовый режим
            SSTab2(SSTab1.Tab).Tab = 0
            SelectStartMode , True
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub OpenTxtInExcel
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPathTxt (String)
'!--------------------------------------------------------------------------------
Private Sub OpenTxtInExcel(ByVal strPathTxt As String)

    Dim ExcelApp As Object
    Dim ExcelDoc As Object

    If IsAppPresent("Excel.Application\CurVer", vbNullString) = False Then
        MsgBox strMessages(19), vbCritical & vbApplicationModal, strProductName
    Else
        Set ExcelApp = CreateObject("Excel.Application")
        'Определяем видимость Excel-a по True - видимый,
        'по False - не видимый (работает только ядро)
        ExcelApp.Visible = False
        'Создаем документ
        Set ExcelDoc = ExcelApp.Workbooks.Open(FileName:=strPathTxt, ReadOnly:=True)
        'активируем его и сохраняем
        ExcelDoc.Activate

        With ExcelApp
            .Cells.Select
            .Cells.EntireColumn.AutoFit
            .Visible = True
        End With

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PutAllDrivers2Log
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub PutAllDrivers2Log()

    Dim i                      As Long
    Dim strTTipTextHeaders     As String
    Dim strTemp                As String
    Dim strLineAll             As String
    Const strTableHwidHeaderDP As String = "Drivers in DriverPack"
    
    If lngSizeRowDPMax < Len(strTableHwidHeaderDP) Then
        lngSizeRowDPMax = Len(strTableHwidHeaderDP)
    End If
    
    'Формируем шапку для подсказки
    strTTipTextHeaders = strTTipTextDrv2Install & vbNewLine & _
                        String$(maxSizeRowAllLineMax, "-") & vbNewLine & _
                        UCase$(strTableHwidHeader1 & Space$(lngSizeRow1Max - lngTableHwidHeader1 + 1) & "| " & _
                        strTableHwidHeaderDP & Space$(lngSizeRowDPMax - Len(strTableHwidHeaderDP) + 1) & "| " & _
                        strTableHwidHeader2 & Space$(lngSizeRow2Max - lngTableHwidHeader2 + 1) & "| " & _
                        strTableHwidHeader4 & Space$(lngSizeRow4Max - lngTableHwidHeader4 + 1) & "| " & _
                        strTableHwidHeader9 & Space$(lngSizeRow9Max - lngTableHwidHeader9 + 1) & "| " & _
                        strTableHwidHeader5 & Space$(lngSizeRow5Max - lngTableHwidHeader5 + 1) & "| " & _
                        strTableHwidHeader6 & Space$(lngSizeRow6Max - lngTableHwidHeader6 + 1) & "| " & _
                        strTableHwidHeader7) & vbNewLine & String$(maxSizeRowAllLineMax, "-") & vbNewLine

    If mbDebugStandart Then DebugMode "===============================List of all found a matched driver===================================" & vbNewLine & strTTipTextHeaders

    ReDim Preserve arrDriversList(12, lngArrDriversIndex - 1)

    QuickSortMDArray arrDriversList, 1, 0

    For i = 0 To UBound(arrDriversList, 2)
        'strDevID
        strTemp = arrDriversList(0, i)
        arrDriversList(0, i) = strTemp & Space$(lngSizeRow1Max - Len(strTemp) + 1) & "| "
        'strDevPath
        strTemp = arrDriversList(1, i)
        arrDriversList(1, i) = strTemp & Space$(lngSizeRow2Max - Len(strTemp) + 1) & "| "
        'strDevVer
        strTemp = arrDriversList(2, i)
        arrDriversList(2, i) = strTemp & Space$(lngSizeRow4Max - Len(strTemp) + 1) & "| "
        'strDevVerLocal
        strTemp = arrDriversList(3, i)
        arrDriversList(3, i) = strTemp & Space$(lngSizeRow5Max - Len(strTemp) + 1) & "| "
        ' strPriznakSravnenia
        strTemp = arrDriversList(6, i)
        arrDriversList(6, i) = strTemp & Space$(lngSizeRow9Max - Len(strTemp) + 1) & "| "
        'strDevStatus & strDevName
        strTemp = arrDriversList(4, i)
        arrDriversList(4, i) = strTemp & Space$(lngSizeRow6Max - Len(strTemp) + 1) & "| "
        ' Секция
        strTemp = arrDriversList(7, i)
        arrDriversList(7, i) = strTemp & Space$(lngSizeRow13Max - Len(strTemp) + 1) & "|"
        ' Имя DP
        strTemp = arrDriversList(12, i)
        arrDriversList(12, i) = strTemp & Space$(lngSizeRowDPMax - Len(strTemp) + 1) & "|"
        ' Итоговый
        strLineAll = (arrDriversList(0, i) & arrDriversList(12, i) & arrDriversList(1, i) & arrDriversList(2, i) & arrDriversList(6, i)) & (arrDriversList(3, i) & arrDriversList(4, i) & arrDriversList(5, i))
        If mbDebugStandart Then DebugMode strLineAll
    Next

    If mbDebugStandart Then DebugMode "===================================================================================================="
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ReadExitCode
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lngCode (Long)
'!--------------------------------------------------------------------------------
Private Function ReadExitCode(ByVal lngCode As Long) As String

    Dim strResultText              As String
    Dim strCode                    As String
    Dim strCodeWW                  As String
    Dim strCodeXX                  As String
    Dim strCodeYY                  As String
    Dim strCodeZZ                  As String
    Dim mbCodeNotInstall           As Boolean
    Dim strCodeNotInstallCount     As Long
    Dim mbCodeInstall              As Boolean
    Dim strCodeInstallCount        As Long
    Dim strCodeReadyToInstallCount As Long
    Dim mbReadyToInstall           As Boolean
    Dim mbCodeReboot               As Boolean

    If mbDebugStandart Then DebugMode str2VbTab & "ReadExitCode-Start" & vbNewLine & _
              str2VbTab & "ReadExitCode: lngCode=" & CStr(lngCode)
    ''0xWW If a driver package could not be installed, the 0x80 bit is set. If a computer restart is necessary, the 0x40 bit is set. Otherwise, no bits are set.
    ''0xXX The number of driver packages that could not be installed.
    ''0xYY The number of driver packages that were copied to the driver store but were not installed on a device.
    ''0xZZ The number of driver packages that were installed on a device.
    strCode = CStr(Hex$(lngCode))

    If Len(strCode) = 8 Then
        strCodeWW = Left$(strCode, 2)
        strCodeXX = Mid$(strCode, 3, 2)
        strCodeYY = Mid$(strCode, 5, 2)
        strCodeZZ = Mid$(strCode, 7, 2)

        ' Получение основного кода
        Select Case strCodeWW

            Case "80"
                mbCodeNotInstall = True

                If Mid$(strCode, 3, 6) = "000000" Then
                    mbCodeNotInstall = False
                    strResultText = "Cancel or Nothing to Install"
                End If

            Case "40"
                mbCodeReboot = True
                mbCodeInstall = True

            Case "C0"
                mbCodeReboot = True
                mbCodeNotInstall = True

            Case "00"
                mbCodeInstall = True
                mbReadyToInstall = True

            Case Else
                mbCodeReboot = False
                mbCodeNotInstall = False
        End Select

    Else

        If Len(strCode) >= 1 Then
            If Len(strCode) <= 2 Then
                If StrComp(strCode, "0") = 0 Then
                    strResultText = "Cancel or Nothing to Install"
                Else
                    mbCodeInstall = True
                    strCodeZZ = strCode
                End If
            End If

        Else

            If Len(strCode) = 4 Then
                strCodeZZ = Mid$(strCode, 3, 2)
                strCodeYY = Left$(strCode, 1)
            ElseIf Len(strCode) = 3 Then
                strCodeZZ = Mid$(strCode, 2, 2)
                strCodeYY = Left$(strCode, 1)
            ElseIf Len(strCode) = 5 Then
                strCodeXX = Left$(strCode, 1)
                strCodeYY = Mid$(strCode, 2, 2)
                strCodeZZ = Mid$(strCode, 4, 2)
            ElseIf Len(strCode) = 6 Then
                strCodeXX = Left$(strCode, 2)
                strCodeYY = Mid$(strCode, 3, 2)
                strCodeZZ = Mid$(strCode, 5, 2)
            End If
        End If
    End If

    ' подсчет кол-ва произведенных операций
    If LenB(strCodeXX) Then
        strCodeNotInstallCount = CLng("&H" & Trim$(strCodeXX))
    End If

    If LenB(strCodeYY) Then
        strCodeReadyToInstallCount = CLng("&H" & Trim$(strCodeYY))
    End If

    If LenB(strCodeZZ) Then
        strCodeInstallCount = CLng("&H" & Trim$(strCodeZZ))
    End If

    ' на основании ко-ва операций устанавливаем признаки добавления в итоговую строку
    If strCodeInstallCount > 0 Then
        mbCodeInstall = True
    End If

    If strCodeReadyToInstallCount > 0 Then
        mbReadyToInstall = True
    End If

    If strCodeNotInstallCount > 0 Then
        mbCodeNotInstall = True
    End If

    ' Формирование итоговой строки
    If mbCodeInstall Then
        strResultText = IIf(LenB(strResultText), strResultText & strSpace, vbNullString) & "Install: " & strCodeInstallCount
    End If

    If mbCodeNotInstall Then
        strResultText = IIf(LenB(strResultText), strResultText & strSpace, vbNullString) & "NotInstall: " & strCodeNotInstallCount
    End If

    If mbReadyToInstall Then
        strResultText = IIf(LenB(strResultText), strResultText & strSpace, vbNullString) & "ReadyToInstall: " & strCodeReadyToInstallCount
    End If

    If mbCodeReboot Then
        strResultText = IIf(LenB(strResultText), strResultText & strSpace, vbNullString) & "Need Reboot"
    End If

    If LenB(strResultText) Then
        ReadExitCode = "(" & strResultText & ")"
    Else
        ReadExitCode = vbNullString
    End If

    If mbDebugStandart Then DebugMode str2VbTab & "ReadExitCode: strResultText=" & strResultText & vbNewLine & _
              str2VbTab & "ReadExitCode-End"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ReadOrSaveToolTip
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPathDevDB (String)
'                              strPathDRP (String)
'                              strPackFileName (String)
'                              Index (Long)
'                              mbSaveToolTip (Boolean = False)
'                              mbReloadToolTip (Boolean = False)
'!--------------------------------------------------------------------------------
Private Sub ReadOrSaveToolTip(ByVal strPathDevDB As String, ByVal strPathDRP As String, ByVal strPackFileName As String, ByVal Index As Long, Optional ByVal mbSaveToolTip As Boolean = False, Optional ByVal mbReloadToolTip As Boolean = False)

    Dim strTTipText            As String
    Dim strTTipTextTemp        As String
    Dim strClassesName         As String
    Dim strTTipTextHeadersTemp As String
    Dim strPackFileNameFull    As String
    Dim strFinishIniPath       As String
    Dim strTTipTextOnlyDrivers As String
    Dim strTTipSizeHeader_x()  As String
    Dim maxLengthRow1          As String
    Dim maxLengthRow2          As String
    Dim maxLengthRow4          As String
    Dim maxLengthRow5          As String
    Dim maxLengthRow6          As String
    Dim maxLengthRow9          As String
    Dim TimeScriptRun          As Long
    Dim TimeScriptFinish       As Long
    Dim objTT                  As TipTool
    Dim mbObjTTNotExist        As Boolean

    If mbDebugStandart Then DebugMode str3VbTab & "ReadOrSaveToolTip - Start - Driverpack: " & strPackFileName
    TimeScriptRun = GetTickCount
    ' Небольшое прерывание для большего отклика от приложения
    DoEvents

    If LenB(strPackFileName) Then
        ' Всплывающая подсказка
        strPackFileNameFull = PathCombine(strPathDRP, strPackFileName)
        
        'Считываем класс пакета из файла
        If mbReadClasses Then
            strFinishIniPath = BackslashAdd2Path(strPathDevDB) & GetFileName_woExt(strPackFileName) & ".ini"
            strClassesName = IniStringPrivate("DriverPack", "classes", strFinishIniPath)

            ' Если такого значения в файле нет, то ничего не добавляем
            If StrComp(strClassesName, "no_key") = 0 Then
                strClassesName = vbNullString
            End If
            
            If LenB(strClassesName) Then
                If Not mbDP_Is_aFolder Then
                    strTTipTextHeadersTemp = (strPathDRP & str2vbNewLine & strPackFileName & vbNewLine) & (strTTipTextFileSize & strSpace & FileSizeApi(strPackFileNameFull) & vbNewLine & strTTipTextClassDRV & strSpace & strClassesName)
                Else
                    ' Пока уберем подсчет размера директории, так как очень медленно
                    'strTTipTextHeadersTemp = strPathDRP & str2vbNewLine & strPackFileName & vbNewLine & strTTipTextFileSize & strSpace & FolderSizeApi(strPackFileNameFull, True) & vbNewLine & strTTipTextClassDRV & strSpace & strClassesName
                    strTTipTextHeadersTemp = strPathDRP & str2vbNewLine & strPackFileName & vbNewLine & strTTipTextClassDRV & strSpace & strClassesName
                End If
            End If
        Else

            If Not mbDP_Is_aFolder Then
                strTTipTextHeadersTemp = (strPathDRP & str2vbNewLine & strPackFileName) & (vbNewLine & strTTipTextFileSize & strSpace & FileSizeApi(strPackFileNameFull))
            Else
                ' Пока уберем подсчет размера директории, так как очень медленно
                'strTTipTextHeadersTemp = strPathDRP & str2vbNewLine & strPackFileName & vbNewLine & strTTipTextFileSize & strSpace & FolderSizeApi(strPackFileNameFull, True)
                strTTipTextHeadersTemp = strPathDRP & str2vbNewLine & strPackFileName
            End If
        End If

        strTTipText = strTTipTextHeadersTemp

        If Not mbReloadToolTip Then
            ' Присваиваем картинку в соответсвии с наличием БД к файлу, а также получаем текст всплывающей подсказки
            strTTipTextTemp = ChangeStatusAndPictureButton(strPathDevDB, strPackFileName, Index)
        Else

            If Not mbFirstStart Then
                strTTipTextTemp = arrTTip(Index)
            End If
        End If

        strTTipTextOnlyDrivers = strTTipTextTemp

        If LenB(strTTipTextTemp) Then
            If StrComp(strTTipTextTemp, "unsupported") <> 0 Then
                If InStr(strTTipTextTemp, "|") Then

                    'Формируем шапку для подсказки
                    If mbReloadToolTip Then
                        strTTipSizeHeader_x = Split(arrTTipSize(Index), ";")
                        maxLengthRow1 = lngTableHwidHeader1
                        maxLengthRow2 = lngTableHwidHeader2
                        maxLengthRow4 = lngTableHwidHeader4
                        maxLengthRow9 = lngTableHwidHeader9
                        maxLengthRow5 = lngTableHwidHeader5
                        maxLengthRow6 = lngTableHwidHeader6
                        maxSizeRowAllLine = strTTipSizeHeader_x(0)
                        lngSizeRow1 = strTTipSizeHeader_x(1)
                        lngSizeRow2 = strTTipSizeHeader_x(2)
                        lngSizeRow4 = strTTipSizeHeader_x(3)
                        lngSizeRow9 = strTTipSizeHeader_x(4)
                        lngSizeRow5 = strTTipSizeHeader_x(5)
                        lngSizeRow6 = strTTipSizeHeader_x(6)
    
                        If lngSizeRow1 < maxLengthRow1 Then
                            lngSizeRow1 = maxLengthRow1
                        End If
    
                        If lngSizeRow2 < maxLengthRow2 Then
                            lngSizeRow2 = maxLengthRow2
                        End If
    
                        If lngSizeRow4 < maxLengthRow4 Then
                            lngSizeRow4 = maxLengthRow4
                        End If
    
                        If lngSizeRow5 < maxLengthRow5 Then
                            lngSizeRow5 = maxLengthRow5
                        End If
    
                        If lngSizeRow6 < maxLengthRow6 Then
                            lngSizeRow6 = maxLengthRow6
                        End If
    
                        If lngSizeRow9 < maxLengthRow9 Then
                            lngSizeRow9 = maxLengthRow9
                        End If
                    End If
    
                    strTTipTextHeaders = strTTipTextHeadersTemp & str2vbNewLine & _
                                         strTTipTextDrv2Install & vbNewLine & _
                                         String$(maxSizeRowAllLine, "-") & vbNewLine & _
                                         UCase$(strTableHwidHeader1 & Space$(lngSizeRow1 - lngTableHwidHeader1 + 1) & "| " & _
                                         strTableHwidHeader2 & Space$(lngSizeRow2 - lngTableHwidHeader2 + 1) & "| " & _
                                         strTableHwidHeader4 & Space$(lngSizeRow4 - lngTableHwidHeader4 + 1) & "| " & _
                                         strTableHwidHeader9 & Space$(lngSizeRow9 - lngTableHwidHeader9 + 1) & "| " & _
                                         strTableHwidHeader5 & Space$(lngSizeRow5 - lngTableHwidHeader5 + 1) & "| " & _
                                         strTableHwidHeader6 & Space$(lngSizeRow6 - lngTableHwidHeader6 + 1) & "| " & _
                                         strTableHwidHeader7) & vbNewLine & String$(maxSizeRowAllLine, "-") & vbNewLine
                    'Текст итоговой подсказки
                    strTTipText = strTTipTextHeaders & strTTipTextTemp & vbNewLine & String$(maxSizeRowAllLine, "-")
                Else
                    strTTipText = strTTipTextHeadersTemp & str2vbNewLine & strTTipTextDrv4UnsupOS
                    strTTipTextOnlyDrivers = strTTipTextDrv4UnsupOS
                End If
            Else
                strTTipText = strTTipTextHeadersTemp & str2vbNewLine & strTTipTextDrv4UnsupOS
                strTTipTextOnlyDrivers = strTTipTextDrv4UnsupOS
            End If
        End If

        ' Сохраняем подсказку в массив, если требуется
        If mbSaveToolTip Then
            If mbFirstStart Then

                ReDim Preserve arrTTip(Index)

                arrTTip(Index) = strTTipTextOnlyDrivers
            Else
                arrTTip(Index) = strTTipText
                If mbDebugDetail Then DebugMode str4VbTab & "ReadOrSaveToolTip: ToolTipArrIndex=" & Index & ":" & UBound(arrTTip)
                If mbDebugStandart Then DebugMode (str4VbTab & "ReadOrSaveToolTip: strTTipText=" & vbNewLine & "=========================================================================================" & vbNewLine) & strTTipText
            End If
        End If
        
        ' Создание/изменение подсказки
        ' В цикле проходит все подсказки и проверяем, создавалась ли такая подсказка ранее для данного контрола
        For Each objTT In TT.Tools
            If StrComp(objTT.hWnd, acmdPackFiles(Index).hWnd) = 0 Then
                ' Если создавалась, то меняем текст
                mbObjTTNotExist = True
                objTT.Text = strTTipText
                Exit For
            End If
        Next
        ' Если подсказка не найдена, то создаем новую
        If Not mbObjTTNotExist Then
            TT.Tools.Add acmdPackFiles(Index).hWnd, , strTTipText, True
        End If
        
        TimeScriptFinish = GetTickCount
        If mbDebugStandart Then DebugMode str3VbTab & "ReadOrSaveToolTip - End - Time to Read Driverpack's - " & strPackFileName & ": " & CalculateTime(TimeScriptRun, TimeScriptFinish, True)
    Else
        If mbDebugDetail Then DebugMode str4VbTab & "ReadOrSaveToolTip: Empty Driverpack's Name"
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ReOrderBtnOnTab2
'! Description (Описание)  :   [Запуск перестроение кнопок на определенной вкладке]
'! Parameters  (Переменные):   lngTab2Tab (Long)
'                              lngBtnPrevCnt (Long)
'                              lngBtnTabCnt (Long)
'                              objScrollControl (Object)
'!--------------------------------------------------------------------------------
Private Sub ReOrderBtnOnTab2(ByVal lngTab2Tab As Long, ByVal lngBtnPrevCnt As Long, ByVal lngBtnTabCnt As Long, objScrollControl As Object)

    Dim i               As Long
    Dim lngStartPosLeft As Long
    Dim lngStartPosTop  As Long
    Dim lngNextPosLeft  As Long
    Dim lngNextPosTop   As Long
    Dim lngMaxLeftPos   As Long
    Dim lngDeltaPosLeft As Long
    Dim lngDeltaPosTop  As Long
    Dim lngBtnPrevNum   As Long
    Dim lngNoDP4ModeCnt As Long

    lngStartPosLeft = lngButtonLeft
    lngStartPosTop = lngButtonTop
    lngBtnPrevNum = 0
    lngNoDP4ModeCnt = 0
    objScrollControl.Visible = False

    If lngTab2Tab = 0 Then
        If objScrollControl.ScrollPositionH Then
            objScrollControl.ScrollPositionH = 0
        End If
    End If
    
    For i = lngBtnPrevCnt To lngBtnTabCnt

        With acmdPackFiles(i)
            If Not (.PictureNormal Is Nothing) Then
    
                Select Case lngTab2Tab
    
                    Case 0
                    
                        GoTo MoveBtn
                        
                    Case 1
    
                        If .PictureNormal = imgOkNew.Picture Then
                            GoTo MoveBtn
                        ElseIf .PictureNormal = imgOkAttentionNew.Picture Then
                            GoTo MoveBtn
                        Else
                            .TabStop = False
                            GoTo NextBtn
                        End If
    
                    Case 2
    
                        If .PictureNormal = imgOkAttention.Picture Then
                            GoTo MoveBtn
                        ElseIf .PictureNormal = imgOkAttentionOLD.Picture Then
                            GoTo MoveBtn
                        ElseIf .PictureNormal = imgOkAttentionNew.Picture Then
                            GoTo MoveBtn
                        Else
                            .TabStop = False
                            GoTo NextBtn
                        End If
    
                    Case 3
    
                        If .PictureNormal = imgOK.Picture Then
                            GoTo MoveBtn
                        ElseIf .PictureNormal = imgOkAttentionOLD.Picture Then
                            GoTo MoveBtn
                        ElseIf .PictureNormal = imgOkAttentionNew.Picture Then
                            GoTo MoveBtn
                        ElseIf .PictureNormal = imgOkNew.Picture Then
                            GoTo MoveBtn
                        ElseIf .PictureNormal = imgOkOld.Picture Then
                            GoTo MoveBtn
                        Else
                            .TabStop = False
                            GoTo NextBtn
                        End If
    
                    Case 4
    
                        If .PictureNormal = imgNoDB.Picture Then
                            GoTo MoveBtn
                        Else
                            .TabStop = False
                            GoTo NextBtn
                        End If
                End Select
    
MoveBtn:
                ' Собственно перемещаем кнопки на другую вкладку
                Set chkPackFiles(i).Container = objScrollControl
                Set .Container = objScrollControl
    
                ' положения кнопок
                If i = 0 Then
                    lngNextPosLeft = lngStartPosLeft
                    lngNextPosTop = lngStartPosTop
                Else
                    
                    If lngBtnPrevNum Then
                        lngDeltaPosLeft = acmdPackFiles(lngBtnPrevNum).Left + lngButtonWidth + lngBtn2BtnLeft - lngStartPosLeft
                    Else
    
                        ' Если первая кнопка подходит, то расчитываем следующее положение исходя из нее
                        If lngTab2Tab Then
                            If IsChildOfControl(acmdPackFiles(0).hWnd, objScrollControl.hWnd) Then
                                lngDeltaPosLeft = acmdPackFiles(0).Left + lngButtonWidth + lngBtn2BtnLeft - lngStartPosLeft
                            End If
    
                        Else
                            If i = lngBtnPrevCnt Then
                                If IsChildOfControl(acmdPackFiles(0).hWnd, objScrollControl.hWnd) = False Then
                                    lngNextPosLeft = lngStartPosLeft
                                    lngNextPosTop = lngStartPosTop
                                Else
                                    lngDeltaPosLeft = acmdPackFiles(0).Left + lngButtonWidth + lngBtn2BtnLeft - lngStartPosLeft
                                End If
                            Else
                                lngDeltaPosLeft = acmdPackFiles(0).Left + lngButtonWidth + lngBtn2BtnLeft - lngStartPosLeft
                            End If
                        End If
                    End If
    
                    lngNextPosLeft = lngStartPosLeft + lngDeltaPosLeft
                    lngMaxLeftPos = lngNextPosLeft + lngButtonWidth + 25
    
                    If lngMaxLeftPos > objScrollControl.Width Then
                        ' Если по горизонтали кнопка не входит, то перешагиваем
                        lngDeltaPosLeft = 0
                        lngDeltaPosTop = lngDeltaPosTop + lngButtonHeight + lngBtn2BtnTop
                        lngNextPosLeft = lngStartPosLeft
                        lngNextPosTop = lngStartPosTop + lngDeltaPosTop
                    Else
                        lngNextPosTop = lngStartPosTop + lngDeltaPosTop
                    End If
                End If
    
                ' Перемещение кнопок и checkbox по расчитанным ранее параметрам
                .Move lngNextPosLeft, lngNextPosTop
                .TabStop = True
                chkPackFiles(i).Move (lngNextPosLeft + 50), (lngNextPosTop + (lngButtonHeight - chkPackFiles(i).Height) / 2)
                chkPackFiles(i).ZOrder 0
                ' Увеличиваем счетчики
                lngBtnPrevNum = i
                lngNoDP4ModeCnt = lngNoDP4ModeCnt + 1
NextBtn:
                ' Clear value
                If .Value Then
                    .Value = False
                End If
            End If
        End With
    Next i

    If lngNoDP4ModeCnt = 0 Then

        With lblNoDP4Mode

            On Error Resume Next

            Set .Container = objScrollControl
            .Left = 100
            .Width = objScrollControl.Width - 200
            .Top = (objScrollControl.Height - .Height) / 2
            .Visible = True
            .ZOrder 0
        End With

    End If

    objScrollControl.Visible = True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function RunDPInst
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strWorkPath (String)
'!--------------------------------------------------------------------------------
Private Function RunDPInst(ByVal strWorkPath As String) As Long

    Dim cmdString As String

    If mbDebugStandart Then DebugMode "RunDPInst-Start" & vbNewLine & _
              "RunDPInst: strWorkPath" & strWorkPath

    cmdString = strKavichki & strDPInstExePath & strKavichki & strSpace & CollectCmdString & "/PATH " & strKavichki & strWorkPath & strKavichki
    ChangeStatusTextAndDebug strMessages(93)

    If RunAndWaitNew(cmdString, GetPathNameFromPath(strDPInstExePath), vbNormalFocus) = False Then
        If Not mbSilentRun Then
            MsgBox strMessages(21) & str2vbNewLine & cmdString, vbInformation, strProductName
        End If

        ChangeStatusTextAndDebug strMessages(21) & strSpace & cmdString
        If mbDebugStandart Then DebugMode "Error on run : " & cmdString
    Else
        RunDPInst = lngExitProc

        If RunDPInst <> 0 Then
            If RunDPInst <> -2147483648# Then
                ' Сбор сведений о PC
                ChangeStatusTextAndDebug strMessages(94)
                RunDevcon
                DevParserLocalHwids2
                ChangeStatusTextAndDebug strMessages(95)
                ' Сбор данных из реестра
                CollectHwidFromReestr
                ChangeStatusTextAndDebug strMessages(96) & strSpace & cmdString
            End If
        End If
    End If

    If mbDebugStandart Then DebugMode "RunDPInst-End"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SelectAllOnTabDP
'! Description (Описание)  :   [Выделение все пакеты драйверов на текущей вкладке]
'! Parameters  (Переменные):   mbIntellectMode (Boolean = True)
'!--------------------------------------------------------------------------------
Private Sub SelectAllOnTabDP(Optional ByVal mbIntellectMode As Boolean = True)

    If SSTab1.Enabled Then
        'MsgBox "Выбираем нужный режим установки"
        If mbIntellectMode Then
            SelectStartMode 1, False
        Else
            SelectStartMode 2, False
        End If

        cmbCheckButton.ListIndex = 3
        cmbCheckButton.Refresh
        DoEvents
        cmdCheck_Click
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SelectNextTab
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SelectNextTab()

    Dim lng2Tab As Long

    If frTabPanel.Visible Then
        'If SSTab1.Visible Then
        lng2Tab = SSTab1.Tab + 1

        Do While lng2Tab <= SSTab1.Tabs - 1

            If lng2Tab = SSTab1.Tabs - 1 Then
                lng2Tab = SetFirstEnableTab
            End If

            If SSTab1.TabEnabled(lng2Tab) Then
                If SSTab1.TabVisible(lng2Tab) Then
                    SSTab1.Tab = lng2Tab
                    SSTab1.SetFocus

                    Exit Do

                End If
            End If

            lng2Tab = lng2Tab + 1
        Loop

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SelectNotInstalledDP
'! Description (Описание)  :   [Выделение пакетов c неустановленными драйверами]
'! Parameters  (Переменные):   mbIntellectMode (Boolean = True)
'!--------------------------------------------------------------------------------
Private Sub SelectNotInstalledDP(Optional ByVal mbIntellectMode As Boolean = True)

    If SSTab1.Enabled Then

        If mbIntellectMode Then
            SelectStartMode 1, False
        Else
            SelectStartMode 2, False
        End If

        cmbCheckButton.ListIndex = 1
        cmbCheckButton.Refresh
        DoEvents
        cmdCheck_Click
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SelectRecommendedDP
'! Description (Описание)  :   [Выделение пакетов рекомендованных к установке]
'! Parameters  (Переменные):   mbIntellectMode (Boolean = True)
'!--------------------------------------------------------------------------------
Private Sub SelectRecommendedDP(Optional ByVal mbIntellectMode As Boolean = True)

    If SSTab1.Enabled Then
        'MsgBox "Выбираем нужный режим установки"

        If mbIntellectMode Then
            SelectStartMode 1, False
        Else
            SelectStartMode 2, False
        End If

        cmbCheckButton.ListIndex = 2
        cmbCheckButton.Refresh
        DoEvents
        cmdCheck_Click
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SelectStartMode
'! Description (Описание)  :   [Выбор стартового режима работы программы]
'! Parameters  (Переменные):   miModeTemp (Long = 0)
'                              mbTab2 (Boolean = True)
'!--------------------------------------------------------------------------------
Private Sub SelectStartMode(Optional miModeTemp As Long = 0, Optional mbTab2 As Boolean = True)

    Dim i_i    As Long
    Dim miMode As Long

    ' Если указан параметр miModeTemp значит это переклбчени вкладок не при старте программы
    If miModeTemp Then
        miMode = miModeTemp
    Else
        miMode = miStartMode
    End If

    If mbDebugStandart Then DebugMode "Start Rezim: " & miMode

    ' Режим при старте
    Select Case miMode

        Case 1

            If optRezim_Intellect.Enabled Then
                'optRezim_Upd.Value = False
                'optRezim_Intellect.Value = False
                optRezim_Intellect.Value = True
                optRezim_Intellect_Click
            Else
                'optRezim_Ust.Value = False
                'optRezim_Intellect.Value = False
                optRezim_Upd.Value = True
                optRezim_Upd_Click
            End If

        Case 2

            If optRezim_Ust.Enabled Then
                'optRezim_Upd.Value = False
                'optRezim_Intellect.Value = False
                optRezim_Ust.Value = True
                optRezim_Ust_Click
            Else
                'optRezim_Ust.Value = False
                'optRezim_Intellect.Value = False
                optRezim_Upd.Value = True
                optRezim_Upd_Click
            End If

        Case 3
            'optRezim_Ust.Value = False
            'optRezim_Intellect.Value = False
            optRezim_Upd.Value = True
            optRezim_Upd_Click
    End Select

    ' выставляем вторую вкладку только при старте программы
    If mbTab2 Then
        If miMode <> 3 Then
            If lngStartModeTab2 Then

                For i_i = SSTab2.LBound To SSTab2.UBound

                    ' Если вкладка активна, то выставляем начальную
                    If SSTab2(i_i).TabEnabled(lngStartModeTab2) = True Then
                        If SSTab2(i_i).Tab <> lngStartModeTab2 Then
                            SSTab2(i_i).Tab = lngStartModeTab2
                        End If
                    Else
                        SSTab2(i_i).Tab = 0
                    End If

                Next

            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function SetFirstEnableTab
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function SetFirstEnableTab() As Long

    Dim i As Long

    For i = 0 To SSTab1.Tabs - 1

        If SSTab1.TabVisible(i) Then
            If SSTab1.TabEnabled(i) Then
                SetFirstEnableTab = i

                Exit For

            End If
        End If

    Next

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetScrollFramePos
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sgnNum (Single)
'                              LngValue (Long)
'                              lngCntTab (Long)
'!--------------------------------------------------------------------------------
Private Sub SetScrollFramePos(ByVal sgnNum As Single, ByVal LngValue As Long, ByVal lngCntTab As Long)

    Dim i                As Integer
    Dim SSTabHeight      As Long
    Dim SSTabTabHeight   As Long
    Dim miValue3         As Long
    Dim lngControlHeight As Long
    Dim lngControlWidth  As Long

    SSTabTabHeight = SSTab1.TabHeight
    SSTabHeight = SSTab1.Height
    miValue3 = frRunChecked.Left + frRunChecked.Width - 50

    For i = SSTab2.LBound To SSTab2.UBound

        With SSTab2(i)

            If Not (SSTab2.item(i) Is Nothing) Then

                If lngCntTab > lngOSCountPerRow Then
                    If sgnNum = LngValue Then
                        .Top = sgnNum * SSTabTabHeight + 35
                        .Height = SSTabHeight - 60 - sgnNum * SSTabTabHeight
                        .Width = miValue3 - 100 * (sgnNum + 1)
                    Else
                        .Top = (LngValue + 1) * SSTabTabHeight + 35
                        .Height = SSTabHeight - 60 - (LngValue + 1) * SSTabTabHeight
                        .Width = miValue3 - 100 * (sgnNum + 1)
                    End If

                Else
                    .Top = SSTabTabHeight + 35
                    .Height = SSTabHeight - 60 - SSTabTabHeight
                    .Width = miValue3 - 55
                End If

                .Visible = SSTab1.TabEnabled(i)

                If .Visible Then
                    lngControlHeight = .Height - .TabHeight - 120
                    lngControlWidth = .Width - 100
                    ctlScrollControl1(i).Height = lngControlHeight
                    ctlScrollControl1(i).Width = lngControlWidth
                    ctlScrollControlTab1(i).Height = lngControlHeight
                    ctlScrollControlTab1(i).Width = lngControlWidth
                    ctlScrollControlTab2(i).Height = lngControlHeight
                    ctlScrollControlTab2(i).Width = lngControlWidth
                    ctlScrollControlTab3(i).Height = lngControlHeight
                    ctlScrollControlTab3(i).Width = lngControlWidth
                    ctlScrollControlTab4(i).Height = lngControlHeight
                    ctlScrollControlTab4(i).Width = lngControlWidth
                End If
            End If

        End With

    Next

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetStartScrollFramePos
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   miUnHideTabTemp (Integer)
'!--------------------------------------------------------------------------------
Private Sub SetStartScrollFramePos(ByVal miUnHideTabTemp As Integer)

    Dim cntUnHideTab As Long
    Dim miValue1     As Long
    Dim miValue2     As Long
    Dim sngNum1      As Single
    Dim sngNum2      As Single

    If mbTabHide Then
        cntUnHideTab = miUnHideTabTemp + 1
        sngNum1 = cntUnHideTab / lngOSCountPerRow
        miValue1 = Round(sngNum1, 0)

        If cntUnHideTab Then
            SetScrollFramePos sngNum1, miValue1, cntUnHideTab
        End If

    Else
        sngNum2 = lngOSCount / lngOSCountPerRow
        miValue2 = Round(sngNum2, 0)
        SetScrollFramePos sngNum2, miValue2, lngOSCount
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetTabProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SetTabProperties()

    With SSTab1
        .Font.Name = strFontTab_Name
        .Font.Size = miFontTab_Size
        .Font.Underline = mbFontTab_Underline
        .Font.Strikethrough = mbFontTab_Strikethru
        .Font.Bold = mbFontTab_Bold
        .Font.Italic = mbFontTab_Italic
        .ForeColor = lngFontTab_Color
        .Font.Charset = lngFont_Charset
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetTabPropertiesTabDrivers
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SetTabPropertiesTabDrivers()

    'Сохранение визуально заданых свойств шрифтов в переменных
    If mbFirstStart Then

        With SSTab2(0)
            .Font.Name = strFontTab2_Name
            .Font.Size = miFontTab2_Size
            .Font.Underline = mbFontTab2_Underline
            .Font.Strikethrough = mbFontTab2_Strikethru
            .Font.Bold = mbFontTab2_Bold
            .Font.Italic = mbFontTab2_Italic
            .ForeColor = lngFontTab2_Color
            .Font.Charset = lngFont_Charset
        End With

    Else

        Dim i As Long

        With SSTab2

            For i = .LBound To .UBound

                With .item(i)
                    .Font.Name = strFontTab2_Name
                    .Font.Size = miFontTab2_Size
                    .Font.Underline = mbFontTab2_Underline
                    .Font.Strikethrough = mbFontTab2_Strikethru
                    .Font.Bold = mbFontTab2_Bold
                    .Font.Italic = mbFontTab2_Italic
                    .ForeColor = lngFontTab2_Color
                    .Font.Charset = lngFont_Charset
                End With

            Next

        End With

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetTabsNameAndCurrTab
'! Description (Описание)  :   [Назначить имена для вкладок и установить текущую на основании версии ОС]
'! Parameters  (Переменные):   mbSecondStart (Boolean)
'!--------------------------------------------------------------------------------
Private Sub SetTabsNameAndCurrTab(ByVal mbSecondStart As Boolean)

    Dim i               As Long
    Dim i_i             As Long
    Dim miTabIndex      As Long
    Dim miFirstTabIndex As Long
    Dim miSymbol        As Long
    Dim strTabIndex     As String
    Dim strTabIndex_x() As String
    Dim strTabIndexTemp As String
    Dim strTabName      As String
    Dim lng_x64         As Long
    Dim lngSupportedOS  As Long

    For i = 0 To UBound(arrOSList)
        strTabName = arrOSList(i).Name
        lng_x64 = CLng(arrOSList(i).is64bit)

        If InStr(arrOSList(i).Ver, strOSCurrentVersion) Then

            ' Если в списке есть ОС x64
            If lng_x64 = 1 Then
                If InStr(strTabName, "64") = 0 Then
                    strTabName = strTabName & " x64"
                End If
            End If

            If lng_x64 = 2 Then
                miTabIndex = i
                strTabIndex = IIf(LenB(strTabIndex), strTabIndex & strSpace, vbNullString) & CStr(miTabIndex)
                lngSupportedOS = lngSupportedOS + 1
            ElseIf lng_x64 = 3 Then
                miTabIndex = i
                strTabIndex = IIf(LenB(strTabIndex), strTabIndex & strSpace, vbNullString) & CStr(miTabIndex)
                lngSupportedOS = lngSupportedOS + 1
            Else

                If CBool(lng_x64) = mbIsWin64 Then
                    miTabIndex = i
                    strTabIndex = IIf(LenB(strTabIndex), strTabIndex & strSpace, vbNullString) & CStr(miTabIndex)
                    lngSupportedOS = lngSupportedOS + 1
                End If
            End If
        End If

        SSTab1.TabCaption(i) = strTabName
    Next

    'Если среди вкладок не найдено поддержки вашей ОС
    mbNoSupportedOS = lngSupportedOS = 0

    miSymbol = InStr(strTabIndex, strSpace)

    If miSymbol Then
        strTabIndexTemp = Trim$(Left$(strTabIndex, miSymbol))
        miFirstTabIndex = CInt(strTabIndexTemp)
    Else
        miFirstTabIndex = miTabIndex
    End If

    If mbSecondStart Then
        strTabIndex_x = Split(strTabIndex, strSpace)

        For i_i = 0 To UBound(strTabIndex_x)

            If arrOSList(strTabIndex_x(i_i)).CntBtn = 0 Then
                miFirstTabIndex = 9999
            Else
                miFirstTabIndex = strTabIndex_x(i_i)

                Exit For

            End If

        Next

    End If

    If mbSecondStart Then
        If miFirstTabIndex <> 9999 Then
            SSTab1.Tab = miFirstTabIndex
            lngSSTabCurrentOS = miFirstTabIndex
        Else
            SetFirstEnableTab
            mbNoSupportedOS = True
            '"Программе не удалось найти комплект драйверов предназначенной для вашей ОС. Возможно неверно указаны настройки программы, или каталогов, указанных в настройках, с драйверами не существует."
        End If

    Else
        SSTab1.Tab = miFirstTabIndex
        lngSSTabCurrentOS = miFirstTabIndex
    End If

    strSSTabCurrentOSList = strTabIndex
    lngFirstActiveTabIndex = SetFirstEnableTab
    If mbDebugStandart Then DebugMode vbTab & "SetTabsNameAndCurrTab: SetCurrentTabOSList=" & strTabIndex & vbNewLine & _
              vbTab & "SetTabsNameAndCurrTab: SetCurrentTab=" & miFirstTabIndex
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetTabsWidth
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   miUnHideTabTemp (Integer)
'!--------------------------------------------------------------------------------
Private Sub SetTabsWidth(ByVal miUnHideTabTemp As Integer)

    Dim cntUnHideTab As Integer
    Dim miValue      As Integer

    If mbTabHide Then
        cntUnHideTab = miUnHideTabTemp + 1
        miValue = frRunChecked.Left + frRunChecked.Width - 50

        With SSTab1

            If cntUnHideTab Then
                If cntUnHideTab < lngOSCountPerRow Then
                    If cntUnHideTab > 1 Then
                        .TabMaxWidth = Round(miValue / cntUnHideTab) - 200
                    Else
                        .TabMaxWidth = Round(miValue / cntUnHideTab) - 800
                    End If

                Else
                    .Height = Me.Height - .Top - 1250
                    .TabMaxWidth = 0
                    .Width = miValue
                    .TabMaxWidth = 0
                End If
            End If

        End With

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ShowMsbBoxForm
'! Description (Описание)  :   [Вызов функции для показана формы с сообщением вместо стандартного MsgBox]
'! Parameters  (Переменные):   strMsgDialog (String)
'                              strMsgFrmCaption (String)
'                              strMsgOKCaption (String)
'!--------------------------------------------------------------------------------
Private Function ShowMsbBoxForm(strMsgDialog As String, strMsgFrmCaption As String, strMsgOKCaption As String) As Long
    lngShowMessageResult = 0
    Load frmShowMessage

    With frmShowMessage
        .txtMessageText.Text = strMsgDialog
        .Caption = strMsgFrmCaption
        .cmdOK.Caption = strMsgOKCaption
        .Show vbModal, Me
    End With

    ShowMsbBoxForm = lngShowMessageResult
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SilentCheckNoDB
'! Description (Описание)  :   [Сценарий запуска тихой установки]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SilentCheckNoDB()
    DoEvents
    SelectStartMode 3, False
    'Выбираем всё рекомендованное для установки
    cmbCheckButton.ListIndex = 3
    cmbCheckButton.Refresh
    DoEvents
    cmdCheck_Click
    'Собственно запускаем сам процесс создания БД
    mbGroupTask = True
    mbSelectInstall = False
    DoEvents
    cmdRunTask_Click
    FindNoDBCount
    mbGroupTask = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SilentInstall
'! Description (Описание)  :   [Сценарий запуска тихой установки]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SilentInstall()
    'Команда для программы DPInst работать в тихом режиме
    mbDpInstQuietInstall = True
    ' для работы в тихом режиме надо обязательно отключать promt
    mbDpInstPromptIfDriverIsNotBetter = False
    ' Включение отладочного режима
    mbDebugStandart = True
    If mbDebugStandart Then DebugMode "SilentInstall-Start" & vbNewLine & _
              vbTab & "SilentInstall: SelectMode: " & strSilentSelectMode
    
    'MsgBox "Выбираем нужный режим установки"
    Select Case strSilentSelectMode

        Case "n"
            ' новые
            SelectRecommendedDP True

        Case "q"
            ' неустановленные
            SelectNotInstalledDP True

        Case "a"
            ' Все на вкладке
            SelectAllOnTabDP True

        Case "n2"
            ' новые
            SelectRecommendedDP False

        Case "q2"
            ' неустановленные
            SelectNotInstalledDP False

        Case "a2"
            ' Все на вкладке
            SelectAllOnTabDP False

        Case Else
            ' по умолчанию (новые)
            SelectRecommendedDP True
    End Select

    'MsgBox "Собственно запускаем сам процесс установки"
    mbGroupTask = True
    mbSelectInstall = False
    DoEvents
    GroupInstallDP
    mbGroupTask = False
    If mbDebugStandart Then DebugMode "SilentInstall-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SilentReindexAllDB
'! Description (Описание)  :   [Сценарий запуска полной реиндексации]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SilentReindexAllDB()
    'Устанавливаем режим обновления
    DoEvents
    SelectStartMode 3, False
    
    'Выбираем все пакеты драйверов
    CheckAllButton True
    DoEvents
    
    'Собственно запускаем сам процесс создания БД
    mbGroupTask = True
    mbSelectInstall = False
    cmdRunTask_Click
    DoEvents
    FindNoDBCount
    mbGroupTask = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub StartReOrderBtnOnTab2
'! Description (Описание)  :   [Запуск перестроение кнопок на активной вкладке]
'! Parameters  (Переменные):   miIndex (Integer)
'                              miPrevTab (Integer)
'!--------------------------------------------------------------------------------
Private Sub StartReOrderBtnOnTab2(ByVal miIndex As Integer, ByVal miPrevTab As Integer)

    Dim lngCntBtnTab      As Long
    Dim lngCntBtnPrevious As Long

    If Not mbFirstStart Then
        lblNoDP4Mode.Visible = False
        lngCntBtnTab = arrOSList(miIndex).CntBtn - 1

        If lngCntBtnTab >= 0 Then
            If miIndex Then
                lngCntBtnPrevious = arrOSList(miIndex - 1).CntBtn

                If lngCntBtnPrevious = 0 Then
                    If miIndex > 1 Then
                        lngCntBtnPrevious = arrOSList(miIndex - 2).CntBtn

                        If lngCntBtnPrevious = 0 Then
                            If miIndex > 2 Then
                                lngCntBtnPrevious = arrOSList(miIndex - 2).CntBtn

                                If lngCntBtnPrevious = 0 Then
                                    If miIndex > 3 Then
                                        lngCntBtnPrevious = arrOSList(miIndex - 3).CntBtn
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            DoEvents

            Select Case SSTab2(miIndex).Tab

                    ' Построение пакетов со всеми драйверами (возврат всех кнопок на место)
                Case 0

                    If miPrevTab Then
                        ReOrderBtnOnTab2 0, lngCntBtnPrevious, lngCntBtnTab, ctlScrollControl1(miIndex)
                    End If

                    ' Построение пакетов с новыми драйверами
                Case 1
                    ReOrderBtnOnTab2 1, lngCntBtnPrevious, lngCntBtnTab, ctlScrollControlTab1(miIndex)

                    ' Построение пакетов с неустановленными драйверами
                Case 2
                    ReOrderBtnOnTab2 2, lngCntBtnPrevious, lngCntBtnTab, ctlScrollControlTab2(miIndex)

                    ' Построение пакетов с установленными драйверами
                Case 3
                    ReOrderBtnOnTab2 3, lngCntBtnPrevious, lngCntBtnTab, ctlScrollControlTab3(miIndex)

                    ' Построение пакетов с "БД не создана"
                Case 4
                    ' Если есть пакеты без БД, тогда
                    If mbNotSupportedDevDB Then
                        ' Переключаемся в режим создания БД
                        mbSet2UpdateFromTab4 = True
                        SelectStartMode 3, False
                    End If
                    
                    ReOrderBtnOnTab2 4, lngCntBtnPrevious, lngCntBtnTab, ctlScrollControlTab4(miIndex)
                    mbSet2UpdateFromTab4 = False
            End Select

        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub TabInstBlockOnUpdate
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbBlock (Boolean)
'!--------------------------------------------------------------------------------
Private Sub TabInstBlockOnUpdate(ByVal mbBlock As Boolean)

    Dim i As Long

    For i = SSTab2.LBound To SSTab2.UBound

        If SSTab1.TabVisible(i) Then

            With SSTab2(i)
                .TabEnabled(1) = Not mbBlock
                .TabEnabled(2) = Not mbBlock
                .TabEnabled(3) = Not mbBlock
            End With

        End If

    Next

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ToolTipBtnReLoad
'! Description (Описание)  :   [Перезагрузка всплывающих подсказок для кнопок с драйверами]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ToolTipBtnReLoad()
    If mbDebugStandart Then DebugMode str2VbTab & "ToolTipBtnReLoad-Start"

    'Если подсказки уже созданы, то очистка
    If TT.Tools.Count Then
        TT.Tools.Clear
        TT.Title = strTTipTextTitle
    End If

    ' Обновляем всплывающую подсказку
    UpdateStatusButtonAll True
    If mbDebugStandart Then DebugMode str2VbTab & "ToolTipBtnReLoad-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ToolTipBtnReLoad
'! Description (Описание)  :   [Перезагрузка всплывающих подсказок для остальных контролов]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ToolTipOtherControlReLoad()
    If mbDebugStandart Then DebugMode str2VbTab & "ToolTipOtherControlReLoad-Start"

    With TTOtherControl
    'Если подсказки уже созданы, то очистка
        If .Tools.Count Then
            .Tools.Clear
            .Font.Name = strFontMainForm_Name
            .Font.Size = lngFontMainForm_Size
        End If
        ' Обновляем всплывающую подсказку
        .Tools.Add optRezim_Intellect.hWnd, , LocaliseString(strPCLangCurrentPath, strFormName, "RezimIntellectTip", optRezim_Intellect.ToolTipText)
        .Tools.Add optRezim_Ust.hWnd, , LocaliseString(strPCLangCurrentPath, strFormName, "RezimUstTip", optRezim_Ust.ToolTipText)
        .Tools.Add optRezim_Upd.hWnd, , LocaliseString(strPCLangCurrentPath, strFormName, "RezimUpdTip", optRezim_Upd.ToolTipText)
    End With

    If mbDebugStandart Then DebugMode str2VbTab & "ToolTipOtherControlReLoad-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ToolTipStatusLoad
'! Description (Описание)  :   [Загрузка статусных соощений]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ToolTipStatusLoad()

    Dim arrTTipStatusIconTemp() As String

    ReDim arrTTipStatusIcon(8)
    ReDim arrTTipStatusIconTemp(8)

    If mbDebugStandart Then DebugMode "ToolTipStatusLoad-Start"
    arrTTipStatusIconTemp(0) = "В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере." & str2vbNewLine & "Такие же драйвера (тех же версий) уже установлены на вашем компьютере." & _
                                str2vbNewLine & "Ваши действия:" & vbNewLine & "Никаких действий не требуется. " & str2vbNewLine & "Примечание:" & vbNewLine & "Если в колонке статус для устроуства стоит '0', то это означает:" & vbNewLine & _
                                " * - устройство блокировано;" & vbNewLine & " * - драйвер для данного устройства не активен (см. сведения в диспетчере устройств)"
    arrTTipStatusIconTemp(1) = "В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере." & str2vbNewLine & "На вашем компьютере эти драйвера не установлены." & str2vbNewLine & _
                                "Ваши действия:" & vbNewLine & _
                                "Переключите программу в один из режимов установки драйверов и нажмите на эту кнопку - это приведет к установке необходимых драйверов из пакета и соответствующему изменению вида кнопки." & str2vbNewLine & _
                                "Примечания:" & vbNewLine & _
                                "1. В некоторых случаях драйвера, из обнаруженных в пакете драйверов, могут не подойти к вашему оборудованию. Сравнение идентификаторов устройств (HWID) происходит по основной части HWID без учета подклассов устройств (SUBSYS|REV|MI)." _
                                & vbNewLine & "2. Если в колонке статус для устроуства стоит '0', то это означает:" & vbNewLine & " * - драйвер для данного устройства не установлен;" & vbNewLine & " * - устройство блокировано;" & vbNewLine & _
                                " * - драйвер для данного устройства не активен (см. сведения в диспетчере устройств)"
    arrTTipStatusIconTemp(2) = "В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере, но более новые, чем те, что уже установлены." & str2vbNewLine & "Ваши действия:" & vbNewLine & _
                                "Переключите программу в один из режимов установки драйверов и нажмите на эту кнопку - это приведет к установке более новых драйверов из пакета и соответствующему изменению вида кнопки." & str2vbNewLine & _
                                "Примечание:" & vbNewLine & _
                                "В некоторых случаях драйвера, из обнаруженных в пакете драйверов, могут не подойти к вашему оборудованию. Сравнение идентификаторов устройств (HWID) происходит по основной части HWID без учета подклассов устройств (SUBSYS|REV|MI)."
    arrTTipStatusIconTemp(3) = "В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере, но более старые, чем те, что уже установлены." & str2vbNewLine & "Ваши действия:" & vbNewLine & _
                                "Ничего делать не надо. Можете поискать в сети более свежие драйвера и обновить (заменить) данный пакет в программе."
    arrTTipStatusIconTemp(4) = "1. В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере, но более новые, чем те, что уже установлены." & str2vbNewLine & "Ваши действия:" & vbNewLine & _
                                "Переключите программу в один из режимов установки драйверов и нажмите на эту кнопку - это приведет к установке более новых драйверов из пакета и соответствующему изменению вида кнопки." & str2vbNewLine & _
                                "2. В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере." & str2vbNewLine & "На вашем компьютере эти драйвера не установлены." & str2vbNewLine _
                                & "Ваши действия:" & vbNewLine & _
                                "Переключите программу в один из режимов установки драйверов и нажмите на эту кнопку - это приведет к установке необходимых драйверов из пакета и соответствующему изменению вида кнопки." & str2vbNewLine & _
                                "Примечания:" & vbNewLine & _
                                "1. В некоторых случаях драйвера, из обнаруженных в пакете драйверов, могут не подойти к вашему оборудованию. Сравнение идентификаторов устройств (HWID) происходит по основной части HWID без учета подклассов устройств (SUBSYS|REV|MI)." _
                                & vbNewLine & "2. Если в колонке статус для устроуства стоит '0', то это означает:" & vbNewLine & " * - драйвер для данного устройства не установлен;" & vbNewLine & " * - устройство блокировано;" & vbNewLine & _
                                " * - драйвер для данного устройства не активен (см. сведения в диспетчере устройств)"
    arrTTipStatusIconTemp(5) = "1. В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере, но более старые, чем те, что уже установлены." & str2vbNewLine & "Ваши действия:" & vbNewLine & _
                                "Ничего делать не надо. Можете поискать в сети более свежие драйвера и обновить (заменить) данный пакет в программе." & str2vbNewLine & _
                                "2. В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере." & str2vbNewLine & "На вашем компьютере эти драйвера не установлены." & str2vbNewLine _
                                & "Ваши действия:" & vbNewLine & _
                                "Переключите программу в один из режимов установки драйверов и нажмите на эту кнопку - это приведет к установке необходимых драйверов из пакета и соответствующему изменению вида кнопки." & str2vbNewLine & _
                                "Примечания:" & vbNewLine & _
                                "1. В некоторых случаях драйвера, из обнаруженных в пакете драйверов, могут не подойти к вашему оборудованию. Сравнение идентификаторов устройств (HWID) происходит по основной части HWID без учета подклассов устройств (SUBSYS|REV|MI)." _
                                & vbNewLine & "2. Если в колонке статус для устроуства стоит '0', то это означает:" & vbNewLine & " * - драйвер для данного устройства не установлен;" & vbNewLine & " * - устройство блокировано;" & vbNewLine & _
                                " * - драйвер для данного устройства не активен (см. сведения в диспетчере устройств)"
    arrTTipStatusIconTemp(6) = "Драйвера из этого пакета программы не имеют отношения к вашему компьютеру." & str2vbNewLine & "Ваши действия:" & vbNewLine & _
                                "Ничего делать не надо. Этот пакет драйверов пригодится вам как-нибудь в другой раз - при замене устройств или на другом компьютере."
    arrTTipStatusIconTemp(7) = "Программа не может определить, что находится в этом пакете драйверов." & str2vbNewLine & "Ваши действия:" & vbNewLine & _
                                "Переключите программу в режим 'Создание или обновление базы данных драйверов', нажмите на эту кнопку - таким образом сведения о драйверах из пакета будут добавлены в базу данных программы и вид кнопки изменится. О дальнейших действиях читайте в пояснении к соответствующему значку."
    arrTTipStatusIconTemp(8) = "Программа производит обновление базы драйверов для выбранного пакета драйверов." & str2vbNewLine & "Ваши действия:" & vbNewLine & "Ничего делать не надо. Остается только ждать завершения работы программы."
    arrTTipStatusIcon(0) = LocaliseString(strPCLangCurrentPath, strFormName, "ToolTipArrStatusIcon1", arrTTipStatusIconTemp(0))
    arrTTipStatusIcon(1) = LocaliseString(strPCLangCurrentPath, strFormName, "ToolTipArrStatusIcon2", arrTTipStatusIconTemp(1))
    arrTTipStatusIcon(2) = LocaliseString(strPCLangCurrentPath, strFormName, "ToolTipArrStatusIcon3", arrTTipStatusIconTemp(2))
    arrTTipStatusIcon(3) = LocaliseString(strPCLangCurrentPath, strFormName, "ToolTipArrStatusIcon4", arrTTipStatusIconTemp(3))
    arrTTipStatusIcon(4) = LocaliseString(strPCLangCurrentPath, strFormName, "ToolTipArrStatusIcon5", arrTTipStatusIconTemp(4))
    arrTTipStatusIcon(5) = LocaliseString(strPCLangCurrentPath, strFormName, "ToolTipArrStatusIcon6", arrTTipStatusIconTemp(5))
    arrTTipStatusIcon(6) = LocaliseString(strPCLangCurrentPath, strFormName, "ToolTipArrStatusIcon7", arrTTipStatusIconTemp(6))
    arrTTipStatusIcon(7) = LocaliseString(strPCLangCurrentPath, strFormName, "ToolTipArrStatusIcon8", arrTTipStatusIconTemp(7))
    arrTTipStatusIcon(8) = LocaliseString(strPCLangCurrentPath, strFormName, "ToolTipArrStatusIcon9", arrTTipStatusIconTemp(8))

    ' Изменяем параметры Всплывающей подсказки для статусных картинок
    With TTStatusIcon

        'Если уже созданы, то очистка
        If .Tools.Count Then
            .Tools.Clear
        End If

        .Font.Name = strFontMainForm_Name
        .Font.Size = lngFontMainForm_Size
        .MaxTipWidth = Me.Width
        .SetDelayTime TipDelayTimeInitial, 200
        .SetDelayTime TipDelayTimeShow, 15000
        .Tools.Add imgOK.hWnd, , arrTTipStatusIcon(0)
        .Tools.Add imgOkAttention.hWnd, , arrTTipStatusIcon(1)
        .Tools.Add imgOkNew.hWnd, , arrTTipStatusIcon(2)
        .Tools.Add imgOkOld.hWnd, , arrTTipStatusIcon(3)
        .Tools.Add imgOkAttentionNew.hWnd, , arrTTipStatusIcon(4)
        .Tools.Add imgOkAttentionOLD.hWnd, , arrTTipStatusIcon(5)
        .Tools.Add imgNo.hWnd, , arrTTipStatusIcon(6)
        .Tools.Add imgNoDB.hWnd, , arrTTipStatusIcon(7)
        .Tools.Add imgUpdBD.hWnd, , arrTTipStatusIcon(8)
    End With

    Erase arrTTipStatusIconTemp
    If mbDebugStandart Then DebugMode "ToolTipStatusLoad-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UnloadAllForms
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   FormToIgnore (String = vbNullString)
'!--------------------------------------------------------------------------------
Public Sub UnloadAllForms(Optional FormToIgnore As String = vbNullString)

    Dim F As Form

    For Each F In Forms

        If Not F Is Nothing Then
            If StrComp(F.Name, FormToIgnore, vbTextCompare) <> 0 Then
                Unload F
                Set F = Nothing
            End If
        End If

    Next F

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function UnPackDPFile
'! Description (Описание)  :   [Извлечение файлов из архива]
'! Parameters  (Переменные):   strPathDRP (String)
'                              strPackFileName (String)
'                              strMaskFile (String)
'                              strDest4OnlyUnpack (String)
'!--------------------------------------------------------------------------------
Private Function UnPackDPFile(ByVal strPathDRP As String, ByVal strPackFileName As String, ByVal strMaskFile As String, ByVal strDest4OnlyUnpack As String) As Boolean

    Dim WorkDir               As String
    Dim strPackGetFileName_woExt As String
    Dim cmdString             As String
    Dim ArchTempPath          As String
    Dim strPhysXPath          As String
    Dim strLangPath           As String
    Dim strRuntimes           As String
    Dim strClassesName        As String
    Dim strFinishIniPath      As String
    Dim ret                   As Long
    Dim strMaskFile_x()       As String
    Dim i                     As Long
    Dim strMaskFile_x_TEMP    As String
    Dim strMaskFile_x_TEMPTo  As String
    Dim strMaskFile_xx()      As String

    If mbDebugStandart Then DebugMode "UnPackDPFile-Start" & vbNewLine & _
              "UnPackDPFile: strMaskFile=" & strMaskFile

    If Not mbOnlyUnpackDP Then
        strPackGetFileName_woExt = GetFileName_woExt(strPackFileName)

        'Рабочий каталог
        If mbGroupTask Then
            WorkDir = strWorkTempBackSL & "GroupInstall\"
            ArchTempPath = strWorkTempBackSL & "GroupInstall"
        Else
            WorkDir = BackslashAdd2Path(strWorkTempBackSL & strPackGetFileName_woExt)
            ArchTempPath = strWorkTempBackSL & strPackGetFileName_woExt

            If PathExists(WorkDir) Then
                DelRecursiveFolder (WorkDir)
            End If
        End If

    Else
        ArchTempPath = strDest4OnlyUnpack
    End If

    If Not mbDP_Is_aFolder Then
        cmdString = strKavichki & strArh7zExePATH & strKavichki & " x -yo" & strKavichki & ArchTempPath & strKavichki & " -r " & strKavichki & strPathDRP & strPackFileName & strKavichki & strSpace & strMaskFile
        ChangeStatusTextAndDebug strMessages(97) & strSpace & strPackFileName
        If mbDebugStandart Then DebugMode "Extract: " & cmdString

        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
            If Not mbSilentRun Then
                MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
            End If

            UnPackDPFile = False
            ChangeStatusTextAndDebug strMessages(13) & strSpace & strPackFileName
            If mbDebugStandart Then DebugMode "Error on run : " & cmdString
        Else

            'Распаковка дополнительных файлов
            ' Если класс пакета считывается при запуске программы, то
            ' Архиватор отработал на все 100%? Если нет то сообщаем
            If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
                ChangeStatusTextAndDebug strMessages(13) & strSpace & strPackFileName

                If Not mbSilentRun Then
                    MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                End If

            Else
                strClassesName = vbNullString

                If mbReadClasses Then
                    strFinishIniPath = PathCombine(arrOSList(SSTab1.Tab).devIDFolderFull, GetFileName_woExt(strPackFileName) & ".ini")
                    strClassesName = IniStringPrivate("DriverPack", "classes", strFinishIniPath)

                    ' Если такого значения в файле нет, то ничего не добавляем
                    If StrComp(strClassesName, "no_key") = 0 Then
                        strClassesName = vbNullString
                    End If

                    If LenB(strClassesName) Then

                        ' Если класс пакета определен как Display, то
                        If StrComp(strClassesName, "Display", vbTextCompare) = 0 Then
                            If Not mbGroupTask Then

                                ' Распаковка strPhysXPath
                                If LenB(arrOSList(SSTab1.Tab).PathPhysX) Then
                                    strPhysXPath = PathCollect(arrOSList(SSTab1.Tab).PathPhysX)
                                    UnPackDPFileAdd strPhysXPath, strPathDRP, ArchTempPath
                                End If

                                ' Распаковка strLangPath
                                If LenB(arrOSList(SSTab1.Tab).PathLanguages) Then
                                    strLangPath = PathCollect(arrOSList(SSTab1.Tab).PathLanguages)
                                    UnPackDPFileAdd strLangPath, strPathDRP, ArchTempPath
                                End If

                                ' Распаковка strRuntimes
                                If LenB(arrOSList(SSTab1.Tab).PathRuntimes) Then
                                    strRuntimes = PathCollect(arrOSList(SSTab1.Tab).PathRuntimes)
                                    UnPackDPFileAdd strRuntimes, strPathDRP, ArchTempPath
                                End If

                            Else
                                mbUnpackAdditionalFile = True
                            End If
                        End If
                    End If
                End If

                UnPackDPFile = True
            End If
        End If

    Else
        ChangeStatusTextAndDebug strMessages(149) & strSpace & strPackFileName
        If mbDebugStandart Then DebugMode "Copy: " & strMaskFile

        If PathExists(WorkDir) = False Then
            CreateNewDirectory WorkDir
        End If

        If InStr(strMaskFile, strSpace) Then
            strMaskFile_x = Split(strMaskFile, strSpace)

            For i = 0 To UBound(strMaskFile_x)
                strMaskFile_x_TEMP = BackslashDelFromPath(strMaskFile_x(i))
                strMaskFile_xx = Split(strMaskFile_x_TEMP, vbBackslash)

                If UBound(strMaskFile_xx) > 1 Then
                    strMaskFile_x_TEMPTo = Left$(strMaskFile_x_TEMP, InStrRev(strMaskFile_x_TEMP, vbBackslash) - 1)
                End If

                ret = ret + CopyFolderByShell(BackslashAdd2Path(strPathDRP & strPackFileName) & strMaskFile_x_TEMP, BackslashAdd2Path(ArchTempPath) & strMaskFile_x_TEMPTo)
            Next

        Else
            strMaskFile_x_TEMP = BackslashDelFromPath(strMaskFile)
            strMaskFile_xx = Split(strMaskFile_x_TEMP, vbBackslash)

            If UBound(strMaskFile_xx) > 1 Then
                strMaskFile_x_TEMPTo = Left$(strMaskFile_x_TEMP, InStrRev(strMaskFile_x_TEMP, vbBackslash) - 1)
            End If

            ret = CopyFolderByShell(BackslashAdd2Path(strPathDRP & strPackFileName) & strMaskFile, BackslashAdd2Path(ArchTempPath) & strMaskFile_x_TEMPTo)
        End If

        UnPackDPFile = Not Abs(ret)
        If mbDebugStandart Then DebugMode "UnPackDPFile-Copy files: " & UnPackDPFile
    End If

    If mbDebugStandart Then DebugMode "UnPackDPFile-End"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UnPackDPFileAdd
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPathAddFile (String)
'                              strPathDRP (String)
'                              strArchTempPath (String)
'!--------------------------------------------------------------------------------
Private Sub UnPackDPFileAdd(ByVal strPathAddFile As String, ByVal strPathDRP As String, ByVal strArchTempPath As String)

    Dim cmdString As String
    Dim strPathAddFilePath As String

    strPathAddFilePath = PathCombine(strPathDRP, strPathAddFile)

    If PathExists(strPathAddFilePath) Then
        If Not PathIsAFolder(strPathAddFilePath) Then
            cmdString = strKavichki & strArh7zExePATH & strKavichki & " x -yo" & strKavichki & strArchTempPath & strKavichki & " -r " & strKavichki & strPathAddFilePath & strKavichki & " *.*"
            ChangeStatusTextAndDebug strMessages(98) & strSpace & strPathAddFilePath
            If mbDebugStandart Then DebugMode "Extract: " & cmdString

            If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
                If Not mbSilentRun Then
                    MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                End If

                ChangeStatusTextAndDebug strMessages(13) & strSpace & strPathAddFilePath
                If mbDebugStandart Then DebugMode "Error on run : " & cmdString
            Else

                ' Архиватор отработал на все 100%? Если нет то сообщаем
                If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
                    ChangeStatusTextAndDebug strMessages(13) & strSpace & strPathAddFilePath

                    If Not mbSilentRun Then
                        MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                    End If
                End If
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function UnpackOtherFile
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strArcDRPPath (String)
'                              strWorkDir (String)
'                              strMaskFile (String)
'!--------------------------------------------------------------------------------
Private Function UnpackOtherFile(ByVal strArcDRPPath As String, ByVal strWorkDir As String, ByVal strMaskFile As String) As Boolean

    Dim cmdString As String

    If mbDebugStandart Then DebugMode "UnpackOtherFile-Start" & vbNewLine & _
              "UnpackOtherFile: strArcDRPPath=" & strArcDRPPath & vbNewLine & _
              "UnpackOtherFile: strMaskFile=" & strMaskFile
     
    cmdString = strKavichki & strArh7zExePATH & strKavichki & " x -yo" & strKavichki & strWorkDir & strKavichki & " -r " & strKavichki & strArcDRPPath & strKavichki & strSpace & strMaskFile
    ChangeStatusTextAndDebug strMessages(99) & strSpace & strArcDRPPath
    If mbDebugStandart Then DebugMode "Extract: " & cmdString
    UnpackOtherFile = True

    If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
        If Not mbSilentRun Then
            MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
        End If

        ChangeStatusTextAndDebug strMessages(13) & strSpace & cmdString
        If mbDebugStandart Then DebugMode "Error on run : " & cmdString
        UnpackOtherFile = False
    Else

        ' Архиватор отработал на все 100%? Если нет то сообщаем
        If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
            ChangeStatusTextAndDebug strMessages(13) & strSpace & GetFileNameFromPath(strArcDRPPath)

            If Not mbSilentRun Then
                MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
            End If

            UnpackOtherFile = False
        End If
    End If

    If mbDebugStandart Then DebugMode "UnpackOtherFile-End"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UpdateStatusButtonAll
'! Description (Описание)  :   [Обновление всех статусов]
'! Parameters  (Переменные):   mbReloadTT (Boolean = False)
'!--------------------------------------------------------------------------------
Public Sub UpdateStatusButtonAll(Optional mbReloadTT As Boolean = False)

    Dim ButtIndex        As Long
    Dim ButtCount        As Long
    Dim i                As Integer
    Dim i_Tab            As Integer
    Dim TimeScriptRun    As Long
    Dim TimeScriptFinish As Long
    Dim AllTimeScriptRun As String
    Dim miPbInterval     As Long
    Dim miPbNext         As Long
    Dim mbDpNoDBExist    As Boolean
    Dim lngSStabStart    As Long
    Dim strPackFileName  As String
    Dim strPathDRP       As String
    Dim strPathDevDB     As String
    Dim lngTabN          As Long
    Dim lngNumButtOnTab  As Long

    If mbDebugStandart Then DebugMode "StatusUpdateAll-Start"
    lngSStabStart = SSTab1.Tab
    ctlUcStatusBar1.PanelText(1) = strMessages(127)

    ' Если кнопка всего одна, то проверяем на какой она вкладке и устанавливаем ее активной
    If acmdPackFiles.Count = 1 Then
        If acmdPackFiles(0).Visible Then

            With SSTab1

                For i_Tab = 0 To .Tabs - 1

                    If .TabVisible(i_Tab) Then
                        .Tab = i_Tab

                        If StrComp(acmdPackFiles(0).Container.Name, "ctlScrollControl1", vbTextCompare) = 0 Then
                            If acmdPackFiles(0).Container.Index = .Tab Then

                                Exit For

                            End If
                        End If
                    End If

                Next

            End With

        End If

    Else
        i_Tab = 0

        If LenB(chkPackFiles(0).Tag) Then
            i_Tab = chkPackFiles(0).Tag
        End If
    End If

    BlockControl False
    DoEvents
    SSTab1.Tab = i_Tab
    TimeScriptRun = 0
    AllTimeScriptRun = vbNullString
    TimeScriptRun = GetTickCount
    ButtIndex = acmdPackFiles.UBound
    ButtCount = acmdPackFiles.Count
    ' Отображаем ProgressBar
    CreateProgressNew
        
    If ButtIndex Then
        ' В цикле обрабатываем обновление
        miPbInterval = 1000 / ButtCount
        miPbNext = 0

        For i = 0 To ButtIndex
            lngTabN = SSTab1.Tab
            lngNumButtOnTab = arrOSList(SSTab1.Tab).CntBtn

            Do While i >= lngNumButtOnTab
                lngTabN = lngTabN + 1
                SSTab1.Tab = lngTabN
                DoEvents
                lngNumButtOnTab = arrOSList(SSTab1.Tab).CntBtn
            Loop

            mbDpNoDBExist = True
            strPathDRP = arrOSList(lngTabN).drpFolderFull
            strPathDevDB = arrOSList(lngTabN).devIDFolderFull

            With acmdPackFiles(i)

                If Not mbReloadTT Then
                    ' Кнопка выглядит нажатой
                    .Value = True
                    
                    Set .PictureNormal = imgUpdBD.Picture
                                        
                    strPackFileName = .Tag
                    ChangeStatusTextAndDebug "(" & i + 1 & strSpace & strMessages(124) & strSpace & ButtCount & "): " & strMessages(89) & strSpace & strPackFileName
                    ' Обновление подсказки
                    ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, i, True
                Else
                    strPackFileName = .Tag
                    ' Только обновление подсказки (используется при смене языка, для изменения шапки таблицы в подсказке)
                    ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, i, , True
                End If

            End With

            miPbNext = miPbNext + miPbInterval

            If miPbNext > 1000 Then
                miPbNext = 1000
            End If

            With ctlProgressBar1
                .Value = miPbNext
                .SetTaskBarProgressValue miPbNext, 1000
            End With

            ChangeFrmMainCaption miPbNext
        Next

    Else

        If Not mbReloadTT Then
            acmdPackFiles_Click 0
        End If

        mbDpNoDBExist = True
    End If

    ' подсчет времни выполнения и формирования сообщения в статусе
    TimeScriptFinish = GetTickCount
    AllTimeScriptRun = CalculateTime(TimeScriptRun, TimeScriptFinish)

    If mbDpNoDBExist Then
        ChangeStatusTextAndDebug strMessages(67) & strSpace & AllTimeScriptRun
    Else
        ChangeStatusTextAndDebug strMessages(68)
    End If

    pbProgressBar.Visible = False
    ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
    ChangeFrmMainCaption
    BlockControl True
    
TheEnd:
    SSTab1.Tab = lngSStabStart
    If mbDebugStandart Then DebugMode "StatusUpdateAll-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UpdateStatusButtonTAB
'! Description (Описание)  :   [Обновление всех статусов]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub UpdateStatusButtonTAB()

    Dim i                 As Integer
    Dim TimeScriptRun     As Long
    Dim TimeScriptFinish  As Long
    Dim AllTimeScriptRun  As String
    Dim miPbInterval      As Long
    Dim miPbNext          As Long
    Dim mbDpNoDBExist     As Boolean
    Dim strPackFileName   As String
    Dim strPathDRP        As String
    Dim strPathDevDB      As String
    Dim lngCntBtnTab      As Long
    Dim lngCntBtnPrevious As Long
    Dim lngSSTab1Tab      As Long
    Dim lngCurrBtn        As Long
    Dim lngSummBtn        As Long

    If mbDebugStandart Then DebugMode "UpdateStatusButtonTAB-Start"
    BlockControl False
    ctlUcStatusBar1.PanelText(1) = strMessages(127)
    DoEvents
    AllTimeScriptRun = vbNullString
    TimeScriptRun = GetTickCount
    
    ' Отображаем ProgressBar
    CreateProgressNew

    With SSTab1
        lngSSTab1Tab = .Tab

        If lngSSTab1Tab Then
            lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 1).CntBtn

            If lngCntBtnPrevious = 0 Then
                If lngSSTab1Tab > 1 Then
                    lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 2).CntBtn
                End If
            End If
        End If

    End With

    lngCntBtnTab = arrOSList(lngSSTab1Tab).CntBtn - 1

    If lngCntBtnTab Then
        ' В цикле обрабатываем обновление
        lngSummBtn = lngCntBtnTab - lngCntBtnPrevious
        miPbInterval = 1000 / lngSummBtn
        miPbNext = 0

        For i = lngCntBtnPrevious To lngCntBtnTab
            lngCurrBtn = lngCurrBtn + 1
            mbDpNoDBExist = True
            strPathDRP = arrOSList(lngSSTab1Tab).drpFolderFull
            strPathDevDB = arrOSList(lngSSTab1Tab).devIDFolderFull

            With acmdPackFiles(i)
                ' Кнопка выглядит нажатой
                .Value = True
                
                Set .PictureNormal = imgUpdBD.Picture
                
                strPackFileName = .Tag
                ChangeStatusTextAndDebug "(" & lngCurrBtn & strSpace & strMessages(124) & strSpace & lngSummBtn & "): " & strMessages(89) & strSpace & strPackFileName
                ' Обновление подсказки
                ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, i
                                
            End With

            miPbNext = miPbNext + miPbInterval

            If miPbNext > 1000 Then
                miPbNext = 1000
            End If

            With ctlProgressBar1
                .Value = miPbNext
                .SetTaskBarProgressValue miPbNext, 1000
            End With

            ChangeFrmMainCaption miPbNext
        Next

    Else
        mbDpNoDBExist = True
    End If

    ' подсчет времни выполнения и формирования сообщения в статусе
    TimeScriptFinish = GetTickCount
    AllTimeScriptRun = CalculateTime(TimeScriptRun, TimeScriptFinish)

    If mbDpNoDBExist Then
        ChangeStatusTextAndDebug strMessages(67) & strSpace & AllTimeScriptRun
    Else
        ChangeStatusTextAndDebug strMessages(68)
    End If

    ChangeFrmMainCaption
    pbProgressBar.Visible = False
    ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
    BlockControl True
TheEnd:
    If mbDebugStandart Then DebugMode "UpdateStatusButtonTAB-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub VerModules
'! Description (Описание)  :   [Отображение версий модулей]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub VerModules()
    MsgBox strMessages(35) & vbNewLine & "DPinst.exe (x86)" & vbTab & GetFileVersionOnly(strDPInstExePath) & vbNewLine & "DPinst.exe (x64)" & vbTab & GetFileVersionOnly(strDPInstExePath64) & vbNewLine & "DevCon.exe (x86)" & vbTab & _
                                GetFileVersionOnly(strDevConExePath) & vbNewLine & "7za.exe (x86)" & vbTab & GetFileVersionOnly(strArh7zExePATH), vbInformation, strProductName
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub WorkWithFinish
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPathDRP (String)
'                              strPackFileName (String)
'                              strWorkPath (String)
'                              strPathDRPList (String)
'!--------------------------------------------------------------------------------
Private Sub WorkWithFinish(ByVal strPathDRP As String, ByVal strPackFileName As String, ByVal strWorkPath As String, ByVal strPathDRPList As String)

    Dim strPathDRPList_x() As String
    Dim strSectionName     As String
    Dim strFinishIniPath   As String
    Dim lngEXCCount        As Long
    Dim i                  As Long
    Dim ii                 As Long

    If mbDebugStandart Then DebugMode "WorkWithFinish-Start"

    If mbLoadFinishFile Then
        If strPathDRPList <> ALL_FILES Then
            strPathDRPList_x = Split(strPathDRPList, strSpace)

            For ii = 0 To UBound(strPathDRPList_x)
                strSectionName = GetFileNameFromPath(BackslashDelFromPath(strPathDRPList_x(ii)))
                ChangeStatusTextAndDebug strMessages(100) & " '" & strSectionName & "'"
                strFinishIniPath = PathCombine(arrOSList(SSTab1.Tab).devIDFolderFull, GetFileName_woExt(strPackFileName) & ".ini")

                If PathExists(strFinishIniPath) Then
                    If Not PathIsAFolder(strFinishIniPath) Then
                        lngEXCCount = IniLongPrivate(strSectionName, "exc_count", strFinishIniPath)

                        ' Если такого значения в файле нет, то ничего не добавляем
                        If lngEXCCount = "9999" Then
                            lngEXCCount = 0
                        End If

                        If lngEXCCount Then

                            For i = 1 To lngEXCCount
                                FindAndInstallPanel strPathDRP & strPackFileName, strFinishIniPath, strSectionName, i, strWorkPath
                            Next

                        End If
                    End If
                End If

            Next

        End If
    End If

    If mbDebugStandart Then DebugMode "WorkWithFinish-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub acmdPackFiles_Click
'! Description (Описание)  :   [Обработка События нажатия кнопки]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub acmdPackFiles_Click(Index As Integer)

    Dim strPackFileName       As String
    Dim strPathDRP            As String
    Dim strPathDevDB          As String
    Dim TimeScriptRun         As Long
    Dim TimeScriptFinish      As Long
    Dim AllTimeScriptRun      As String
    Dim strFileName_woExt     As String
    Dim cmdString             As String
    Dim ArchTempPath          As String
    Dim strDevPathShort       As String
    Dim DPInstExitCode        As Long
    Dim ReadExitCodeString    As String
    Dim strTemp_x()           As String
    Dim strTempLine_x()       As String
    Dim i_arr                 As Long
    Dim lngRetMsgBox          As Long

    If mbDebugStandart Then DebugMode "acmdPackFiles_Click-Start: Index Button=" & Index
               
    strPathDRPList = vbNullString
    
    ' Блокируем форму при обработке пакета
    If Not mbGroupTask Then
        BlockControl False
    End If

    ' Запущен ли другой процесс обработки пакета?
    If mbDevParserRun Then
        MsgBox strMessages(22), vbInformation, strProductName
    Else
        mbStatusHwid = True
        strPackFileName = acmdPackFiles(Index).Tag

        'Если пакет драйверов реальный, то....
        If LenB(strPackFileName) Then
            
            acmdPackFiles(Index).Value = True
            strPathDRP = arrOSList(SSTab1.Tab).drpFolderFull
            strPathDevDB = arrOSList(SSTab1.Tab).devIDFolderFull
            mbDevParserRun = True
            lngExitProc = 0

            '------------------------------------------------------
            '---------------- Режим обновления БД -----------------
            '------------------------------------------------------
            If optRezim_Upd.Value Then
                If mbIsDriveCDRoom Then
                    MsgBox strMessages(16), vbInformation, strProductName
                Else
                    TimeScriptRun = 0
                    TimeScriptFinish = 0
                    AllTimeScriptRun = vbNullString
                    Set acmdPackFiles(Index).PictureNormal = imgUpdBD.Picture
                    ' изменение положения прогресс-анимации
                    ChangeStatusTextAndDebug strMessages(61) & strSpace & strPackFileName, strMessages(128)
                    TimeScriptRun = GetTickCount
                    ' запуск построения БД
                    DevParserByRegExp strPackFileName, strPathDRP, strPathDevDB
                    ' Обновление подсказки
                    ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, Index, True
                    TimeScriptFinish = GetTickCount
                    AllTimeScriptRun = CalculateTime(TimeScriptRun, TimeScriptFinish, True)
                    ChangeStatusTextAndDebug strMessages(62) & strSpace & AllTimeScriptRun
                    If mbDebugStandart Then DebugMode vbTab & "Create Index: All time for create Base for file: " & AllTimeScriptRun
                    
                    If Not mbGroupTask Then
                        ' Обновить список неизвестных дров и описание для кнопки
                        LoadCmdViewAllDeviceCaption
                    End If
                End If

                '------------------------------------------------------
                '-------- Установка всех драйверов в пакете -----------
                '------------------------------------------------------
            ElseIf optRezim_Ust.Value Then
                                
                ChangeStatusTextAndDebug strMessages(63) & strSpace & strPackFileName, strMessages(129)
                'Имя папки с распакованными драйверами
                strFileName_woExt = GetFileName_woExt(strPackFileName)
                ArchTempPath = strWorkTempBackSL & strFileName_woExt

                ' Создаем точку восстановления
                If Not mbOnlyUnpackDP Then
                    If mbCreateRestorePoint Then
                        ' Проверяем создавалась ли точка восстановления ранее
                        If Not mbCreateRestorePointDone Then
                            If mbSilentRun Then
                                CreateRestorePoint
                            Else
                                lngRetMsgBox = MsgBox(strMessages(115) & vbNewLine & strMessages(120) & str2vbNewLine & strMessages(153), vbQuestion + vbYesNoCancel, strProductName)
                                ' Click "Yes"
                                If lngRetMsgBox = vbYes Then
                                    CreateRestorePoint
                                ' Click "Cancel" - Do not remind
                                ElseIf lngRetMsgBox = vbCancel Then
                                    mbCreateRestorePointDone = True
                                End If
                            End If
                        End If
                    End If
                End If
                
                'Извлечение драйверов из файла
                If UnPackDPFile(strPathDRP, strPackFileName, ALL_FILES, ArchTempPath) = False Then
                    If Not mbSilentRun Then
                        MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                    End If

                    ChangeStatusTextAndDebug strMessages(13) & strSpace & strPackFileName
                    If mbDebugStandart Then DebugMode "Error on run : " & cmdString
                Else
                    ' установка драйверов
                    DPInstExitCode = RunDPInst(ArchTempPath)
                End If

                ' Обновление подсказки
                ReadExitCodeString = ReadExitCode(DPInstExitCode)

                If DPInstExitCode <> 0 Then
                    If DPInstExitCode <> -2147483648# Then
                        If InStr(1, ReadExitCodeString, "Cancel or Nothing to Install", vbTextCompare) = 0 Then
                            ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, Index, True
                        End If
                    End If
                End If
    
                ChangeStatusTextAndDebug strMessages(64) & " (" & strPackFileName & "): " & ReadExitCodeString
                If mbDebugStandart Then DebugMode "Install from : " & strPackFileName & " finished."

                '------------------------------------------------------
                '------- Установка избранных драйверов в пакете--------
                '------------------------------------------------------
            Else
                                
                'Имя папки с распакованными драйверами
                strFileName_woExt = GetFileName_woExt(strPackFileName)
                ' Временный путь распаковки
                ArchTempPath = strWorkTempBackSL & strFileName_woExt

                ' Создаем точку восстановления
                If Not mbOnlyUnpackDP Then
                    ChangeStatusTextAndDebug strMessages(63) & strSpace & strPackFileName, strMessages(129)
                    If mbCreateRestorePoint Then
                        ' Проверяем создавалась ли точка восстановления ранее
                        If Not mbCreateRestorePointDone Then
                            If mbSilentRun Then
                                CreateRestorePoint
                            Else
                                lngRetMsgBox = MsgBox(strMessages(115) & vbNewLine & strMessages(120) & str2vbNewLine & strMessages(153), vbQuestion + vbYesNoCancel, strProductName)
                                ' Click "Yes"
                                If lngRetMsgBox = vbYes Then
                                    CreateRestorePoint
                                ' Click "Cancel" - Do not remind
                                ElseIf lngRetMsgBox = vbCancel Then
                                    mbCreateRestorePointDone = True
                                End If
                            End If
                        End If
                    End If
                Else
                    ChangeStatusTextAndDebug strMessages(154) & strSpace & strPackFileName, strMessages(155)
                    '# Диалог выбора каталога
                    With New CommonDialog
                        .InitDir = ArchTempPath
                        .DialogTitle = strMessages(131)
                        .Flags = CdlBIFNewDialogStyle
        
                        If .ShowFolder Then
                            ArchTempPath = .FileName
                        Else
                            '# if user cancel #
                            ChangeStatusTextAndDebug strMessages(132) & strSpace & strPackFileName
                            mbDevParserRun = False
                            
                            Exit Sub
                
                        End If
        
                    End With
                    
                    If LenB(ArchTempPath) = 0 Then
                        '# if user cancel #
                        ChangeStatusTextAndDebug strMessages(132) & strSpace & strPackFileName
                        mbDevParserRun = False
                        
                        Exit Sub
        
                    End If
        
                    If mbDebugStandart Then DebugMode "Unpack: Destination=" & ArchTempPath

                End If
                
                ' если выборочная установка, то получаем список каталогов для распаковки
                If mbSelectInstall Then
                    If IsFormLoaded("frmListHwid") = False Then
                        frmListHwid.Show vbModal, Me
                    Else
                        frmListHwid.FormLoadDefaultParam
                        frmListHwid.FormLoadAction
                        frmListHwid.Show vbModal, Me
                    End If

                    ' если на форме нажали отмену или закрыли ее, то завершаем обработку
                    If Not mbCheckDRVOk Then
                        mbDevParserRun = False
                        
                        'acmdPackFiles(Index).Value = False
                        
                        If Not mbGroupTask Then
                            BlockControl True
                        End If
                        ChangeStatusTextAndDebug strMessages(65) & strSpace & strPackFileName
                        cmdRunTask.Enabled = FindCheckCount(False)

                        Exit Sub

                    End If

                Else

                    ' иначе список строится сам

                    strTemp_x = Split(arrTTip(Index), vbNewLine)

                    For i_arr = 0 To UBound(strTemp_x)
                        strTempLine_x = Split(strTemp_x(i_arr), " | ")

                        If LenB(Trim$(strTemp_x(i_arr))) Then
                            strDevPathShort = Trim$(GetPathNameFromPath(strTempLine_x(1)))

                            ' Если данного пути нет в списке, то добавляем
                            If InStr(1, strPathDRPList, strDevPathShort, vbTextCompare) = 0 Then
                                AppendStr strPathDRPList, strDevPathShort, strSpace
                            End If
                        End If

                    Next i_arr

                End If

                strPathDRPList = Trim$(strPathDRPList)

                ' Если по каким либо причинам список папок не получился, то извлекаем все.
                If LenB(strPathDRPList) = 0 Then
                    strPathDRPList = ALL_FILES
                End If

                'Извлечение драйверов из файла
                If UnPackDPFile(strPathDRP, strPackFileName, strPathDRPList, ArchTempPath) = False Then
                    If Not mbSilentRun Then
                        MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                    End If

                    ChangeStatusTextAndDebug strMessages(13) & strSpace & strPackFileName
                    If mbDebugStandart Then DebugMode "Error on run : " & cmdString
                Else
                    ' Если выбран режим, не только распаковки, то запускаем установку
                    If Not mbOnlyUnpackDP Then
                        ' установка драйверов
                        DPInstExitCode = RunDPInst(ArchTempPath)
                        ReadExitCodeString = ReadExitCode(DPInstExitCode)
    
                        If DPInstExitCode <> 0 Then
                            If DPInstExitCode <> -2147483648# Then
                                If InStr(1, ReadExitCodeString, "Cancel or Nothing to Install", vbTextCompare) = 0 Then
                                    ' Обрабатываем файл finish
                                    If mbLoadFinishFile Then
                                        WorkWithFinish strPathDRP, strPackFileName, ArchTempPath, strPathDRPList
                                    End If
                                    ' Обновление подсказки
                                    ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, Index, True
                                End If
                            End If
                        End If
                        ChangeStatusTextAndDebug strMessages(64) & strSpace & strPackFileName & " finish. " & ReadExitCodeString
                        If mbDebugStandart Then DebugMode "Install from : " & strPackFileName & " finish."
                    Else
                        ChangeStatusTextAndDebug strMessages(125) & strSpace & ArchTempPath
            
                        If MsgBox(strMessages(125) & str2vbNewLine & strMessages(133), vbYesNo, strProductName) = vbYes Then
                            RunUtilsShell ArchTempPath, False
                        End If
                    End If
                End If
            End If

            mbDevParserRun = False

            If mbGroupTask Then
                ' Удаление временных файлов
                ChangeStatusTextAndDebug strMessages(81), strMessages(130)
                strFileName_woExt = GetFileName_woExt(strPackFileName)
                ArchTempPath = strWorkTempBackSL & strFileName_woExt
                
                If PathExists(ArchTempPath) Then
                    DelRecursiveFolder ArchTempPath
                End If
            Else
                BlockControl True
            End If
            acmdPackFiles(Index).Value = False
        End If

        If Not mbGroupTask Then
            ' Проверка выделенных пакетов
            cmdRunTask.Enabled = FindCheckCount(False)
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub acmdPackFiles_KeyDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'                              KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub acmdPackFiles_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeySpace Then
        
        If chkPackFiles(Index).Value Then
            chkPackFiles(Index).Value = 0
        Else
            chkPackFiles(Index).Value = 1
        End If
        
        cmdRunTask.Enabled = FindCheckCount
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub acmdPackFiles_MouseDown
'! Description (Описание)  :   [Обработка события нажатия правой кнопкой мыши]
'! Parameters  (Переменные):   Index (Integer)
'                              Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub acmdPackFiles_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        mnuContextTxt.Enabled = True
        mnuContextXLS.Enabled = True
        mnuContextToolTip.Enabled = True
        mnuContextDeleteDevIDs.Enabled = True
        mnuContextCopyHWIDs.Enabled = True

        If acmdPackFiles(Index).PictureNormal = imgNo.Picture Then
            mnuContextToolTip.Enabled = False
            mnuContextDeleteDevIDs.Enabled = False
            mnuContextCopyHWIDs.Enabled = False
        ElseIf acmdPackFiles(Index).PictureNormal = imgNoDB.Picture Then
            mnuContextTxt.Enabled = False
            mnuContextXLS.Enabled = False
            mnuContextToolTip.Enabled = False
            mnuContextDeleteDevIDs.Enabled = False
            mnuContextCopyHWIDs.Enabled = False
        End If
        
        If mnuContextDeleteDevIDs.Enabled Then
            ' создаем меню для удаления драйверов устройств
            CreateMenuDevIDIndexDelMenu arrDevIDs(Index)
        End If
        If mnuContextCopyHWIDs.Enabled Then
            ' создаем меню для копирования HWID устройств
            CreateMenuDevIDIndexCopyMenu arrDevIDs(Index)
        End If

        lngCurrentBtnIndex = Index
    End If

End Sub

Private Sub acmdPackFiles_MouseEnter(Index As Integer)
        lngCurrentBtnIndex = Index
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkPackFiles_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub chkPackFiles_Click(Index As Integer)

    Dim lngCheckCount As Long

    chkPackFiles(Index).Value = chkPackFiles(Index).Value
    lngCheckCount = FindCheckCount
    cmdRunTask.Enabled = lngCheckCount

    If lngCheckCount Then
        ChangeStatusTextAndDebug strMessages(104) & strSpace & lngCheckCount, , False
    Else
        ChangeStatusTextAndDebug strMessages(105), , False
    End If

    chkPackFiles(Index).Refresh
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbCheckButton_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbCheckButton_Click()

    Dim strTextforCheck As String

    strTextforCheck = cmbCheckButton.Text

    If StrComp(strTextforCheck, strCmbChkBtnListElement2, vbTextCompare) = 0 Then
        LoadIconImage2Object cmdCheck, "BTN_UNCHECKMARK", strPathImageMainWork
    Else
        LoadIconImage2Object cmdCheck, "BTN_CHECKMARK", strPathImageMainWork
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdBreakUpdateDB_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdBreakUpdateDB_Click()
    mbBreakUpdateDBAll = True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdCheck_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdCheck_Click()

    Dim i                 As Long
    Dim strTextforCheck   As String
    Dim lngCntBtnTab      As Long
    Dim lngCntBtnPrevious As Long
    Dim lngCheckCount     As Long
    Dim lngSSTab1Tab      As Long

    If mbDebugStandart Then DebugMode "cmdCheck_Click-Start"
    CheckAllButton False
    strTextforCheck = cmbCheckButton.Text

    If strTextforCheck <> strCmbChkBtnListElement3 Then
        If strTextforCheck <> strCmbChkBtnListElement2 Then

            lngSSTab1Tab = SSTab1.Tab

            If lngSSTab1Tab Then
                lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 1).CntBtn

                If lngCntBtnPrevious = 0 Then
                    If lngSSTab1Tab > 1 Then
                        lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 2).CntBtn
                    End If
                End If
            End If

            lngCntBtnTab = arrOSList(lngSSTab1Tab).CntBtn - 1
        End If
    End If

    'Выбор режима выделения
    Select Case strTextforCheck

            '"все"
        Case strCmbChkBtnListElement3
            CheckAllButton True

            '"Все на текущей вкладке"
        Case strCmbChkBtnListElement1

            For i = lngCntBtnPrevious To lngCntBtnTab

                If lngCntBtnPrevious <> lngCntBtnTab Then

                    With acmdPackFiles(i)

                        If Not (.PictureNormal Is Nothing) Then
                            If .Visible Then
                                If .Left Then
                                    chkPackFiles(i).Value = True
                                End If
                            End If
                        End If

                    End With

                End If

            Next

            '"Все новые"
        Case strCmbChkBtnListElement4

            With acmdPackFiles
                For i = .LBound To .UBound
    
                    If Not (.item(i).PictureNormal Is Nothing) Then
                        If .item(i).PictureNormal = imgNoDB.Picture Then
                            If .item(i).Visible Then
                                chkPackFiles(i).Value = True
                            End If
                        End If
                    End If
    
                Next
            End With

            '"Неустановленные"
        Case strCmbChkBtnListElement5

            For i = lngCntBtnPrevious To lngCntBtnTab

                With acmdPackFiles(i)

                    If .Left Then
                        If Not (.PictureNormal Is Nothing) Then
                            If .PictureNormal = imgOkAttention.Picture Then
                                If .Visible Then
                                    chkPackFiles(i).Value = True
                                End If
                            End If
                        End If
                    End If

                End With

            Next

            '"Рекомендуемые"
        Case strCmbChkBtnListElement6

            For i = lngCntBtnPrevious To lngCntBtnTab

                With acmdPackFiles(i)

                    If Not (.PictureNormal Is Nothing) Then
                        If .Left Then
                            If .Visible Then
                                chkPackFiles(i).Value = True
                            End If

                            If .PictureNormal = imgNo.Picture Then
                                chkPackFiles(i).Value = False
                            End If

                            If .PictureNormal = imgNoDB.Picture Then
                                chkPackFiles(i).Value = False
                            End If

                            If .PictureNormal = imgOK.Picture Then
                                chkPackFiles(i).Value = False
                            End If

                            If mbCompareDrvVerByDate Then
                                If .PictureNormal = imgOkOld.Picture Then
                                    chkPackFiles(i).Value = False
                                End If
                            End If
                        End If
                    End If

                End With

            Next

            '"Сброс отметок"
        Case strCmbChkBtnListElement2
            CheckAllButton False

        Case Else
            cmbCheckButton.ListIndex = 0
    End Select

    lngCheckCount = FindCheckCount
    cmdRunTask.Enabled = lngCheckCount
    
    If lngCheckCount Then
        ChangeStatusTextAndDebug strMessages(104) & strSpace & lngCheckCount
    Else
        ChangeStatusTextAndDebug strMessages(105)
    End If

    If mbDebugStandart Then DebugMode "cmdCheck_Click-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdRunTask_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdRunTask_Click()
    mbSelectInstall = False
    mbGroupTask = True
    BlockControl False
    BaseUpdateOrRunTask False, True
    BlockControl True
    cmdRunTask.Enabled = FindCheckCount(False)
    mbGroupTask = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdViewAllDevice_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdViewAllDevice_Click()

    'MsgBox frRezim.Font.Name
    If IsFormLoaded("frmListHwidAll") = False Then
        frmListHwidAll.Show vbModal, Me
    Else
        frmListHwidAll.FormLoadDefaultParam
        frmListHwidAll.FormLoadAction
        frmListHwidAll.Show vbModal, Me
    End If

    If mbDeleteDriverByHwid Then
        If MsgBox(strMessages(113), vbQuestion + vbYesNo, strProductName) = vbYes Then
            mnuReCollectHWID_Click
        End If
    End If

    mbDeleteDriverByHwid = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Activate
'! Description (Описание)  :   [Событие активации формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()

    Dim lStart           As Long
    Dim lEnd             As Long
    Dim cntFindUnHideTab As Integer

    If mbFirstStart Then
        If mbStartMaximazed Then
            Me.WindowState = vbMaximized
            DoEvents
        ElseIf mbChangeResolution Then
            Me.WindowState = vbMaximized
            DoEvents
        End If

        ' Создаем элемент ProgressBar
        CreateProgressNew
        DoEvents

        ' поиск устройств при запуске
        With ctlProgressBar1
            .Value = 100
            .SetTaskBarProgressValue .Value, 1000
        End With

        ChangeFrmMainCaption 100

        If mbSearchOnStart Then
            RunDevconRescan lngPauseAfterSearch
        End If

        ChangeStatusTextAndDebug strMessages(3)

        ' сбор данных о компе
        If RunDevcon Then

            With ctlProgressBar1
                .Value = 150
                .SetTaskBarProgressValue 150, 1000
            End With

            ChangeFrmMainCaption 150
            DevParserLocalHwids2
            'Get the start time
            lStart = GetTickCount
            Me.Enabled = False
            'CollectHwid
            CollectHwidFromReestr
            Me.Enabled = True
            'Get the end time
            lEnd = GetTickCount
            If mbDebugStandart Then DebugMode "Time to Collect INFO from Reestr: =" & CalculateTime(lStart, lEnd, True)

            With ctlProgressBar1
                .Value = 250
                .SetTaskBarProgressValue 250, 1000
            End With

            ChangeFrmMainCaption 250
            ChangeStatusTextAndDebug strMessages(80)
            
            ' Назначить имена для вкладок и установить текущую на основании версии ОС
            SetTabsNameAndCurrTab False
            ' Загрузить все кнопки
            LoadButton
            'выставить настройки шрифта
            SetTTFontProperties TT
            'сохранить найденные драйвера в файл
            SaveHWIDs2File
    
            ' Вывести в лог список всех драйверов
            If lngArrDriversIndex Then
                PutAllDrivers2Log
            End If
    
            SetTabsNameAndCurrTab True
            DoEvents
            BlockControl True
            ' Активация меню языков и локализации приложения
            mnuMainLang.Enabled = mbMultiLanguage
    
            FindCheckCount
            frTabPanel.Visible = True
    
            If SSTab1.Tab = 0 Then
                If Not SSTab1.TabEnabled(0) Then
                    If acmdPackFiles.Count <= 1 Then
                        acmdPackFiles(0).Visible = False
                        chkPackFiles(0).Visible = False
                    End If
                End If
            End If
    
            mbFirstStart = False
            ' Режим при старте
            SelectStartMode
            ' Активируем скрытую форму
            frTabPanel.Visible = True
            SSTab1.Enabled = True
    
            ' устанавливаем размера табконтрола и положения FrameScroll
            With frTabPanel
                cntFindUnHideTab = FindUnHideTab
    
                If .Visible Then
                    SetTabsWidth cntFindUnHideTab
                    SetStartScrollFramePos cntFindUnHideTab
                End If
    
            End With
    
            ' подсчитываем кол-во неизвестных драйверов и изменяем текст кнопки
            LoadCmdViewAllDeviceCaption
            
            ' Загружаем описание значков иконок
            ToolTipStatusLoad
            Unload frmLicence
            Set frmLicence = Nothing
            dtEndTimeProg = GetTickCount
            dtAllTimeProg = CalculateTime(dtStartTimeProg, dtEndTimeProg)
            
            ChangeStatusTextAndDebug strMessages(59) & strSpace & dtAllTimeProg
            If mbDebugStandart Then DebugMode "End Start Operation" & " StartTime is: " & dtAllTimeProg
    
            If mbRunWithParam Then
                ChangeStatusTextAndDebug strMessages(60)
                If mbDebugStandart Then DebugMode "Program start in silentMode"
                frmSilent.Show vbModal, Me
    
                If mbSilentRun Then
    
                    If Not mbNoSupportedOS Then
                        '"Начинается автоматическая установка"
                        SilentInstall
                        ' после установки закрываем программу
                        Unload Me
                    End If
    
                End If
    
            Else
                ' Разные сообщения если нет поддерживаемых вкладок, или что-то нет так с пакетами
                EventOnActivateForm
    
                ' Проверка обновлений при старте, только если не тихий режим установки
                If mbUpdateCheck Then
                    ctlUcStatusBar1.PanelText(1) = strMessages(145)
                    ChangeStatusTextAndDebug strMessages(58)
                    CheckUpd
                Else
                    ShowUpdateToolTip
                End If
            End If
        Else
            MsgBox strHwidsTxtPath & vbNewLine & strMessages(46), vbInformation, strProductName
            Unload Me
        End If

    End If

    mbFirstStart = False
    mbLoadAppEnd = True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_KeyDown
'! Description (Описание)  :   [обработка нажатий клавиш клавиатуры]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Нажата кнопка "Ctrl"
    If Shift = 2 Then

        Select Case KeyCode

            Case vbKeyA '65
                ' Ctrl+A (Выделение всех пакетов для установки)
                CheckAllButton True

            Case vbKeyZ '90
                ' Ctrl+Z (Отмена выделения всех)
                CheckAllButton False

            Case vbKeyS '83
                ' Ctrl+S (Выделение всех пакетов на вкладке)
                SelectAllOnTabDP True

            Case vbKeyN '78
                ' Ctrl+N (Выделение всех пакетов с новыми драйверами)
                SelectRecommendedDP True

            Case vbKeyQ '81
                ' Ctrl+Q (Выделение пакетов с не установленными)
                SelectNotInstalledDP True

            Case vbKeyI '73
                ' Ctrl+I (Установка выделенных пакетов)
                InsOrUpdSelectedDP True

            Case vbKeyU '85
                ' Ctrl+U (Обновление БД выделенных пакетов)
                InsOrUpdSelectedDP False

            Case vbKeyTab
                ' CTRL+Tab (Переключение по вкладкам SSTab1)
                If SSTab1.Tabs Then
                    SelectNextTab
                End If

            Case 19
                ' CTRL+Break (Прерывание групповой обработки)
                If cmdBreakUpdateDB.Visible Then
                    mbBreakUpdateDBAll = True
                End If
            
            Case vbKeyX '88
                ' Ctrl+X (Прерывание групповой обработки в IDE)
                If cmdBreakUpdateDB.Visible Then
                    mbBreakUpdateDBAll = True
                End If

        End Select

    ElseIf Shift = 0 Then
        ' Выход из программы по "Escape"
        If Not mbFirstStart Then
            If KeyCode = vbKeyEscape Then
                If Not mbCheckUpdNotEnd Then
                    If VBA.MsgBox(strMessages(34), vbQuestion + vbYesNo, strProductName) = vbYes Then
                        Unload Me
                    End If
                End If
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Load
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()

    Dim i  As Long
    Dim ii As Long

    If mbDebugStandart Then DebugMode "MainForm Show"
    SetupVisualStyles Me

    With Me
        ' изменяем иконки формы и приложения
        ' Icon for Exe-file
        SetIcon .hWnd, "APPICOTAB", True
        SetIcon .hWnd, "FRMMAIN", False
        DoEvents
        ' Смена заголовка формы
        strFormName = .Name
        ChangeFrmMainCaption
        ' Изменяем размеры формы
        .Width = lngMainFormWidth
        .Height = lngMainFormHeight
        ' Центрируем форму на экране
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With
    
    ' Инициализируем переменные для поиска совместимых HWID
    Set objHashOutput = New Scripting.Dictionary
    objHashOutput.CompareMode = TextCompare
    Set objHashOutput2 = New Scripting.Dictionary
    objHashOutput2.CompareMode = TextCompare
    Set objHashOutput3 = New Scripting.Dictionary
    objHashOutput3.CompareMode = TextCompare
    
    LoadIconImage
    ' Подчеркавание меню (аля 3D)
    Me.Line (0, 15)-(ScaleWidth, 15), vbWhite
    Me.Line (0, 0)-(ScaleWidth, 0), GetSysColor(COLOR_BTNSHADOW)
    
    frRezim.Top = 500
    frRunChecked.Top = 500
    frDescriptionIco.Top = 2100
    frTabPanel.Top = 3150

    ' Начальные параметры статус панели
    With ctlUcStatusBar1
        .AddPanel strMessages(127)
        .AddPanel strMessages(1), , False
        .PanelWidth(2) = (lngMainFormWidth \ Screen.TwipsPerPixelX) - .PanelWidth(1)
    End With

    ' Запись файла настроек в ini
    PrintFileInDebugLog strSysIni
    ' Инициализация клипбоарда
    InitClipboard

    If lngOSCount < lngOSCountPerRow Then
        SSTab1.TabsPerRow = lngOSCount
    Else
        SSTab1.TabsPerRow = lngOSCountPerRow
    End If

    ' информация о системе свернута
    frInfo.Collapsado = False
    frTabPanel.Visible = False
    
    'Устанавливаем неактичность элементов управления
    BlockControl False
    
    ' Проверка доступности стандартных дополнительных утилит
    CheckMenuUtilsPath
    
    ' Начальные позиции некоторых элементов управления
    frTabPanel.Top = 3100
    frTabPanel.Left = 75
    lblOSInfo.Left = 75

    With acmdPackFiles(0)
        .ButtonStyle = lngStatusBtnStyle
        .ColorScheme = lngStatusBtnStyleColor
        If lngStatusBtnStyleColor = 3 Then
            .BackColor = lngStatusBtnBackColor
        End If
        .Left = lngButtonLeft
        .Top = lngButtonTop
        .Width = lngButtonWidth
        .Height = lngButtonHeight
        .CheckExist = True
        .ForeColor = lngFontBtn_Color
        ' Устанавливаем шрифт кнопок
        SetBtnStatusFontProperties acmdPackFiles(0)
    End With

    With chkPackFiles(0)
        .Width = 200
        .Height = 200
        .Left = lngButtonLeft + lngButtonWidth - 225
        .Top = lngButtonTop + 30
    End With

    ' Устанавливаем шрифт закладок
    SetTabProperties
    SetTabPropertiesTabDrivers
    ' Свойства скороллформы
    ctlScrollControl1(0).BorderStyle = vbBSNone
    ctlScrollControlTab1(0).BorderStyle = vbBSNone
    ctlScrollControlTab2(0).BorderStyle = vbBSNone
    ctlScrollControlTab3(0).BorderStyle = vbBSNone
    ctlScrollControlTab4(0).BorderStyle = vbBSNone

    If lngOSCount <> 9999 Then
        If lngOSCount <> 0 Then
            SSTab1.Tabs = lngOSCount
        End If
    End If

    If mbDebugStandart Then DebugMode "LoadTabList" & vbNewLine & _
              "TabsPerRow: " & SSTab1.TabsPerRow & vbNewLine & _
              "TabsCount: " & SSTab1.Tabs

    ' Загрузка меню утилит
    If arrUtilsList(0, 1) <> "List_Empty" Then
        If mbDebugStandart Then DebugMode "CreateUtilsList: " & UBound(arrUtilsList)

        For i = UBound(arrUtilsList) To 0 Step -1
            CreateMenuIndex arrUtilsList(i, 0)
        Next

    End If

    ' Загрузка меню языков и локализация приложения
    If mbMultiLanguage Then
        If mbDebugStandart Then DebugMode "CreateLangList: " & UBound(arrLanguage) + 1

        ' Создаем меню поддержки языков
        CreateMenuLng
        
        ' Локализация приложения
        Localise strPCLangCurrentPath
        
        ' Устанавливаем галочку на активном языке
        For ii = mnuLang.LBound To mnuLang.UBound
            mnuLang(ii).Checked = arrLanguage(1, ii + 1) = strPCLangCurrentPath
        Next
        
        ' Устанавливаем галочку на автовыборе языка
        mnuLangStart.Checked = Not mbAutoLanguage
    End If

    If mbDebugStandart Then DebugMode "OsInfo: " & lblOSInfo.Caption & vbNewLine & _
              "PCModel: " & lblPCInfo.Caption
    ' Выставляем шрифт
    FontCharsetChange

    ' Изменяем параметры Всплывающей подсказки для кнопок
    With TT
        .MaxTipWidth = lngRightWorkArea
        .SetDelayTime TipDelayTimeInitial, 400
        .SetDelayTime TipDelayTimeShow, 15000
        .Title = strTTipTextTitle
        'SetTTFontProperties TT
    End With

    ' Изменяем параметры кнопок и картинок
    imgOK.BorderStyle = 0
    imgOkAttention.BorderStyle = 0
    imgOkNew.BorderStyle = 0
    imgOkOld.BorderStyle = 0
    imgOkAttentionNew.BorderStyle = 0
    imgOkAttentionOLD.BorderStyle = 0
    imgNo.BorderStyle = 0
    imgNoDB.BorderStyle = 0
    imgUpdBD.BorderStyle = 0
    'загрузка меню кнопки CmdRunTask
    LoadCmdRunTask
    'заполнение списка на выделение
    LoadListChecked
    mbFirstStart = True

    If mbIsWin64 Then
        If PathExists(PathCollect("Tools\SIV\SIV64X.exe")) Then
            lblOSInfo.ToolTipText = "View system info using System Information Viewer"
        End If

    Else

        If PathExists(PathCollect("Tools\SIV\SIV32X.exe")) Then
            lblOSInfo.ToolTipText = "View system info using System Information Viewer"
        End If
    End If

    mnuAutoInfoAfterDelDRV.Checked = mbAutoInfoAfterDelDRV
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_QueryUnload
'! Description (Описание)  :   [Корректная выгрузка формы]
'! Parameters  (Переменные):   Cancel (Integer)
'                              UnloadMode (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' Проверяем закончена ли проверка обновления, если нет то прерываем выход из программы, иначе программа вылетит
    If mbCheckUpdNotEnd Then
        Cancel = UnloadMode = vbFormControlMenu Or vbFormCode
        Exit Sub
    End If

    ' Удаление временных файлов если есть и если опция включена
    If mbDelTmpAfterClose Then
        ChangeStatusTextAndDebug strMessages(81), strMessages(130)

        'Чистим если только не перезапуск программы
        If Not mbRestartProgram Then
            'Me.Hide
            DelTemp
        End If
    End If

    
    Dim i As Long
    For i = acmdPackFiles.LBound To acmdPackFiles.UBound
        acmdPackFiles(i).UnsetPopupMenu
        acmdPackFiles(i).UnsetPopupMenuRBT
    Next i
    
    ' сохранение параметров при выходе
    If mbSaveSizeOnExit Then
        FRMStateSave
    End If

    ' Сохраняем язык при старте
    If Not mbIsDriveCDRoom Then
        If mnuLangStart.Checked Then
            IniWriteStrPrivate "Main", "StartLanguageID", strPCLangCurrentID, strSysIni
        End If

        IniWriteStrPrivate "Main", "AutoLanguage", CStr(Abs(Not mnuLangStart.Checked)), strSysIni
        IniWriteStrPrivate "Main", "AutoInfoAfterDelDRV", CStr(Abs(mnuAutoInfoAfterDelDRV.Checked)), strSysIni
    End If

    SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP", False

    If mbLoadIniTmpAfterRestart Then
        SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP_PATH", "-"

        If StrComp(GetFileNameFromPath(strSysIni), "Settings_DIA_TMP.ini", vbTextCompare) = 0 Then
            DeleteFiles strSysIni
        End If
    End If
    
    If lngFrameTime < 0 Then lngFrameTime = 2
    If lngFrameCount < 1 Then lngFrameCount = 40
    If Me.WindowState <> vbMinimized Then
        AnimateForm Me, aUnload, eZoomOut, lngFrameTime, lngFrameCount
    End If

    Set objHashOutput = Nothing
    Set objHashOutput2 = Nothing
    Set objHashOutput3 = Nothing
    
    Set frmMain = Nothing
        
    ' Выгружаем из памяти формы
    UnloadAllForms strFormName
    
    ' Выгружаем из памяти главную форму
    Unload Me
    Set frmMain = Nothing
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Resize
'! Description (Описание)  :   [Изменение размеров контролов при изменении размеров формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Form_Resize()

    Dim OptWidth             As Long
    Dim OptWidthDelta        As Long
    Dim ImgWidth             As Long
    Dim imgWidthDelta        As Long
    Dim miDeltafrmMainWidth  As Long
    Dim miDeltafrmMainHeight As Long
    Dim cntFindUnHideTab     As Integer

    On Error Resume Next

    ' если форма не свернута, то изменяем размеры
    If Me.WindowState <> vbMinimized Then

        ' если форма не максимизирована, то изменяем размеры формы
        If OSCurrVersionStruct.VerFull >= "6.0" Then
            miDeltafrmMainWidth = 120
            miDeltafrmMainHeight = 120
            '            miDeltafrmMainHeight = 405
            '            If mbAeroEnabled Then
            '                miDeltafrmMainWidth = 216
            '                miDeltafrmMainHeight = 540
            '            End If
        Else

            If mbAppThemed Then
                miDeltafrmMainWidth = 0
            Else
                miDeltafrmMainWidth = 0
            End If
        End If

        If Me.Width < lngMainFormWidthMin Then
            Me.Width = lngMainFormWidthMin
            Me.Enabled = False
            Me.Enabled = True

            Exit Sub

        End If

        If Me.Height < lngMainFormHeightMin Then
            Me.Height = lngMainFormHeightMin
            Me.Enabled = False
            Me.Enabled = True

            Exit Sub

        End If

        With Me
            frMainPanel.Top = 0
            frMainPanel.Left = 0
            frMainPanel.Height = (.Height - 2.1 * ctlUcStatusBar1.Height - miDeltafrmMainHeight)
            frMainPanel.Width = (.Width)
        End With

        If Not (ctlUcStatusBar1 Is Nothing) Then
            If ctlUcStatusBar1.PanelCount > 1 Then
                ctlUcStatusBar1.PanelWidth(2) = (Me.Width \ Screen.TwipsPerPixelX) - ctlUcStatusBar1.PanelWidth(1)
                ctlUcStatusBar1.Refresh
            End If
        End If

        pbProgressBar.Align = 0
        pbProgressBar.Width = Me.Width
        pbProgressBar.Align = 2
        pbProgressBar.Refresh
        ctlProgressBar1.Width = pbProgressBar.Width
        ctlProgressBar1.Refresh
        frRunChecked.Left = frMainPanel.Width - frRunChecked.Width - 150 - miDeltafrmMainWidth
        frRezim.Width = frRunChecked.Left - frRezim.Left - 50
        frInfo.Left = frRezim.Left
        frInfo.Width = frRezim.Width + frRunChecked.Width + 50

        ' устанавливаем размера табконтрола и положения FrameScroll
        With frTabPanel
            .Height = frMainPanel.Height - .Top - 50
            .Width = frRunChecked.Left + frRunChecked.Width - 50

            SSTab1.Height = frTabPanel.Height - 20
            SSTab1.Width = frTabPanel.Width - 20

            ' Изменение размеров FrameScroll и ширины вкладок sstab
            cntFindUnHideTab = FindUnHideTab

            If .Visible Then
                SetTabsWidth cntFindUnHideTab
                SetStartScrollFramePos cntFindUnHideTab
            End If

        End With

        ' устанавливаем ширину кнопок выбора режима
        OptWidth = frRezim.Width / 3 - 125
        OptWidthDelta = OptWidth + 100
        optRezim_Intellect.Width = OptWidth
        optRezim_Intellect.Left = 100
        optRezim_Ust.Width = OptWidth
        optRezim_Ust.Left = optRezim_Intellect.Left + OptWidthDelta
        optRezim_Upd.Width = OptWidth
        optRezim_Upd.Left = optRezim_Ust.Left + OptWidthDelta
        ' устанавливаем ширину иконок и описаний статусов кнопок
        ImgWidth = imgOK.Width
        frDescriptionIco.Width = frRezim.Width
        imgWidthDelta = ((frDescriptionIco.Width - imgOK.Width * 9) / 9)
        imgOK.Left = (frDescriptionIco.Width - imgOK.Width * 9 - imgWidthDelta * 8) / 2
        imgOkAttention.Left = imgOK.Left + ImgWidth + imgWidthDelta
        imgOkNew.Left = imgOkAttention.Left + ImgWidth + imgWidthDelta
        imgOkOld.Left = imgOkNew.Left + ImgWidth + imgWidthDelta
        imgOkAttentionNew.Left = imgOkOld.Left + ImgWidth + imgWidthDelta
        imgOkAttentionOLD.Left = imgOkAttentionNew.Left + ImgWidth + imgWidthDelta
        imgNo.Left = imgOkAttentionOLD.Left + ImgWidth + imgWidthDelta
        imgNoDB.Left = imgNo.Left + ImgWidth + imgWidthDelta
        imgUpdBD.Left = imgNoDB.Left + ImgWidth + imgWidthDelta
        lblOSInfo.Width = frInfo.Width - 200
        lblPCInfo.Width = frInfo.Width - 200
        cmdViewAllDevice.Width = optRezim_Upd.Left + optRezim_Upd.Width - cmdViewAllDevice.Left
        ' Удаление иконки в трее если есть
        SetTrayIcon NIM_DELETE, Me.hWnd, 0&, vbNullString

        With lblNoDPInProgram
            .Left = 100

            ' Изменяем положение лабел
            Dim cntUnHideTab   As Long
            Dim miValue1       As Long
            Dim sngNum1        As Single
            Dim SSTabTabHeight As Long

            SSTabTabHeight = SSTab1.TabHeight
            cntUnHideTab = FindUnHideTab

            If cntUnHideTab Then
                sngNum1 = (cntUnHideTab + 1) / lngOSCountPerRow
                miValue1 = Round(sngNum1, 0)
            Else
                miValue1 = 1
            End If

            If sngNum1 = miValue1 Then
                .Width = SSTab1.Width - 150 * (sngNum1 + 1)
                .Top = (SSTab1.Height - .Height + (SSTabTabHeight * (miValue1))) / 2
            Else
                .Width = SSTab1.Width - 150 * (sngNum1 + 1)
                .Top = (SSTab1.Height - .Height + (SSTabTabHeight * (miValue1))) / 2
            End If

            .AutoSize = False
        End With

        With lblNoDP4Mode
            .Left = 100
            .Width = ctlScrollControl1(0).Width - 200
            .Top = (ctlScrollControl1(0).Height - .Height) / 2
            .ZOrder 0
        End With

        If Not mbFirstStart Then
            StartReOrderBtnOnTab2 SSTab1.Tab, 1
        End If

    Else
        ' Добавляеи иконку в трей
        SetTrayIcon NIM_ADD, Me.hWnd, Me.Icon, "Drivers Installation Assistant"
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub frDescriptionIco_MouseDown
'! Description (Описание)  :   [Контектсное меню для формы со списком обозначений кнопок]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub frDescriptionIco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, PanelArea As m_PanelArea)
    If Button = vbRightButton Then
        OpenContextMenu Me, Me.mnuContextMenu2
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lblOsInfo_MouseDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub lblOsInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mnuUtils_SIV.Enabled Then mnuUtils_SIV_Click
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuAbout_Click
'! Description (Описание)  :   [Меню - О программе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuAutoInfoAfterDelDRV_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuAutoInfoAfterDelDRV_Click()
    mnuAutoInfoAfterDelDRV.Checked = Not mnuAutoInfoAfterDelDRV.Checked
    mbAutoInfoAfterDelDRV = Not mbAutoInfoAfterDelDRV
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuCheckUpd_Click
'! Description (Описание)  :   [еню - Проверить обновление]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuCheckUpd_Click()
    ctlUcStatusBar1.PanelText(1) = strMessages(145)
    ChangeStatusTextAndDebug strMessages(58)
    CheckUpd False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextCopyHWID2Clipboard_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuContextCopyHWID2Clipboard_Click(Index As Integer)

    Dim strValue          As String
    Dim strValueDevID     As String
    Dim strValueDevID_x() As String

    strValue = mnuContextDeleteDevID(Index).Caption
    strValueDevID = Left$(strValue, InStr(strValue, vbTab) - 1)

    If InStr(strValueDevID, vbBackslash) Then
        strValueDevID_x = Split(strValueDevID, vbBackslash)
        strValueDevID = strValueDevID_x(0) & vbBackslash & strValueDevID_x(1)
    End If

    ' Копируем текст в клипбоард
    CBSetText strValueDevID
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextDeleteDevID_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuContextDeleteDevID_Click(Index As Integer)

    Dim strValue                 As String
    Dim strValueDevID            As String
    Dim strValueDevID_x()        As String
    Dim mbDeleteDriverByHwidTemp As Boolean

    strValue = mnuContextDeleteDevID(Index).Caption
    strValueDevID = Left$(strValue, InStr(strValue, vbTab) - 1)

    If InStr(strValueDevID, vbBackslash) Then
        strValueDevID_x = Split(strValueDevID, vbBackslash)
        strValueDevID = strValueDevID_x(0) & vbBackslash & strValueDevID_x(1)
    End If

    mbDeleteDriverByHwidTemp = DeleteDriverbyHwid(strValueDevID)

    If mbDeleteDriverByHwidTemp Then
        If Not mbDeleteDriverByHwid Then
            mbDeleteDriverByHwid = True
        End If
    End If

    If mbDeleteDriverByHwid Then
        If Not mbAutoInfoAfterDelDRV Then
            If MsgBox(strMessages(113), vbQuestion + vbYesNo, strProductName) = vbYes Then
                mnuReCollectHWID_Click
            End If

        Else
            mnuReCollectHWID_Click
        End If
    End If

    mbDeleteDriverByHwid = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextDeleteDRP_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuContextDeleteDRP_Click()

    Dim i                 As Long
    Dim strPathDRP        As String
    Dim strPathDB         As String
    Dim strFullPathDRP    As String
    Dim strFullPathDB     As String
    Dim strFullPathDBHwid As String
    Dim strFullPathDBIni  As String

    If mbIsDriveCDRoom Then
        MsgBox strMessages(16), vbInformation, strProductName
    Else
        i = SSTab1.Tab
        strPathDRP = arrOSList(i).drpFolderFull
        strPathDB = arrOSList(i).devIDFolderFull
        strFullPathDRP = PathCombine(strPathDRP, acmdPackFiles(lngCurrentBtnIndex).Tag)
        strFullPathDB = PathCombine(strPathDB, GetFileNameFromPath(strCurSelButtonPath))
        strFullPathDBIni = Replace$(strFullPathDB, ".txt", ".ini", , , vbTextCompare)
        strFullPathDBHwid = Replace$(strFullPathDB, ".txt", ".hwid", , , vbTextCompare)

        If MsgBox(strMessages(17) & " '" & acmdPackFiles(lngCurrentBtnIndex).Tag & "' ?", vbQuestion + vbYesNo, strProductName) = vbYes Then
            ' удаление пакета 7z
            If PathExists(strFullPathDRP) Then
                If Not PathIsAFolder(strFullPathDRP) Then
                    If mbDebugStandart Then DebugMode "Delete file: " & strFullPathDRP
                    DeleteFiles strFullPathDRP
                End If
                
                ' удаление файла txt
                If PathExists(strFullPathDB) Then
                    If Not PathIsAFolder(strFullPathDB) Then
                        If mbDebugStandart Then DebugMode "Delete file: " & strFullPathDB
                        DeleteFiles strFullPathDB
                        'Удаление секции о данном пакете из ini-файла
                        IniDelAllKeyPrivate GetFileName_woExt(GetFileNameFromPath(strCurSelButtonPath)), PathCombine(strPathDB, "DevDBVersions.ini")
                    End If
                End If
                
                ' удаление файла hwid
                If PathExists(strFullPathDBHwid) Then
                    If Not PathIsAFolder(strFullPathDBHwid) Then
                        If mbDebugStandart Then DebugMode "Delete file: " & strFullPathDBHwid
                        DeleteFiles strFullPathDBHwid
                    End If
                End If
                
                ' удаление файла ini
                If PathExists(strFullPathDBIni) Then
                    If Not PathIsAFolder(strFullPathDBIni) Then
                        If mbDebugStandart Then DebugMode "Delete file: " & strFullPathDBIni
                        DeleteFiles strFullPathDBIni
                    End If
                End If
                
            End If
            
            acmdPackFiles(lngCurrentBtnIndex).Visible = False
            chkPackFiles(lngCurrentBtnIndex).Visible = False
            chkPackFiles(lngCurrentBtnIndex).Value = False
            ChangeStatusTextAndDebug strMessages(88) & strSpace & strFullPathDRP
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextEditDPName_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuContextEditDPName_Click()

    If Not FileisReadOnly(strSysIni) Then
        EditOrReadDPName lngCurrentBtnIndex
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextInstall_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuContextInstallGroupDP_Click(Index As Integer)
    mbGroupTask = True
    mbOnlyUnpackDP = False

    Select Case Index

        Case 0
            mbSelectInstall = False
            mbOnlyUnpackDP = False

        Case 2
            mbSelectInstall = True
            mbOnlyUnpackDP = False

        Case 4
            mbSelectInstall = False
            mbOnlyUnpackDP = True

        Case 5
            mbSelectInstall = True
            mbOnlyUnpackDP = True
    End Select

    GroupInstallDP
    mbGroupTask = False
    BlockControl True
    cmdRunTask.Enabled = FindCheckCount(False)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextInstall_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuContextInstallSingleDP_Click(Index As Integer)
    mbGroupTask = False
    mbOnlyUnpackDP = False

    Select Case Index

        Case 0
            mbSelectInstall = False
            mbOnlyUnpackDP = False

        Case 2
            mbSelectInstall = True
            mbOnlyUnpackDP = False

        Case 4
            mbSelectInstall = False
            mbOnlyUnpackDP = True

        Case 5
            mbSelectInstall = True
            mbOnlyUnpackDP = True
    End Select

    acmdPackFiles_Click CInt(lngCurrentBtnIndex)
    
    mbGroupTask = False
    BlockControl True
    cmdRunTask.Enabled = FindCheckCount(False)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextLegendIco_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuContextLegendIco_Click()
    frmLegendIco.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextTestDRP_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuContextTestDRP_Click()

    Dim cmdString       As String
    Dim strPackFileName As String
    Dim strPathDRP      As String

    strPackFileName = acmdPackFiles(lngCurrentBtnIndex).Tag
    strPathDRP = arrOSList(SSTab1.Tab).drpFolderFull
    cmdString = strKavichki & strArh7zExePATH & strKavichki & " t " & strKavichki & strPathDRP & strPackFileName & strKavichki & " -r"
    ChangeStatusTextAndDebug strMessages(109) & strSpace & strPackFileName
    BlockControl False

    If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
        MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
    Else

        ' Архиватор отработал на все 100%? Если нет то сообщаем
        If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
            ChangeStatusTextAndDebug strMessages(13) & strSpace & strPackFileName
            MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
        Else
            ChangeStatusTextAndDebug strMessages(110) & strSpace & strPackFileName
            MsgBox strMessages(110) & strSpace & strPackFileName, vbInformation, strProductName
        End If
    End If

    BlockControl True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextToolTip_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuContextToolTip_Click()
    mbSelectInstall = False

    If IsFormLoaded("frmListHwid") = False Then
        frmListHwid.Show vbModal, Me
    Else
        frmListHwid.FormLoadDefaultParam
        frmListHwid.FormLoadAction
        frmListHwid.Show vbModal, Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextTxt_Click
'! Description (Описание)  :   [Меню - Файл БД в текстовом виде]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuContextTxt_Click()
    RunUtilsShell strCurSelButtonPath, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextUpdStatus_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuContextUpdStatus_Click()

    Dim strPackFileName As String
    Dim strPathDRP      As String
    Dim strPathDevDB    As String

    strPathDRP = arrOSList(SSTab1.Tab).drpFolderFull
    strPathDevDB = arrOSList(SSTab1.Tab).devIDFolderFull
    strPackFileName = acmdPackFiles(lngCurrentBtnIndex).Tag
    ' Обновление подсказки
    ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, CInt(lngCurrentBtnIndex), , True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuContextXLS_Click
'! Description (Описание)  :   [Меню - Файл БД в Excel]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuContextXLS_Click()

    Dim strCurSelButtonPathTemp As String

    strCurSelButtonPathTemp = strWorkTempBackSL & GetFileNameFromPath(strCurSelButtonPath)
    ' Копируем файл БД во временный каталог
    CopyFileTo strCurSelButtonPath, strCurSelButtonPathTemp
    ' Открываем в Excel
    OpenTxtInExcel strCurSelButtonPathTemp
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuCreateBackUp_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuCreateBackUp_Click()

    Dim lngMsgRet As Long

    lngMsgRet = MsgBox(strMessages(123), vbYesNo + vbQuestion, strProductName)

    Select Case lngMsgRet

        Case vbYes
            mnuHomePage1_Click
    End Select

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuCreateRestorePoint_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuCreateRestorePoint_Click()

    If MsgBox(strMessages(115), vbQuestion + vbYesNo, strProductName) = vbYes Then
        ' Блокируем форму при создании точки восстановления
        BlockControl False
        ' Собственно создание точки
        CreateRestorePoint
        ' РазБлокируем форму после создания точки восстановления
        BlockControl True
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuDelDuplicateOldDP_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuDelDuplicateOldDP_Click()
    DelDuplicateOldDP

    If mbRestartProgram Then
        ShellExecute Me.hWnd, "open", strAppEXEName, vbNullString, strAppPath, SW_SHOWNORMAL
        Unload Me

        End

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuDonate_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuDonate_Click()
    frmDonate.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuDriverPacks_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuDriverPacks_Click()
    RunUtilsShell "http://driverpacks.net/driverpacks", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuDriverPacksOnMySite_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuDriverPacksOnMySite_Click()
    RunUtilsShell "http://adia-project.net/forum/index.php?topic=789.0", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuHelp_Click
'! Description (Описание)  :   [Меню - Помощь]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuHelp_Click()

    Dim strFilePathTemp As String

    strFilePathTemp = strAppPathBackSL & strToolsDocs_Path & "\" & strPCLangCurrentID & "\Help.html"

    If PathExists(strFilePathTemp) = False Then
        strFilePathTemp = strAppPathBackSL & strToolsDocs_Path & "\0409\Help.html"
    End If

    RunUtilsShell strFilePathTemp, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuHistory_Click
'! Description (Описание)  :   [Меню - История изменений]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuHistory_Click()

    Dim strFilePathTemp As String

    strFilePathTemp = strAppPathBackSL & strToolsDocs_Path & "\" & strPCLangCurrentID & "\history.txt"

    If PathExists(strFilePathTemp) = False Then
        strFilePathTemp = strAppPathBackSL & strToolsDocs_Path & "\0409\history.txt"
    End If

    RunUtilsShell strFilePathTemp, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuHomePage1_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuHomePage1_Click()
    RunUtilsShell strUrl_MainWWWSite, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuHomePage_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuHomePage_Click()
    RunUtilsShell "http://forum.oszone.net/thread-139908.html", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuLang_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuLang_Click(Index As Integer)

    Dim i                      As Long
    Dim ii                     As Long
    Dim strPathLng             As String
    Dim strPCLangCurrentIDTemp As String
    Dim strPCLangCurrentID_x() As String

    i = Index + 1

    For ii = mnuLang.LBound To mnuLang.UBound
        mnuLang(ii).Checked = ii = Index
    Next

    strPathLng = arrLanguage(1, i)
    strPCLangCurrentPath = strPathLng
    strPCLangCurrentIDTemp = arrLanguage(3, i)
    strPCLangCurrentLangName = arrLanguage(2, i)
    lngFont_Charset = GetCharsetFromLng(CLng(arrLanguage(6, i)))

    If InStr(strPCLangCurrentIDTemp, ";") Then
        strPCLangCurrentID_x = Split(strPCLangCurrentIDTemp, ";")
        strPCLangCurrentID = strPCLangCurrentID_x(0)
    Else
        strPCLangCurrentID = strPCLangCurrentIDTemp
    End If
    
    

    ' Собственно локализация
    Localise strPCLangCurrentPath

    ' ПереВыставляем шрифт основной формы
    With Me.Font
        .Name = strFontMainForm_Name
        .Size = lngFontMainForm_Size
        .Charset = lngFont_Charset
    End With
    
    ChangeFrmMainCaption

    ChangeStatusTextAndDebug strMessages(142) & strSpace & arrLanguage(2, i), , False

    If mbNoSupportedOS Then
        SelectStartMode 3, False
        BlockControl True
        BlockControlEx False
    End If
    
    cmdRunTask.Enabled = FindCheckCount(False)

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuLangStart_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuLangStart_Click()
    mnuLangStart.Checked = Not mnuLangStart.Checked
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuLicence_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuLicence_Click()
    frmLicence.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuLinks_Click
'! Description (Описание)  :   [Меню - Ссылки]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuLinks_Click()

    Dim strFilePathTemp As String

    strFilePathTemp = strAppPathBackSL & strToolsDocs_Path & "\" & strPCLangCurrentID & "\Links.html"

    If PathExists(strFilePathTemp) = False Then
        strFilePathTemp = strAppPathBackSL & strToolsDocs_Path & "\0409\Links.html"
    End If

    RunUtilsShell strFilePathTemp, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuLoadOtherPC_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuLoadOtherPC_Click()
    frmEmulate.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuModulesVersion_Click
'! Description (Описание)  :   [Меню - Версии модулей]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuModulesVersion_Click()
    VerModules
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuOptions_Click
'! Description (Описание)  :   [Меню - Настройки]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuOptions_Click()
    ctlUcStatusBar1.PanelText(1) = strMessages(146)
    ChangeStatusTextAndDebug strMessages(146)

    If IsFormLoaded("frmOptions") = False Then
        frmOptions.Show vbModal, Me
    Else
        frmOptions.FormLoadAction
        frmOptions.Show vbModal, Me
    End If

    If mbRestartProgram Then
        ShellExecute Me.hWnd, "open", strAppEXEName, vbNullString, strAppPath, SW_SHOWNORMAL
        Unload Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuReCollectHWID_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuReCollectHWID_Click()
    ' Сначала повторно собираем данные
    ReCollectHWID
    ' А теперь Обновляем статус всех пакетов
    UpdateStatusButtonAll
    SaveHWIDs2File
    ' Обновить список неизвестных дров и описание для кнопки
    LoadCmdViewAllDeviceCaption
    ChangeStatusTextAndDebug strMessages(114)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuReCollectHWIDTab_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuReCollectHWIDTab_Click()
    ' Сначала повторно собираем данные
    ReCollectHWID
    ' А теперь Обновляем статус всех пакетов для текущей вкладки
    UpdateStatusButtonTAB
    SaveHWIDs2File
    ' Обновить список неизвестных дров и описание для кнопки
    LoadCmdViewAllDeviceCaption
    ChangeStatusTextAndDebug strMessages(114)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuRezimBaseDrvClean_Click
'! Description (Описание)  :   [Меню - Очистка лишних файлов БД]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuRezimBaseDrvClean_Click()
    DeleteUnUsedBase
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuRezimBaseDrvUpdateALL_Click
'! Description (Описание)  :   [Меню - Обновление всех баз поочередно]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuRezimBaseDrvUpdateALL_Click()
    
    SilentReindexAllDB
    ' Обновить список неизвестных дров и описание для кнопки
    LoadCmdViewAllDeviceCaption
    ' возвращаяем обратно стартовый режим
    SelectStartMode , True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuRezimBaseDrvUpdateNew_Click
'! Description (Описание)  :   [Меню - Обновление новых баз поочередно]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuRezimBaseDrvUpdateNew_Click()

    If FindNoDBCount Then
        SilentCheckNoDB
        ' Обновить список неизвестных дров и описание для кнопки
        LoadCmdViewAllDeviceCaption
        ' возвращаяем обратно стартовый режим
        SelectStartMode
    Else
        ChangeStatusTextAndDebug strMessages(68)
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuRunSilentMode_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuRunSilentMode_Click()

    If MsgBox(strMessages(18), vbQuestion + vbYesNo, strProductName) = vbYes Then
        'Команда для программы DPInst работать в тихом режиме
        mbDpInstQuietInstall = True
        ' Включаем тихий режим
        mbSilentRun = True
        ' Начинаем тихую установку
        SilentInstall
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuSaveInfoPC_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuSaveInfoPC_Click()

    Dim strFilePathTo As String

    With New CommonDialog
        .Filter = "Text Files (*.TXT)|*.TXT"
        .DefaultExt = ".txt"
        .InitDir = GetSpecialFolderPath(CSIDL_DESKTOPDIRECTORY)
        .DialogTitle = strMessages(151)
        If mbIsNotebok Then
            If Not OSCurrVersionStruct.ClientOrServer Then
                .FileName = ExpandFileNamebyEnvironment("hwids_%PCMODEL%-Notebook_" & strOSCurrentVersion & "-Server_%OSBIT%")
            Else
                .FileName = ExpandFileNamebyEnvironment("hwids_%PCMODEL%-Notebook_" & strOSCurrentVersion & "_%OSBIT%")
            End If
        Else
            If Not OSCurrVersionStruct.ClientOrServer Then
                .FileName = ExpandFileNamebyEnvironment("hwids_%PCMODEL%_" & strOSCurrentVersion & "-Server_%OSBIT%")
            Else
                .FileName = ExpandFileNamebyEnvironment("hwids_%PCMODEL%_" & strOSCurrentVersion & "_%OSBIT%")
            End If
        End If

        If .ShowSave = True Then
            strFilePathTo = .FileName
        End If

    End With

    If LenB(strFilePathTo) Then
        If PathExists(strResultHwidsExtTxtPath) Then
            CopyFileTo strResultHwidsExtTxtPath, strFilePathTo
        Else

            If SaveHwidsArray2File(strResultHwidsExtTxtPath, arrHwidsLocal) Then
                If PathExists(strResultHwidsExtTxtPath) Then
                    CopyFileTo strResultHwidsExtTxtPath, strFilePathTo
                Else
                    MsgBox strMessages(45) & vbNewLine & strFilePathTo, vbInformation, strProductName
                End If

            Else
                MsgBox strMessages(45) & vbNewLine & strFilePathTo, vbInformation, strProductName
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuService_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuService_Click()
    mnuViewDPInstLog.Enabled = PathExists(strWinDir & "DPINST.LOG")
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuShowHwidsAll_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuShowHwidsAll_Click()

    If IsFormLoaded("frmListHwidAll") = False Then
        frmListHwidAll.Show vbModal, Me
    Else
        frmListHwidAll.FormLoadDefaultParam
        frmListHwidAll.FormLoadAction
        frmListHwidAll.Show vbModal, Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuShowHwidsTxt_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuShowHwidsTxt_Click()

    If PathExists(strHwidsTxtPathView) = False Then
        RunDevconView
    End If

    RunUtilsShell strHwidsTxtPathView, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuShowHwidsXLS_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuShowHwidsXLS_Click()
    OpenTxtInExcel strResultHwidsTxtPath
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuUpdateStatusAll_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuUpdateStatusAll_Click()
    UpdateStatusButtonAll
    ' Обновить список неизвестных дров и описание для кнопки
    LoadCmdViewAllDeviceCaption
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuUpdateStatusTab_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuUpdateStatusTab_Click()
    UpdateStatusButtonTAB
    ' Обновить список неизвестных дров и описание для кнопки
    LoadCmdViewAllDeviceCaption
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuUtils_Click
'! Description (Описание)  :   [Запуск дополнительной утилиты]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuUtils_Click(Index As Integer)

    Dim i         As Long
    Dim PathExe   As String
    Dim PathExe64 As String
    Dim Params    As String
    Dim cmdString As String

    i = Index
    PathExe = PathCollect(arrUtilsList(i, 1))
    PathExe64 = PathCollect(arrUtilsList(i, 2))

    If mbIsWin64 Then
        If LenB(PathExe64) Then
            PathExe = PathExe64
        End If
    End If

    Params = arrUtilsList(i, 3)

    If LenB(Params) = 0 Then
        cmdString = strKavichki & PathExe & strKavichki
    Else
        cmdString = strKavichki & PathExe & strKavichki & strSpace & Params
    End If

    RunUtilsShell cmdString, False, False, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuUtils_DevManView_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuUtils_DevManView_Click()

    If mbIsWin64 Then
        RunUtilsShell strDevManView_Path64
    Else
        RunUtilsShell strDevManView_Path
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuUtils_devmgmt_Click
'! Description (Описание)  :   [Запуск диспетчера устройств]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuUtils_devmgmt_Click()
    RunUtilsShell "devmgmt.msc", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuUtils_DoubleDriver_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuUtils_DoubleDriver_Click()
    RunUtilsShell strDoubleDriver_Path
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuUtils_SIV_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuUtils_SIV_Click()

    If mbIsWin64 Then
        RunUtilsShell strSIV_Path64
    Else
        RunUtilsShell strSIV_Path
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuUtils_UDI_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuUtils_UDI_Click()
    RunUtilsShell strUDI_Path
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuUtils_UnknownDevices_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuUtils_UnknownDevices_Click()
    RunUtilsShell strUnknownDevices_Path, , True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuViewDPInstLog_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuViewDPInstLog_Click()

    Dim strLogPath    As String
    Dim strLogPathNew As String

    strLogPath = strWinDir & "DPINST.LOG"
    strLogPathNew = strWorkTempBackSL & "DPINST.LOG.TXT"

    If PathExists(strLogPath) Then
        CopyFileTo strLogPath, strLogPathNew
        RunUtilsShell strLogPathNew, False
    Else
        If mbDebugStandart Then DebugMode "cmdString - File not exist: " & strLogPath
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optRezim_CaptionBold
'! Description (Описание)  :   [Made Bold caption for Active Rezim Mode]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optRezim_CaptionBold(lngCurrMode As Long)
    Select Case lngCurrMode
    Case 1
        optRezim_Intellect.Font.Bold = True
        optRezim_Ust.Font.Bold = False
        optRezim_Upd.Font.Bold = False
    Case 2
        optRezim_Intellect.Font.Bold = False
        optRezim_Ust.Font.Bold = True
        optRezim_Upd.Font.Bold = False
    Case 3
        optRezim_Intellect.Font.Bold = False
        optRezim_Ust.Font.Bold = False
        optRezim_Upd.Font.Bold = True
    End Select
End Sub
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optRezim_Intellect_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optRezim_Intellect_Click()

    Dim ButtIndex                 As Long
    Dim strSSTabCurrentOSListTemp As String
    Dim i                         As Integer
    Dim i_i                       As Integer
    Dim cntFindUnHideTab          As Integer

    If Not mbFirstStart Then
        ButtIndex = acmdPackFiles.UBound

        For i = 0 To ButtIndex

            If ButtIndex Then

                With acmdPackFiles(i)

                    If Not (.PictureNormal Is Nothing) Then
                        If .PictureNormal = imgNo.Picture Or .PictureNormal = imgNoDB.Picture Then
                            If .Enabled Then
                                .Enabled = False
                                chkPackFiles(i).Enabled = False
                            End If

                            .DropDownEnable = False
                        Else
                            .DropDownEnable = True
                        End If
                    End If

                End With

            Else

                With acmdPackFiles(0)

                    If .Visible Then
                        If Not (.PictureNormal Is Nothing) Then
                            If .PictureNormal = imgNo.Picture Or .PictureNormal = imgNoDB.Picture Then
                                If .Enabled Then
                                    .Enabled = False
                                    chkPackFiles(0).Enabled = False
                                End If

                                .DropDownEnable = False
                            Else
                                .DropDownEnable = True
                            End If
                        End If
                    End If

                End With

            End If

        Next

    End If

    If mbTabBlock Then
        strSSTabCurrentOSListTemp = strSSTabCurrentOSList & strSpace

        For i = 0 To SSTab1.Tabs - 1

            If InStr(strSSTabCurrentOSListTemp, i & strSpace) = 0 Then
                SSTab1.TabEnabled(i) = False

                If mbTabHide Then
                    SSTab1.TabVisible(i) = False
                End If

            Else

                If arrOSList(i).CntBtn = 0 Then
                    SSTab1.TabEnabled(i) = False
                End If
            End If

        Next

    End If

    With SSTab1

        If .Tab <> lngSSTabCurrentOS Then
            If .TabVisible(lngSSTabCurrentOS) Then
                .Tab = lngSSTabCurrentOS
            End If
        End If

    End With

    With cmdRunTask
        .Enabled = FindCheckCount
        .DropDownEnable = True
        .DropDownSeparator = True
        .DropDownSymbol = 6
        .Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdRunTask1", .Caption)
    End With
    
    'заполнение списка на выделение
    LoadListChecked
    ' Изменение размеров FrameScroll и ширины вкладок sstab
    cntFindUnHideTab = FindUnHideTab
    'SSTab1.Visible = cntFindUnHideTab >= 0
    frTabPanel.Visible = cntFindUnHideTab >= 0

    If frTabPanel.Visible Then
        'If SSTab1.Visible Then
        SetTabsWidth cntFindUnHideTab
        SetStartScrollFramePos cntFindUnHideTab
    End If

    ' Активируем возможно блокированные подвкладок
    TabInstBlockOnUpdate False

    ' если активна вкладка 4 то тогда в этом режиме переставляем на стартовую или 0
    If SSTab2(SSTab1.Tab).Tab = 4 Then

        For i_i = SSTab2.LBound To SSTab2.UBound

            If lngStartModeTab2 Then

                ' Если вкладка активна, то выставляем начальную
                If SSTab2(i_i).TabEnabled(lngStartModeTab2) = True Then
                    SSTab2(i_i).Tab = lngStartModeTab2
                Else
                    SSTab2(i_i).Tab = 0
                End If
            End If

        Next

    End If
    
    'BoldCaption
    optRezim_CaptionBold 1

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optRezim_Upd_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optRezim_Upd_Click()

    Dim i                As Integer
    Dim i_i              As Integer
    Dim cntFindUnHideTab As Integer

    If Not mbFirstStart Then

        With acmdPackFiles
            For i = 0 To .UBound
    
                If Not .item(i).Enabled Then
                    .item(i).Enabled = True
                    chkPackFiles(i).Enabled = True
                End If
    
                .item(i).DropDownEnable = False
            Next
        End With

    End If

    If mbTabBlock Then

        With SSTab1
        
            For i = 0 To .Tabs - 1
    
                If Not arrOSList(i).DPFolderNotExist Then
                    If arrOSList(i).CntBtn = 0 Then
                        .TabEnabled(i) = False
                    Else
    
                        If Not .TabVisible(i) Then .TabVisible(i) = True
                        If Not .TabEnabled(i) Then .TabEnabled(i) = True
                    End If
    
                Else
    
                    If mbTabHide Then
                        .TabVisible(i) = False
                    End If
                End If
    
            Next
        End With

    End If

    ' Если кнопка всего одна, то проверяем на какой она вкладке
    If acmdPackFiles.Count = 1 Then
        If acmdPackFiles(0).Visible Then

            With SSTab1

                For i = 0 To .Tabs - 1

                    If .TabVisible(i) Then
                        .Tab = i

                        If StrComp(acmdPackFiles(0).Container.Name, "ctlScrollControl1", vbTextCompare) = 0 Then
                            If acmdPackFiles(0).Container.Index <> .Tab Then
                                .TabEnabled(i) = False
                            End If
                        End If
                    End If

                Next

            End With

        End If
    End If

    With cmdRunTask
        .Enabled = FindCheckCount
        .DropDownEnable = False
        .DropDownSeparator = False
        .DropDownSymbol = 0
        .Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdRunTask", .Caption)
    End With
    
    'заполнение списка на выделение
    LoadListChecked

    ' Если переключились в режим обновления БД путем перехода с 4-ой вкладки, то не делаем ничего
    If Not mbSet2UpdateFromTab4 Or acmdPackFiles.Count = 1 Then
        lngFirstActiveTabIndex = SetFirstEnableTab
        SSTab1.Tab = lngFirstActiveTabIndex
    End If

    ' Изменение размеров FrameScroll и ширины вкладок sstab
    cntFindUnHideTab = FindUnHideTab
    frTabPanel.Visible = cntFindUnHideTab >= 0

    If frTabPanel.Visible Then
        SetTabsWidth cntFindUnHideTab
        SetStartScrollFramePos cntFindUnHideTab
    End If

    ' Блокировка подвкладок
    TabInstBlockOnUpdate True

    ' если активна вкладка c 1-3, то тогда в этом режиме переставляем на 0
    If SSTab2(SSTab1.Tab).Tab Then
        If SSTab2(SSTab1.Tab).Tab < 4 Then

            For i_i = SSTab2.LBound To SSTab2.UBound
                SSTab2(i_i).Tab = 0
            Next

        End If
    End If

    mbSet2UpdateFromTab4 = False
    
    'BoldCaption
    optRezim_CaptionBold 3
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optRezim_Ust_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optRezim_Ust_Click()

    Dim ButtIndex                 As Integer
    Dim i                         As Integer
    Dim i_i                       As Integer
    Dim strSSTabCurrentOSListTemp As String
    Dim cntFindUnHideTab          As Integer

    If Not mbFirstStart Then
        ButtIndex = acmdPackFiles.UBound

        For i = 0 To ButtIndex

            If ButtIndex Then

                With acmdPackFiles(i)

                    If .Enabled = imgNoDB.Picture Then
                        If .Enabled Then
                            .Enabled = False
                            chkPackFiles(i).Enabled = False
                        End If

                    Else

                        If Not .Enabled Then
                            .Enabled = True
                            chkPackFiles(i).Enabled = True
                        End If
                    End If

                    .DropDownEnable = False
                End With

            End If

        Next

    End If

    If mbTabBlock Then
        strSSTabCurrentOSListTemp = strSSTabCurrentOSList & strSpace

        For i = 0 To SSTab1.Tabs - 1

            If InStr(strSSTabCurrentOSListTemp, i & strSpace) = 0 Then
                SSTab1.TabEnabled(i) = False

                If mbTabHide Then
                    SSTab1.TabVisible(i) = False
                End If

            Else

                If arrOSList(i).CntBtn = 0 Then
                    SSTab1.TabEnabled(i) = False
                End If
            End If

        Next

    End If

    With SSTab1

        If .Tab <> lngSSTabCurrentOS Then
            If .TabVisible(lngSSTabCurrentOS) Then
                .Tab = lngSSTabCurrentOS
            End If
        End If

    End With

    With cmdRunTask
        .Enabled = FindCheckCount
        .DropDownEnable = False
        .DropDownSeparator = False
        .DropDownSymbol = 0
        .Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdRunTask1", .Caption)
    End With
    
    'заполнение списка на выделение
    LoadListChecked
    ' Изменение размеров FrameScroll и ширины вкладок sstab
    cntFindUnHideTab = FindUnHideTab
    'SSTab1.Visible = cntFindUnHideTab >= 0
    frTabPanel.Visible = cntFindUnHideTab >= 0

    If frTabPanel.Visible Then
        'If SSTab1.Visible Then
        SetTabsWidth cntFindUnHideTab
        SetStartScrollFramePos cntFindUnHideTab
    End If

    ' Активируем возможно блокированные подвкладок
    TabInstBlockOnUpdate False

    ' если активна вкладка 4 то тогда в этом режиме переставляем на стартовую или 0
    If SSTab2(SSTab1.Tab).Tab = 4 Then

        For i_i = SSTab2.LBound To SSTab2.UBound

            If lngStartModeTab2 Then

                ' Если вкладка активна, то выставляем начальную
                If SSTab2(i_i).TabEnabled(lngStartModeTab2) = True Then
                    SSTab2(i_i).Tab = lngStartModeTab2
                Else
                    SSTab2(i_i).Tab = 0
                End If
            End If

        Next

    End If

    'BoldCaption
    optRezim_CaptionBold 2
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub pbProgressBar_Resize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub pbProgressBar_Resize()
    cmdBreakUpdateDB.Left = (pbProgressBar.Width - cmdBreakUpdateDB.Width) / 2
    cmdBreakUpdateDB.Top = (pbProgressBar.Height - cmdBreakUpdateDB.Height) / 2
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SSTab1_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PreviousTab (Integer)
'!--------------------------------------------------------------------------------
Private Sub SSTab1_Click(PreviousTab As Integer)

    If acmdPackFiles(0).Visible Then
        If acmdPackFiles.UBound > 1 Then
            mbNextTab = True
        End If
    End If

    If ctlScrollControl1.UBound >= SSTab1.Tab Then
        If arrOSList(SSTab1.Tab).CntBtn Then
            ctlScrollControl1(SSTab1.Tab).Refresh
        End If
    End If

    If optRezim_Upd.Value Then

        ' если активна вкладка c 1-3, то тогда в этом режиме переставляем на 0
        If SSTab2(SSTab1.Tab).Tab Then
            If SSTab2(SSTab1.Tab).Tab < 4 Then
                SSTab2(SSTab1.Tab).Tab = 0
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SSTab2_Click
'! Description (Описание)  :   [Нажатие кнопки на SStab2]
'! Parameters  (Переменные):   Index (Integer)
'                              PreviousTab (Integer)
'!--------------------------------------------------------------------------------
Private Sub SSTab2_Click(Index As Integer, PreviousTab As Integer)
    
    If SSTab2(Index).Tab = 0 Then
        If PreviousTab Then
            ctlScrollControl1(Index).Visible = False
        End If
    End If

    StartReOrderBtnOnTab2 Index, PreviousTab

    If SSTab2(Index).Tab = 0 Then
        If PreviousTab Then
            If ctlScrollControl1(Index).Visible = False Then
                ctlScrollControl1(Index).Visible = True
            End If
        End If
    End If

End Sub

