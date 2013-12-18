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
      Name            =   "Courier New"
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
      TabIndex        =   18
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
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   11265
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9345
      Visible         =   0   'False
      Width           =   11265
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
      Begin prjDIADBS.ctlJCbutton cmdBreakUpdateDB 
         Height          =   385
         Left            =   4200
         TabIndex        =   26
         Top             =   75
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
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
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
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
         ThemeColor      =   1
         Begin prjDIADBS.LabelW lblPCInfo 
            Height          =   255
            Left            =   75
            TabIndex        =   31
            Top             =   850
            Width           =   10995
            _ExtentX        =   19394
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
            Caption         =   "Модель PC:"
         End
         Begin prjDIADBS.LabelW lblOsInfo 
            Height          =   255
            Left            =   75
            TabIndex        =   30
            ToolTipText     =   "Starting ""System Information Viewer"""
            Top             =   480
            Width           =   10995
            _ExtentX        =   19394
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
            Left            =   200
            TabIndex        =   4
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   4155
            Left            =   0
            TabIndex        =   1
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
               TabIndex        =   2
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
                  TabIndex        =   22
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
                  TabIndex        =   24
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
                  TabIndex        =   21
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
                  TabIndex        =   23
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
                  TabIndex        =   25
                  Top             =   350
                  Width           =   4095
                  _ExtentX        =   7223
                  _ExtentY        =   2778
                  AutoScrollToFocus=   0   'False
               End
               Begin prjDIADBS.LabelW lblNoDP4Mode 
                  Height          =   285
                  Left            =   105
                  TabIndex        =   28
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
               TabIndex        =   29
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
         Begin prjDIADBS.ctlXpButton acmdPackFiles 
            Height          =   555
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   4200
            Visible         =   0   'False
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   979
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
               Height          =   330
               Left            =   120
               TabIndex        =   5
               Top             =   480
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
               Text            =   "frmMain.frx":0144
               CueBanner       =   "frmMain.frx":0180
               Sorted          =   -1  'True
            End
            Begin prjDIADBS.ctlJCbutton cmdCheck 
               Height          =   390
               Left            =   120
               TabIndex        =   6
               Top             =   840
               Width           =   3075
               _ExtentX        =   5424
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
               ButtonStyle     =   8
               BackColor       =   12244692
               Caption         =   "Выделить"
               PictureAlign    =   0
               PicturePushOnHover=   -1  'True
               PictureShadow   =   -1  'True
               CaptionEffects  =   0
               TooltipBackColor=   0
               ColorScheme     =   3
            End
         End
         Begin prjDIADBS.ctlJCbutton cmdRunTask 
            Height          =   675
            Left            =   120
            TabIndex        =   27
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
            Caption         =   "Выполнить задание для выбранных пакетов драйверов"
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            CaptionEffects  =   0
            TooltipBackColor=   0
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
         Begin prjDIADBS.ctlJCbutton cmdViewAllDevice 
            Height          =   510
            Left            =   120
            TabIndex        =   7
            Top             =   925
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
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin prjDIADBS.ctlJCbutton optRezim_Intellect 
            Height          =   510
            Left            =   120
            TabIndex        =   8
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
            Mode            =   2
            Value           =   -1  'True
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   3
         End
         Begin prjDIADBS.ctlJCbutton optRezim_Upd 
            Height          =   510
            Left            =   5280
            TabIndex        =   9
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
            Left            =   2640
            TabIndex        =   10
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
            Mode            =   2
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            CaptionEffects  =   0
            TooltipBackColor=   0
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
         GradientHeaderStyle=   1
         Begin VB.PictureBox imgOkAttentionOld 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            DrawStyle       =   5  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   19
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
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   510
            Left            =   6840
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   350
            Width           =   510
         End
         Begin VB.PictureBox imgNoDB 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   6000
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgNo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   5160
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgOkOld 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   4320
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgOkNew 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3480
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgOkAttention 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2640
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
         Begin VB.PictureBox imgOK 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00F0D4C0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1800
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   12
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
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   960
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   350
            Width           =   495
         End
      End
   End
   Begin prjDIADBS.ToolTip TTStatusIcon 
      Left            =   900
      Top             =   9000
      _ExtentX        =   450
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualStyles    =   -1  'True
      Title           =   "frmMain.frx":01A0
   End
   Begin prjDIADBS.ToolTip TT 
      Left            =   300
      Top             =   9000
      _ExtentX        =   450
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualStyles    =   -1  'True
      Title           =   "frmMain.frx":01C0
   End
   Begin VB.Menu mnuRezim 
      Caption         =   "Обновление баз данных"
      Begin VB.Menu mnuRezimBaseDrvUpdateALL 
         Caption         =   "Обновить базы для ВСЕХ пакетов драйверов"
      End
      Begin VB.Menu mnuRezimBaseDrvUpdateNew 
         Caption         =   "Обновить базы только для НОВЫХ пакетов драйверов"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRezimBaseDrvClean 
         Caption         =   "Удалить файлы баз данных отсутствующих пакетов драйверов"
      End
      Begin VB.Menu mnuDelDuplicateOldDP 
         Caption         =   "Удалить устаревшие версии пакетов драйверов"
      End
      Begin VB.Menu mnuSep26 
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
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowHwidsAll 
         Caption         =   "Показать ПОЛНЫЙ СПИСОК УСТРОЙСТВ компьютера"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuShowHwidsAllBase 
         Caption         =   "Поиск по всей базе драйверов"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep4 
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
      Begin VB.Menu mnuSep21 
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
      Begin VB.Menu mnuSep25 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateRestorePoint 
         Caption         =   "Создать точку восстановления системы"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateBackUp 
         Caption         =   "Создать резервную копию драйверов"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDPInstLog 
         Caption         =   "Просмотреть DPinst.log"
      End
      Begin VB.Menu mnuSep7 
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
      Begin VB.Menu mnuSep8 
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
      Begin VB.Menu mnuSep9 
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
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckUpd 
         Caption         =   "Проверить обновление программы"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModulesVersion 
         Caption         =   "Модули..."
      End
      Begin VB.Menu mnuSep12 
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
   Begin VB.Menu mnuContextMenu 
      Caption         =   "Контекстное меню"
      Begin VB.Menu mnuContextXLS 
         Caption         =   "Открыть файл базы данных в программе Excel"
      End
      Begin VB.Menu mnuContextTxt 
         Caption         =   "Открыть файл базы данных в текстовом виде"
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextToolTip 
         Caption         =   "Показать список доступных драйверов для компьютера"
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextUpdStatus 
         Caption         =   "Обновить статус пакета драйверов"
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextEditDPName 
         Caption         =   "Изменить отображаемое имя пакета драйверов в программе"
      End
      Begin VB.Menu mnuSep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextTestDRP 
         Caption         =   "Протестировать данный пакет драйверов программой 7-zip"
      End
      Begin VB.Menu mnuSep18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextDeleteDRP 
         Caption         =   "Удалить пакет драйверов"
      End
      Begin VB.Menu mnuSep19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextDeleteDevIDs 
         Caption         =   "Удалить драйвера устройств:"
         Begin VB.Menu mnuContextDeleteDevIDDesc 
            Caption         =   "Список драйверов доступных для удаления"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSep20 
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
         Begin VB.Menu mnuSep23 
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
      Begin VB.Menu mnuContextInstall 
         Caption         =   "Обычная установка"
         Index           =   0
      End
      Begin VB.Menu mnuContextInstall 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContextInstall 
         Caption         =   "Выборочная установка"
         Index           =   2
      End
      Begin VB.Menu mnuContextInstall 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuContextInstall 
         Caption         =   "Распаковать в каталог - Все подобранные драйвера"
         Index           =   4
      End
      Begin VB.Menu mnuContextInstall 
         Caption         =   "Распаковать в каталог - Выбрать драйвера..."
         Index           =   5
      End
   End
   Begin VB.Menu mnuMainLang 
      Caption         =   "Язык"
      Begin VB.Menu mnuLangStart 
         Caption         =   "Использовать выбранный язык при запуске (отмена автовыбора)"
      End
      Begin VB.Menu mnuSep17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLang 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngCntBtn                       As Long
Private mbNextTab                       As Boolean
Private strCurSelButtonPath             As String
Private mbStatusHwid                    As Boolean
Private mbStatusNewer                   As Boolean
Private mbStatusOlder                   As Boolean
Private lngSSTabCurrentOS               As Long
Private strSSTabCurrentOSList           As String
Private strCmbChkBtnListElement1        As String
Private strCmbChkBtnListElement2        As String
Private strCmbChkBtnListElement3        As String
Private strCmbChkBtnListElement4        As String
Private strCmbChkBtnListElement5        As String
Private strCmbChkBtnListElement6        As String
Private strContextInstall1              As String
Private strContextInstall2              As String
Private strContextInstall3              As String
Private strContextInstall4              As String
Private strTTipTextTitle                As String
Private strTTipTextFileSize             As String
Private strTTipTextClassDRV             As String
Private strTTipTextDrv2Install          As String
Private strTTipTextDrv4UnsupOS          As String
Private strTTipTextTitleStatus          As String
Private mbUnpackAdditionalFile          As Boolean
Private lngFirstActiveTabIndex          As Long
Private mbNoSupportedOS                 As Boolean
Private mbNotSupportedDevDB             As Boolean
Private strSSTabTypeDPTab1              As String
Private strSSTabTypeDPTab2              As String
Private strSSTabTypeDPTab3              As String
Private strSSTabTypeDPTab4              As String
Private strSSTabTypeDPTab5              As String
Private mbSet2UpdateFromTab4            As Boolean
Private lngNotFinedDriversInDP          As Long
Private mbLoadAppEnd                    As Boolean

Private objHashOutput                   As Scripting.Dictionary
Private objHashOutput2                  As Scripting.Dictionary
Private objHashOutput3                  As Scripting.Dictionary

' Положение кнопок пакетов драйверов на основной вкладке
Private strFormName                     As String

'! -----------------------------------------------------------
'!  Функция     :  acmdPackFiles_Click
'!  Переменные  :  Index As Integer
'!  Описание    :  Обработка События нажатия кнопки
'! -----------------------------------------------------------
Private Sub acmdPackFiles_Click(Index As Integer)

Dim strPackFileName                     As String
Dim strPathDRP                          As String
Dim strPathDevDB                        As String
Dim TimeScriptRun                       As Long
Dim TimeScriptFinish                    As Long
Dim AllTimeScriptRun                    As String
Dim strPackFileName_woExt               As String
Dim cmdString                           As String
Dim ArchTempPath                        As String
Dim strDevPathShort                     As String
Dim DPInstExitCode                      As Long
Dim ReadExitCodeString                  As String

    DebugMode "acmdPackFiles_Click-Start"
    DebugMode vbTab & "acmdPackFiles_Click: Index=" & Index

    strPathDRPList = vbNullString
    BlockControl False

    If mbDevParserRun Then
        MsgBox strMessages(22), vbInformation, strProductName
    Else
        mbStatusHwid = True

        strPackFileName = acmdPackFiles(Index).Tag

        'Если пакет драйверов реальный, то....
        If LenB(strPackFileName) > 0 Then

            FlatBorderButton acmdPackFiles(Index).hWnd
            acmdPackFiles(Index).Refresh

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
                    Set acmdPackFiles(Index).Picture = imgUpdBD.Picture
                    ' изменение положения прогресс-анимации
                    ChangeStatusTextAndDebug strMessages(61) & " " & strPackFileName, , , , strMessages(128)
                    TimeScriptRun = GetTickCount
                    ' запуск построения БД
                    DevParserByRegExp strPackFileName, strPathDRP, strPathDevDB
                    ' Обновление подсказки
                    ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, Index, True
                    TimeScriptFinish = GetTickCount
                    AllTimeScriptRun = CalculateTime(TimeScriptRun, TimeScriptFinish, True)
                    ChangeStatusTextAndDebug strMessages(62) & " " & AllTimeScriptRun, "DevParserByRegExp: All time for create Base for file finish: " & strPackFileName
                End If

'------------------------------------------------------
'-------- Установка всех драйверов в пакете -----------
'------------------------------------------------------
            ElseIf optRezim_Ust.Value Then
                ChangeStatusTextAndDebug strMessages(63) & " " & strPackFileName, , , , strMessages(129)
                'Имя папки с распакованными драйверами
                strPackFileName_woExt = FileName_woExt(strPackFileName)
                ArchTempPath = strWorkTempBackSL & strPackFileName_woExt

                'Извлечение драйверов из файла
                If UnPackDPFile(strPathDRP, strPackFileName, ALL_FILES, ArchTempPath) = False Then
                    If Not mbSilentRun Then
                        MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                    End If

                    ChangeStatusTextAndDebug strMessages(13) & " " & strPackFileName, "Error on run : " & cmdString
                Else
                    ' установка драйверов
                    DPInstExitCode = RunDPInst(ArchTempPath)
                End If

                ' Обновление подсказки
                ReadExitCodeString = ReadExitCode(DPInstExitCode)

                If DPInstExitCode <> 0 Then
                    If DPInstExitCode <> -2147483648# Then
                        If InStr(1, ReadExitCodeString, "Cancel or Nothing to Install", vbTextCompare) = 0 Then
                            ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, Index
                        End If
                    End If
                End If

                ChangeStatusTextAndDebug strMessages(64) & " (" & strPackFileName & "): " & ReadExitCodeString, "Install from : " & strPackFileName & " finished."
'------------------------------------------------------
'------- Установка избранных драйверов в пакете--------
'------------------------------------------------------
            Else
                ChangeStatusTextAndDebug strMessages(63) & " " & strPackFileName, , , , strMessages(129)
                'Имя папки с распакованными драйверами
                strPackFileName_woExt = FileName_woExt(strPackFileName)

                ' если выборочная установка, то получаем список каталогов для распаковки
                If mbooSelectInstall Then

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
                        FlatBorderButton acmdPackFiles(Index).hWnd, False
                        acmdPackFiles(Index).Refresh
                        BlockControl True
                        ChangeStatusTextAndDebug strMessages(65) & " " & strPackFileName
                        cmdRunTask.Enabled = FindCheckCount(False)
                        Exit Sub
                    End If

                Else
                    ' иначе список строится сам
                    Dim strTemp_x() As String
                    Dim strTempLine_x() As String
                    Dim i_arr As Long
                    
                    strTemp_x = Split(arrTTip(Index), vbNewLine)
                    
                    For i_arr = LBound(strTemp_x) To UBound(strTemp_x)
                        strTempLine_x = Split(strTemp_x(i_arr), " | ")
        
                        If LenB(Trim$(strTemp_x(i_arr))) Then
                            strDevPathShort = Trim$(strTempLine_x(1))
                            ' Если данного пути нет в списке, то добавляем
                            If InStr(1, strPathDRPList, strDevPathShort, vbTextCompare) = 0 Then
                                strPathDRPList = AppendStr(strPathDRPList, strDevPathShort, " ")
                            End If
                        End If
                    Next i_arr
                End If
                
                strPathDRPList = Trim$(strPathDRPList)

                ' Если по каким либо причинам список папок не получился, то извлекаем все.
                If LenB(strPathDRPList) = 0 Then
                    strPathDRPList = ALL_FILES
                End If

                ArchTempPath = strWorkTempBackSL & strPackFileName_woExt

                'Извлечение драйверов из файла
                If UnPackDPFile(strPathDRP, strPackFileName, strPathDRPList, ArchTempPath) = False Then
                    If Not mbSilentRun Then
                        MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                    End If

                    ChangeStatusTextAndDebug strMessages(13) & " " & strPackFileName, "Error on run : " & cmdString
                Else
                    ' установка драйверов
                    DPInstExitCode = RunDPInst(ArchTempPath)
                    ReadExitCodeString = ReadExitCode(DPInstExitCode)

                    If DPInstExitCode <> 0 Then
                        If DPInstExitCode <> -2147483648# Then
                            If InStr(1, ReadExitCodeString, "Cancel or Nothing to Install", vbTextCompare) = 0 Then
                                ' Обрабатываем файл finish
                                WorkWithFinish strPathDRP, strPackFileName, ArchTempPath, strPathDRPList
                                ' Обновление подсказки
                                ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, Index
                            End If
                        End If
                    End If
                End If

                ChangeStatusTextAndDebug strMessages(64) & " " & strPackFileName & " finish. " & ReadExitCodeString, "Install from : " & strPackFileName & " finish."
            End If

            mbDevParserRun = False
            BlockControl True
            FlatBorderButton acmdPackFiles(Index).hWnd, False

            If Not optRezim_Upd.Value Then
                ' Удаление временных файлов
                strPackFileName_woExt = FileName_woExt(strPackFileName)
                ArchTempPath = strWorkTempBackSL & strPackFileName_woExt

                If PathExists(ArchTempPath) Then
                    DelRecursiveFolder ArchTempPath
                End If
            End If
        End If

        ' Проверка выделенных пакетов
        cmdRunTask.Enabled = FindCheckCount(False)
        acmdPackFiles(Index).Refresh
    End If

    DebugMode "acmdPackFiles_Click-End"

End Sub

Private Sub acmdPackFiles_ClickMenu(Index As Integer, mnuIndex As Integer)
    mbGroupTask = False

    Select Case mnuIndex

        Case 0
            mbooSelectInstall = False
            acmdPackFiles_Click Index

        Case 2
            CurrentSelButtonIndex = Index
            mbooSelectInstall = True
            acmdPackFiles_Click Index

    End Select

End Sub

Private Sub acmdPackFiles_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = 32 Then
        chkPackFiles(Index).Value = Not chkPackFiles(Index).Value
        FindCheckCount
    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  acmdPackFiles_MouseDown
'!  Переменные  :  Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single
'!  Описание    :  Обработка события нажатия правой кнопкой мыши
'! -----------------------------------------------------------
Private Sub acmdPackFiles_MouseDown(Index As Integer, _
                                    Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

Dim strPackFileName                     As String
Dim strPathDevDB                        As String

    If Button = vbRightButton Then
        mnuContextTxt.Enabled = True
        mnuContextXLS.Enabled = True
        mnuContextToolTip.Enabled = True
        strPackFileName = acmdPackFiles(Index).Tag
        strPathDevDB = arrOSList(SSTab1.Tab).devIDFolderFull

        If Not CheckExistDB(strPathDevDB, strPackFileName) Then
            mnuContextTxt.Enabled = False
            mnuContextXLS.Enabled = False
            mnuContextToolTip.Enabled = False
            mnuContextDeleteDevIDs.Enabled = False
            mnuContextCopyHWIDs.Enabled = False
        Else
            ' создаем меню для удаления драйверов устройств
            CreateMenuDevIDIndexDelMenu arrDevIDs(Index)
            ' создаем меню для копирования HWID устройств
            CreateMenuDevIDIndexCopyMenu arrDevIDs(Index)
        End If

        If acmdPackFiles(Index).Picture = imgNo.Picture Then
            mnuContextToolTip.Enabled = False
            mnuContextDeleteDevIDs.Enabled = False
            mnuContextCopyHWIDs.Enabled = False
        End If

        CurrentSelButtonIndex = Index
        OpenContextMenu Me, Me.mnuContextMenu
    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  BaseUpdateOrRunTask
'!  Переменные  :  Optional mbOnlyNew As Boolean = False
'!  Описание    :  Обновление всех баз или только новых поочередно
'! -----------------------------------------------------------
Private Sub BaseUpdateOrRunTask(Optional ByVal mbOnlyNew As Boolean = False, _
                                Optional ByVal mbTasks As Boolean = False)

Dim ButtIndex                           As Long
Dim ButtCount                           As Long
Dim i                                   As Integer
Dim TimeScriptRun                       As Long
Dim TimeScriptFinish                    As Long
Dim AllTimeScriptRun                    As String
Dim miPbInterval                        As Long
Dim miPbNext                            As Long
Dim strTextNew                          As String
Dim mbDpNoDBExist                       As Boolean
Dim strMsg                              As String
Dim lngFindCheckCountTemp               As Long
Dim lngSStabStart                       As Long
Dim lngNumButtOnTab                     As Long

    DebugMode "BaseUpdateOrRunTask-Start"
    mbBreakUpdateDBAll = False
    lngSStabStart = SSTab1.Tab
    strTextNew = " "
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

    If ButtIndex > 0 Then
        ' В цикле обрабатываем обновление
        miPbInterval = 1000 / ButtCount

        If mbTasks Then
            lngFindCheckCountTemp = FindCheckCount

            If lngFindCheckCountTemp > 0 Then
                miPbInterval = 1000 / lngFindCheckCountTemp
            End If
        End If

        miPbNext = 0

        For i = 0 To ButtIndex

            lngNumButtOnTab = arrOSList(SSTab1.Tab).CntBtn

            Do While i >= lngNumButtOnTab
                SSTab1.Tab = SSTab1.Tab + 1
                DoEvents
                Sleep 100
                lngNumButtOnTab = arrOSList(SSTab1.Tab).CntBtn
            Loop

            ' Прерываем процесс обновления
            If mbBreakUpdateDBAll Then
                MsgBox strMessages(27) & vbNewLine & acmdPackFiles(i).Tag, vbInformation, strProductName
                Exit For
            End If

            If mbOnlyNew Then
                If acmdPackFiles(i).Picture = imgNoDB.Picture Then
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
            If acmdPackFiles(0).Picture = imgNoDB.Picture Then
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
        ChangeStatusTextAndDebug strMessages(66) & " " & AllTimeScriptRun, , True
    Else

        If mbDpNoDBExist Then
            ChangeStatusTextAndDebug strMessages(67) & " " & AllTimeScriptRun, , True
        Else
            ChangeStatusTextAndDebug strMessages(68), , True
        End If
    End If

    ChangeFrmMainCaption
    pbProgressBar.Visible = False
    ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
    
    cmdBreakUpdateDB.Visible = False
    BlockControl True
    
TheEnd:
    mbTasks = False
    SSTab1.Tab = lngSStabStart
    DoEvents
    Sleep 100
    DebugMode "BaseUpdateOrRunTask-End"

End Sub

'! -----------------------------------------------------------
'!  Функция     :  BlockControl
'!  Описание    :  Блокировка(Разблокировка) некоторых элементов формы при работе сложных функций
'! -----------------------------------------------------------
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

'! -----------------------------------------------------------
'!  Функция     :  BlockControlEx
'!  Переменные  :  ByVal mbBlock As Boolean
'!  Описание    :  Блокировка(Разблокировка) элементов
'! -----------------------------------------------------------
Private Sub BlockControlEx(ByVal mbBlock As Boolean)
    mnuRunSilentMode.Enabled = mbBlock
    optRezim_Ust.Enabled = mbBlock
    optRezim_Intellect.Enabled = mbBlock
    optRezim_Upd.Enabled = mbBlock
End Sub

'! -----------------------------------------------------------
'!  Функция     :  BlockControlInNoDP
'!  Переменные  :  ByVal mbBlock As Boolean
'!  Описание    :  Блокировка(Разблокировка) элементов если нет пакетов драйверов
'! -----------------------------------------------------------
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

'! -----------------------------------------------------------
'!  Функция     :  CalculateUnknownDrivers
'!  Описание    :  Подсчитываем кол-во неизвестных драйверов
'! -----------------------------------------------------------
Private Function CalculateUnknownDrivers() As Long

Dim ii                                  As Long
Dim lngCountUnknown                     As Long

    For ii = LBound(arrHwidsLocal) To UBound(arrHwidsLocal)

        If arrHwidsLocal(ii).DRVExist = 0 Then

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

'! -----------------------------------------------------------
'!  Функция     :  ChangeFrmMainCaption
'!  Переменные  :  Optional ByVal lngPercentage As Long
'!  Описание    :  Изменение Caption Формы
'! -----------------------------------------------------------
Private Sub ChangeFrmMainCaption(Optional ByVal lngPercentage As Long)

Dim strProgressValue                    As String

    Select Case strPCLangCurrentID

        Case "0419"
            strFrmMainCaptionTemp = "Помощник установки драйверов"
            strFrmMainCaptionTempDate = " (Дата релиза: "

        Case Else
            strFrmMainCaptionTemp = "Drivers Installer Assistant"
            strFrmMainCaptionTempDate = " (Date Build: "

    End Select

    If lngPercentage Mod 999 Then
        If ctlProgressBar1.Visible Then
            strProgressValue = (lngPercentage \ 10) & "% (" & ctlUcStatusBar1.PanelText(1) & ") - "
        End If
    End If

    If LenB(strThisBuildBy) = 0 Then
        Me.Caption = strProgressValue & strFrmMainCaptionTemp & " v." & strProductVersion & " @" & App.CompanyName
        ' & strProgressValue
    Else
        Me.Caption = strProgressValue & strFrmMainCaptionTemp & " v." & strProductVersion & " " & strThisBuildBy
        ' & strProgressValue
    End If

End Sub

Private Sub ChangeMenuCaption()

Dim ButtIndex                           As Long
Dim i                                   As Long

    ButtIndex = acmdPackFiles.UBound

    If ButtIndex > 0 Then

        For i = 0 To ButtIndex

            With acmdPackFiles(i)

                If .MenuCount > 0 Then
                    .MenuCaption(0) = strContextInstall1
                    .MenuCaption(2) = strContextInstall2
                    .Refresh
                End If
            End With

        Next
    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  ChangeStatusAndPictureButton
'!  Переменные  :  strPathDevDB, strPackFileName, ButtonIndex As Integer
'!  Описание    :  Присваиваем картинку в соответсвии с наличием БД к файлу
'! -----------------------------------------------------------
Private Function ChangeStatusAndPictureButton(ByVal strPathDevDB As String, _
                                              ByVal strPackFileName As String, _
                                              ByVal ButtonIndex As Long) As String

Dim strTextHwids                        As String
Dim mbUnSuppOS                          As Boolean

    DebugMode str3VbTab & "ChangeStatusAndPictureButton-Start"
    DebugMode str4VbTab & "ChangeStatusAndPictureButton: strPackFileName=" & strPackFileName
    ' Небольшое прерывание для большего отклика от приложения
    DoEvents
    ChangeStatusAndPictureButton = vbNullString

    With acmdPackFiles(ButtonIndex)

        If CheckExistDB(strPathDevDB, strPackFileName) Then

            ' Ищем совпадения HWID в DP в новом режиме
            If mbFirstStart Then
                If mbLoadUnSupportedOS Then
                    strTextHwids = FindHwidInBaseNew(strPathDevDB, UCase$(FileNameFromPath(strPackFileName)), ButtonIndex)
                Else

                    If InStr(arrOSList(SSTab1.Tab).Ver, strOsCurrentVersion) Then
                        If arrOSList(SSTab1.Tab).is64bit = 2 Or arrOSList(SSTab1.Tab).is64bit = 3 Then
                            strTextHwids = FindHwidInBaseNew(strPathDevDB, UCase$(FileNameFromPath(strPackFileName)), ButtonIndex)
                        Else

                            If mbIsWin64 = CBool(arrOSList(SSTab1.Tab).is64bit) Then
                                strTextHwids = FindHwidInBaseNew(strPathDevDB, UCase$(FileNameFromPath(strPackFileName)), ButtonIndex)
                            Else
                                mbUnSuppOS = True
                            End If
                        End If

                    Else
                        mbUnSuppOS = True
                    End If
                End If

            Else
                strTextHwids = FindHwidInBaseNew(strPathDevDB, UCase$(FileNameFromPath(strPackFileName)), ButtonIndex)
            End If

            If LenB(strTextHwids) > 0 Then
                ChangeStatusAndPictureButton = strTextHwids
                DebugMode str4VbTab & "Hwids in file for PC: " & vbNewLine & strTextHwids

                If mbStatusHwid Then
                    If mbStatusNewer Then
                        Set .Picture = Nothing
                        Set .Picture = imgOkNew.Picture
                        DebugMode str3VbTab & "*ImageForButton: imgOkNew"
                    ElseIf mbStatusOlder Then
                        Set .Picture = Nothing
                        Set .Picture = imgOkOld.Picture
                        DebugMode str3VbTab & "ImageForButton: imgOkOld"
                    Else
                        Set .Picture = Nothing
                        Set .Picture = imgOK.Picture
                        DebugMode str3VbTab & "ImageForButton: imgOK"
                    End If

                Else

                    If mbIgnorStatusHwid Then
                        If mbDRVNotInstall Then
                            If mbStatusNewer Then
                                Set .Picture = Nothing
                                Set .Picture = imgOkAttentionNew.Picture
                                DebugMode str3VbTab & "ImageForButton: imgOkAttentionNew"
                            ElseIf mbStatusOlder Then
                                Set .Picture = Nothing
                                Set .Picture = imgOkAttentionOLD.Picture
                                DebugMode str3VbTab & "ImageForButton: imgOkAttentionOLD"
                            Else
                                Set .Picture = Nothing
                                Set .Picture = imgOkAttention.Picture
                                DebugMode str3VbTab & "ImageForButton: imgOkAttention"
                            End If

                        Else

                            If mbStatusNewer Then
                                Set .Picture = Nothing
                                Set .Picture = imgOkNew.Picture
                                DebugMode str3VbTab & "ImageForButton: imgOkNew"
                            ElseIf mbStatusOlder Then
                                Set .Picture = Nothing
                                Set .Picture = imgOkOld.Picture
                                DebugMode str3VbTab & "ImageForButton: imgOkOld"
                            Else
                                Set .Picture = Nothing
                                Set .Picture = imgOK.Picture
                                DebugMode str3VbTab & "ImageForButton: imgOK"
                            End If
                        End If

                    Else

                        If mbStatusNewer Then
                            Set .Picture = Nothing
                            Set .Picture = imgOkAttentionNew.Picture
                            DebugMode str3VbTab & "ImageForButton: imgOkAttentionNew"
                        ElseIf mbStatusOlder Then
                            Set .Picture = Nothing
                            Set .Picture = imgOkAttentionOLD.Picture
                            DebugMode str3VbTab & "ImageForButton: imgOkAttentionOLD"
                        Else
                            Set .Picture = Nothing
                            Set .Picture = imgOkAttention.Picture
                            DebugMode str3VbTab & "ImageForButton: imgOkAttention"
                        End If
                    End If
                End If

                If .MenuCount = 0 Then
                    .AddMenu strContextInstall1
                    .AddMenu "-"
                    .AddMenu strContextInstall2
                End If

                .MenuExist = optRezim_Intellect.Value
            Else
                Set .Picture = Nothing
                Set .Picture = imgNo.Picture
                DebugMode str3VbTab & "ImageForButton: imgNo"
                .MenuExist = False

                If mbUnSuppOS Then
                    ChangeStatusAndPictureButton = "Unsupported"
                End If
            End If

        Else
            Set .Picture = Nothing
            Set .Picture = imgNoDB.Picture
            DebugMode str3VbTab & "ImageForButton: imgNoDB"
            .MenuExist = False
        End If
    End With

    DebugMode str3VbTab & "ChangeStatusAndPictureButton-End"

End Function

'! -----------------------------------------------------------
'!  Функция     :  CheckAllButton
'!  Переменные  :  ByVal mbCheckAll As Boolean
'!  Описание    :  Выделение всех кнопок
'! -----------------------------------------------------------
Private Sub CheckAllButton(ByVal mbCheckAll As Boolean)

Dim i                                   As Long

    For i = acmdPackFiles.LBound To acmdPackFiles.UBound

        If Not (acmdPackFiles(i).Picture Is Nothing) Then
            If acmdPackFiles(i).Visible Then
                chkPackFiles(i).Value = mbCheckAll
            End If
        End If

    Next
    FindCheckCount

End Sub

'! -----------------------------------------------------------
'!  Функция     :  CheckExistDB
'!  Переменные  :  ByRef DevDBPath As String, ByRef strPackFileName As String
'!  Возвр. знач.:  As Boolean
'!  Описание    :  Проверяет есть ли txt/hwid файл для данного архива
'! -----------------------------------------------------------
Private Function CheckExistDB(ByVal strDevDBPath As String, _
                              ByVal strPackFileName As String) As Boolean

Dim strFileNameDevDB                    As String
Dim strPathFileNameDevDB                As String
Dim strPathFileNameDevDBHwid            As String
Dim lngFileDBSize                       As Long

    DebugMode str4VbTab & "CheckExistDB-Start"
    CheckExistDB = False
    
    strFileNameDevDB = Replace$(strPackFileName, ".7Z", ".TXT", , , vbTextCompare)
    If InStr(1, strPackFileName, ".zip", vbTextCompare) Then
        strFileNameDevDB = Replace$(strPackFileName, ".ZIP", ".TXT", , , vbTextCompare)
    End If
    
    strDevDBPath = BackslashAdd2Path(strDevDBPath)
    If Not mbDP_Is_aFolder Then
        strPathFileNameDevDB = PathCombine(strDevDBPath, FileNameFromPath(strFileNameDevDB))
        strPathFileNameDevDBHwid = Replace$(strPathFileNameDevDB, ".TXT", ".HWID")
    Else
        strPathFileNameDevDB = PathCombine(strDevDBPath, FileNameFromPath(strPackFileName) & ".TXT")
        strPathFileNameDevDBHwid = Replace$(strPathFileNameDevDB, ".TXT", ".HWID")
    End If
    strCurSelButtonPath = strPathFileNameDevDB

    If PathExists(strPathFileNameDevDBHwid) Then
        lngFileDBSize = GetFileSizeByPath(strPathFileNameDevDBHwid)
        DebugMode str5VbTab & "CheckExistDB: Find file=" & strPathFileNameDevDBHwid & " (FileSize: " & lngFileDBSize & " bytes)"

        If lngFileDBSize > 0 Then
            If PathExists(strPathFileNameDevDB) Then
                lngFileDBSize = GetFileSizeByPath(strPathFileNameDevDB)
                DebugMode str5VbTab & "CheckExistDB: Find file=" & strPathFileNameDevDB & " (FileSize: " & lngFileDBSize & " bytes)"
    
                If lngFileDBSize > 0 Then
                    If CompareDevDBVersion(strPathFileNameDevDB) Then
                        CheckExistDB = True
                    Else
                        CheckExistDB = False
                        mbNotSupportedDevDB = True
                    End If
                Else
                    DebugMode str5VbTab & "CheckExistDB: File is zero = 0 bytes"
                End If
            Else
                DebugMode str5VbTab & "CheckExistDB: NOT FIND DB FILE=" & strPathFileNameDevDB
            End If
        Else
            DebugMode str5VbTab & "CheckExistDB: File is zero = 0 bytes"
        End If

    Else
        DebugMode str5VbTab & "CheckExistDB: NOT FIND DB FILE=" & strPathFileNameDevDBHwid
    End If

    DebugMode str4VbTab & "CheckExistDB-End"

End Function

Private Sub CheckMenuUtilsPath()

    If mbIsWin64 Then
        If PathExists(PathCollect(strDevManView_Path64)) = False Then
            mnuUtils_DevManView.Visible = False
        End If

        If PathExists(PathCollect(strSIV_Path64)) = False Then
            mnuUtils_SIV.Visible = False
            lblOsInfo.MousePointer = 0
            lblOsInfo.ToolTipText = vbNullString
        End If

    Else

        If PathExists(PathCollect(strDevManView_Path)) = False Then
            mnuUtils_DevManView.Visible = False
        End If

        If PathExists(PathCollect(strSIV_Path)) = False Then
            mnuUtils_SIV.Visible = False
            lblOsInfo.MousePointer = 0
            lblOsInfo.ToolTipText = vbNullString
        End If
    End If

    If PathExists(PathCollect(strDoubleDriver_Path)) = False Then
        mnuUtils_DoubleDriver.Visible = False
    End If

    If PathExists(PathCollect(strUDI_Path)) = False Then
        mnuUtils_UDI.Visible = False
    End If

    If PathExists(PathCollect(strUnknownDevices_Path)) = False Then
        mnuUtils_UnknownDevices.Visible = False
    End If

End Sub

Private Sub chkPackFiles_Click(Index As Integer)

Dim lngCheckCount                       As Long

    chkPackFiles(Index).Value = chkPackFiles(Index).Value
    lngCheckCount = FindCheckCount

    If lngCheckCount > 0 Then
        ChangeStatusTextAndDebug strMessages(104) & " " & lngCheckCount, , , False
    Else
        ChangeStatusTextAndDebug strMessages(105), , , False
    End If

    chkPackFiles(Index).Refresh

End Sub

Private Sub cmbCheckButton_Click()

Dim strTextforCheck                     As String

    strTextforCheck = cmbCheckButton.Text
    If StrComp(strTextforCheck, strCmbChkBtnListElement2, vbTextCompare) = 0 Then
        LoadIconImage2BtnJC cmdCheck, "BTN_UNCHECKMARK", strPathImageMainWork
    Else
        LoadIconImage2BtnJC cmdCheck, "BTN_CHECKMARK", strPathImageMainWork
    End If

End Sub

Private Sub cmdBreakUpdateDB_Click()
    mbBreakUpdateDBAll = True

End Sub

Private Sub cmdCheck_Click()

Dim i                                   As Long
Dim strTextforCheck                     As String
Dim lngCntBtnTab                        As Long
Dim lngCntBtnPrevious                   As Long
Dim lngCheckCount                       As Long
Dim lngSSTab1Tab                        As Long

    DebugMode "cmdCheck_Click-Start"
    CheckAllButton False
    strTextforCheck = cmbCheckButton.Text

    If strTextforCheck <> strCmbChkBtnListElement3 Then
        If strTextforCheck <> strCmbChkBtnListElement2 Then

            With SSTab1

                lngSSTab1Tab = .Tab

                If lngSSTab1Tab > 0 Then
                    lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 1).CntBtn

                    If lngCntBtnPrevious = 0 Then
                        If lngSSTab1Tab > 1 Then
                            lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 2).CntBtn
                        End If
                    End If
                End If
            End With

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

                        If Not (.Picture Is Nothing) Then
                            If .Visible Then
                                If .Left > 0 Then
                                    chkPackFiles(i).Value = True
                                End If
                            End If
                        End If
                    End With
                End If

            Next

            '"Все новые"
        Case strCmbChkBtnListElement4

            For i = acmdPackFiles.LBound To acmdPackFiles.UBound

                If Not (acmdPackFiles(i).Picture Is Nothing) Then
                    If acmdPackFiles(i).Picture = imgNoDB.Picture Then
                        If acmdPackFiles(i).Visible Then
                            chkPackFiles(i).Value = True
                        End If
                    End If
                End If

            Next

            '"Неустановленные"
        Case strCmbChkBtnListElement5

            For i = lngCntBtnPrevious To lngCntBtnTab

                With acmdPackFiles(i)

                    If .Left > 0 Then
                        If Not (.Picture Is Nothing) Then
                            If .Picture = imgOkAttention.Picture Then
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

                    If Not (.Picture Is Nothing) Then
                        If .Left > 0 Then
                            If .Visible Then
                                chkPackFiles(i).Value = True
                            End If

                            If .Picture = imgNo.Picture Then
                                chkPackFiles(i).Value = False
                            End If

                            If .Picture = imgNoDB.Picture Then
                                chkPackFiles(i).Value = False
                            End If

                            If .Picture = imgOK.Picture Then
                                chkPackFiles(i).Value = False
                            End If

                            If mbCompareDrvVerByDate Then
                                If .Picture = imgOkOld.Picture Then
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

    If lngCheckCount > 0 Then
        ChangeStatusTextAndDebug strMessages(104) & " " & lngCheckCount
    Else
        ChangeStatusTextAndDebug strMessages(105)
    End If

    DebugMode "cmdCheck_Click-End"

End Sub

Private Sub cmdRunTask_Click()
    mbooSelectInstall = False
    BaseUpdateOrRunTask False, True
    BlockControl True
    cmdRunTask.Enabled = FindCheckCount(False)
End Sub

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

Private Function CompatibleDriver4OS(ByVal strSection As String, _
                                     ByVal strDPFileName As String, _
                                     ByVal strDPInfPath As String, _
                                     ByVal strSectionUnsupported As String) As Boolean

Dim mbOSx64                             As Boolean
Dim strOsVer                            As String
Dim strDRVx64                           As String
Dim lngDRVx64                           As Long
Dim strDRVOSVer                         As String
Dim objRegExp                           As RegExp
Dim objMatch                            As Match
Dim objMatches                          As MatchCollection
Dim mbCompatibleByArch                  As Boolean
Dim mbCompatibleByVer                   As Boolean
Dim mbVerFromSection                    As Boolean
Dim mbArchFromSection                   As Boolean
Dim mbVerFromMarkers                    As Boolean
Dim mbArchFromMarkers                   As Boolean
Dim mbVerFromDPName                     As Boolean
Dim mbArchFromDPName                    As Boolean
Dim strRegExpMarkerPattern              As String
Dim mbMarkerCheckExist                  As Boolean
Dim mbMarkerSTRICTCheckExist            As Boolean
Dim strSection_x()                      As String
Dim strSectionMain                      As String
Dim strSectionUnsupportedTemp           As String
Dim mbMarkerFORCEDCheckExist            As Boolean
Dim strDRVOSVerUnsupported              As String

    mbOSx64 = mbIsWin64

    If Not mbSearchCompatibleDriverOtherOS Then
        strOsVer = arrOSList(SSTab1.Tab).Ver
    Else
        strOsVer = strOsCurrentVersion
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
    mbMarkerCheckExist = CheckExistbyRegExp(strDPInfPath, strVer_All_Known_Ver)
    ' проверяем есть ли вообще маркер FORCED в пути
    'strDPInfPath = "5x86\FORCED\M\N\"
    mbMarkerFORCEDCheckExist = CheckExistbyRegExp(strDPInfPath, strVerFORCED & vbBackslashDouble)
    ' проверяем есть ли вообще маркер STRICT в пути
    'strDPInfPath = "5x86\STRICT\M\N\"
    mbMarkerSTRICTCheckExist = CheckExistbyRegExp(strDPInfPath, strVerSTRICT & vbBackslashDouble)

    ' Если нет маркера FORCED, то получаем данные из секции
    'Debug.Print strDPInfPath
    'Debug.Print mbMarkerFORCEDCheckExist
    If Not mbMarkerFORCEDCheckExist Then
        Set objRegExp = New RegExp

        With objRegExp
            .Pattern = "\.NT(X86|AMD64|IA64|)(?:\.(\d(?:.\d)))?"
            .IgnoreCase = True
            'strSection = "AMD.NTAMD64.5.1.1"
            Set objMatches = .Execute(strSection)
        End With

        'Получаем значения версии ОС драйвера и архитектуры
        With objMatches

            'данные берем из секции Manufactured
            If .Count > 0 Then
                Set objMatch = .Item(0)
                strDRVx64 = UCase$(Trim$(objMatch.SubMatches(0)))
                strDRVOSVer = UCase$(Trim$(objMatch.SubMatches(1)))
                lngDRVx64 = InStr(strDRVx64, "64")
            End If
        End With
    Else
        DebugMode str5VbTab & "CompatibleDriver4OS: Check Inf-Section: " & strSection & " Inf-Path: " & strDPInfPath & " contained FORCED marker, section [Manufactured] not analyzing"
    End If

    ' Если секция не имеет окончания вида .NTX86.6.0 - т.е .Count=0, то тогда мы не можем определить точно подходит или нет.
    ' Считаем раз лежит в данной папке, то подходит.
    ' Если в секции manufactured не указано ля каких систем драйвер подходит, то анализируем имя файла
    ' !!! После сделаю поддержку маркеров

    ' Если версия не определена, определяем версию по маркерам или по имени DP
    If LenB(strDRVOSVer) = 0 Then

CheckVerByMarkers:

        Select Case strOsCurrentVersion

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
        mbVerFromMarkers = CheckExistbyRegExp(strDPInfPath, strRegExpMarkerPattern)

        If mbVerFromMarkers Then
            strDRVOSVer = strOsCurrentVersion
        Else

            ' Если по маркерам определить нельзя, определяем версию по имени DP
            If mbMatchHWIDbyDPName Then
                If Not mbMarkerCheckExist Then
                    If InStr(strDPFileName, "WXP") Or InStr(strDPFileName, "WNT5") Then
                        strDRVOSVer = "5.0;5.1;5.2"
                    ElseIf InStr(strDPFileName, "WNT6") Then
                        strDRVOSVer = "6.0;6.1;6.2;6.3"
                    Else

                        If mbOSx64 Then
                            If InStr(strDRVx64, "AMD64") Then
                                strDRVOSVer = strOsCurrentVersion
                            End If

                        Else

                            If InStr(strDRVx64, "X86") Then
                                strDRVOSVer = strOsCurrentVersion
                            End If
                        End If
                    End If
                End If
            End If
        End If

    Else

        If mbMarkerSTRICTCheckExist Then
            DebugMode str5VbTab & "CompatibleDriver4OS: Check Inf-Section: " & strSection & " Inf-Path: " & strDPInfPath & " contained STRICT marker, section [Manufactured] not analyzing by Version"
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
        mbArchFromMarkers = CheckExistbyRegExp(strDPInfPath, strRegExpMarkerPattern, True, strDRVx64)

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
    If LenB(strDRVOSVer) > 0 Then
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

            If strOsCurrentVersion <> "5.0" Then

                Set objRegExp = New RegExp

                With objRegExp
                
                    If mbOSx64 Then
                        .Pattern = "\.NT[AMD64|IA64]*"
                    Else
                        .Pattern = "\.NT[X86]*"
                    End If

                    .Pattern = strSectionMain & .Pattern & "(?:\.(\d(?:.\d)*)*)*,"
                    'ATHEROS\.NT[AMD64|IA64]*(?:\.(\d(?:.\d)*)*)*,
                    'Debug.Print .Pattern
                    .IgnoreCase = True
                    'strSection = "AMD.NTAMD64.5.1.1"
                    '.Pattern = "Atheros,Atheros.NT.6.0,Atheros.NTamd64.6.0"
                    Set objMatches = .Execute(strSectionUnsupportedTemp)
                End With

                'Если несовместиые секции найдены
                With objMatches

                    'данные берем из секции Manufactured
                    If .Count > 0 Then
                        Set objMatch = .Item(0)
                        strDRVOSVerUnsupported = Trim$(objMatch.SubMatches(0))

                        'strDRVOSVer = Trim$(objMatch.SubMatches(1))
                        If LenB(strDRVOSVerUnsupported) > 0 Then

                            ' Если в inf неподдерживаемые секции с версией например 6.0, то неподдериваются ос 6.0 и выше
                            ' т.е если текущая ОС больше чем найденная в inf пустая секция, т.е драйвер не поддерживается
                            If CompareByVersion(strOsVer, strDRVOSVerUnsupported) = ">" Or CompareByVersion(strOsVer, strDRVOSVerUnsupported) = "=" Then

                                ' Если в inf неподдерживаемые секции с версией например 6.0, а драйвер найден в секции 6.1, то драйвер найден правильно, иначе
                                If CompareByVersion(strDRVOSVerUnsupported, strDRVOSVer) = ">" Then
                                    DebugMode str5VbTab & "CompatibleDriver4OS: Check Inf-Section: " & strSection & " Result: " & CompatibleDriver4OS & " by SectionUnsupported:" & strSectionUnsupported, 1
                                    mbCompatibleByArch = False
                                    mbCompatibleByVer = False
                                End If
                            End If
                        End If
                    End If
                End With
            Else
                If UBound(strSection_x) < 1 Then
                    DebugMode str5VbTab & "CompatibleDriver4OS: verOS=" & strOsCurrentVersion & " Check Inf-Section: " & strSection & " Result: " & CompatibleDriver4OS & " by SectionUnsupported:" & strSectionUnsupported, 1
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
    Set objRegExp = Nothing
    Set objMatch = Nothing
    Set objMatches = Nothing
    DebugMode str5VbTab & "CompatibleDriver4OS: Check Inf-Section: " & strSection & " Result: " & CompatibleDriver4OS & " (by Version-" & mbCompatibleByVer & "; by Architecture-" & mbCompatibleByArch & "; by ManufacturedSection:Ver/Arch-" & mbVerFromSection & "/" & mbArchFromSection & "; by Markers:Ver/Arch-" & mbVerFromMarkers & "/" & mbArchFromMarkers & ")", 1

End Function

' проверяем совместимость драйвера по вендору ноутбука
Private Function CheckDRVbyNotebookVendor(ByVal strInfPath As String) As Boolean
Dim i                                   As Long
Dim ii                                  As Long
Dim strFilterList                       As String
Dim strFilterList_x()                   As String
Dim mbFind                              As Boolean

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

' Изменяем имя пакета драйверов (удаление лишних Символов)
Private Function ConvertDPName(ByVal strButtonName As String) As String
Dim strButtonNameTemp                   As String

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
        If InStr(strButtonNameTemp, "_") Then strButtonName = Replace$(strButtonName, "_", " ")
        If InStr(strButtonNameTemp, "-") Then strButtonName = Replace$(strButtonName, "-", " ")
        If InStr(strButtonName, "   ") Then strButtonName = Replace$(strButtonName, "   ", " ")
        If InStr(strButtonName, "  ") Then strButtonName = Replace$(strButtonName, "  ", " ")
        strButtonName = Trim$(strButtonName)
    End If

    ' все в верхний регистр
    If mbButtonTextUpCase Then
        strButtonName = UCase$(strButtonName)
    End If

    ConvertDPName = strButtonName

End Function

'! -----------------------------------------------------------
'!  Функция     :  CreateButtonsonSSTab
'!  Переменные  :  strDrpPath As String, strDevDBPath As String, miTabIndex As Long
'!  Описание    :  Создание кнопок на выбранной вкладке табконтрола
'! -----------------------------------------------------------
Private Sub CreateButtonsonSSTab(ByVal strDrpPath As String, _
                                 ByVal strDevDBPath As String, _
                                 ByVal miTabIndex As Long, _
                                 ByVal lngProgressDelta As Long)

Dim strButtonName                       As String
Dim strPackFileName                     As String
Dim StartPositionLeft                   As Long
Dim StartPositionTop                    As Long
Dim NextPositionLeft                    As Long
Dim NextPositionTop                     As Long
Dim MaxLeftPosition                     As Long
Dim DeltaPositionLeft                   As Long
Dim DeltaPositionTop                    As Long
Dim mbStep                              As Boolean
Dim tabN                                As Long
Dim TabHeight                           As Long
Dim ii                                  As Long
Dim strFileList_x()                     As String
Dim miOffSideCountTemp                  As Long
Dim strPhysXPath                        As String
Dim strLangPath                         As String
Dim strRuntimes                         As String
Dim lngFileCount                        As Long
Dim lngProgressDeltaTemp                As Single

    On Error Resume Next

    DebugMode vbTab & "CreateButtonsonSSTab-Start"
    DebugMode str2VbTab & "CreateButtonsonSSTab: miTabIndex=" & miTabIndex

    If PathExists(strDrpPath) Then
        tabN = miTabIndex
        TabHeight = SSTab1.Height
        Sleep 200
        DoEvents
        SSTab1.Tab = tabN
        StartPositionLeft = miButtonLeft
        StartPositionTop = miButtonTop

        If tabN > 0 Then
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
        strFileListInFolder = vbNullString
        DebugMode str2VbTab & "CreateButtonsonSSTab: Recursion: " & mbRecursion
        DebugMode str2VbTab & "CreateButtonsonSSTab: Get ListFile in folder: " & strDrpPath

        'Строим список файлов 7z
        If Not mbDP_Is_aFolder Then
            strFileList_x = SearchFilesInRoot(strDrpPath, "DP*.7z;DP*.zip", mbRecursion, False, False, True)
            'Иначе это каталоги, а не 7z
        Else
            If FolderContainsSubfolders(strDrpPath) Then
                strFileList_x = SearchFoldersInRoot(strDrpPath, "DP*", False, False)
            End If
        End If

        DebugMode str2VbTab & "CreateButtonsonSSTab: FileCount: " & UBound(strFileList_x, 2)

        If UBound(strFileList_x, 2) = 0 Then
            If LenB(strFileList_x(0, 0)) = 0 Then
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

        strPhysXPath = FileNameFromPath(arrOSList(tabN).PathPhysX)
        strLangPath = FileNameFromPath(arrOSList(tabN).PathLanguages)
        strRuntimes = FileNameFromPath(arrOSList(tabN).PathRuntimes)
        strExcludeFileName = arrOSList(tabN).ExcludeFileName
        lngFileCount = UBound(strFileList_x, 2) - LBound(strFileList_x, 2) + 1
        pbProgressBar.Refresh

        For ii = LBound(strFileList_x, 2) To UBound(strFileList_x, 2)
            strPackFileName = Replace$(strFileList_x(0, ii), BackslashAdd2Path(strDrpPath), vbNullString, , , vbTextCompare)
            DebugMode "===================================================================================================="
            DebugMode str2VbTab & "Work with File: " & strPackFileName
            ChangeStatusTextAndDebug strMessages(69) & " " & strDrpPath & " " & vbNewLine & strMessages(70) & "(" & (ii + 1) & " " & strMessages(124) & " " & lngFileCount & "):" & strPackFileName
            mbStatusHwid = True

            If Not mbDP_Is_aFolder Then
                strButtonName = FileNameFromPath(strPackFileName)
            Else
                strButtonName = strPackFileName
            End If

            ' проверяем что файл не является дополнительным для обработки вместе с архивом
            If MatchSpec(strButtonName, strPhysXPath) Or MatchSpec(strButtonName, strLangPath) Or MatchSpec(strButtonName, strRuntimes) Or MatchSpec(strButtonName, strExcludeFileName) Then
                GoTo NextFiles
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
                    DeltaPositionLeft = acmdPackFiles(lngCntBtn - 1).Left + miButtonWidth + miBtn2BtnLeft - StartPositionLeft
                    NextPositionLeft = StartPositionLeft + DeltaPositionLeft

                    ' Если Кол-во ОС больше кол-ва вкладок на строку
                    If lngOSCount > lngOSCountPerRow Then
                        MaxLeftPosition = NextPositionLeft + miButtonWidth + 100 * (Abs(lngOSCount / lngOSCountPerRow) - 1)
                    Else
                        MaxLeftPosition = NextPositionLeft + miButtonWidth + 25
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
                        DeltaPositionTop = DeltaPositionTop + miButtonHeight + miBtn2BtnTop
                        NextPositionLeft = StartPositionLeft
                        NextPositionTop = StartPositionTop + DeltaPositionTop

                        If NextPositionTop > TabHeight Then
                            mbOffSideButton = True
                            miOffSideCountTemp = miOffSideCountTemp + 1
                        End If

                        mbStep = False
                    End If
                End If
            End If

            ' Загружаем кнопку и чекбокс
            If lngCntBtn > 0 Then
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
                .Top = NextPositionTop + (miButtonHeight - .Height) / 2
                .ZOrder 0
                .Tag = tabN
            End With

            'Считываем подменяемое имя пакета из файла
            EditOrReadDPName lngCntBtn, True

            ' массив HWID для будущего использования для каждой кнопки
            ReDim Preserve arrDevIDs(acmdPackFiles.UBound) As String
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

    If tabN > 0 Then
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
    DebugMode str2VbTab & "CreateButtonsonSSTab: cntButton=" & lngCntBtn

    If miOffSideCountTemp > miOffSideCount Then
        miOffSideCount = miOffSideCountTemp
    End If

    On Error GoTo 0

    DebugMode vbTab & "CreateButtonsonSSTab-End"

End Sub

'! -----------------------------------------------------------
'!  Функция     :  CreateMenuIndex
'!  Переменные  :  Name As String
'!  Описание    :
'! -----------------------------------------------------------
Private Sub CreateMenuIndex(ByVal strName As String)

Dim i                                   As Long

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

'! -----------------------------------------------------------
'!  Функция     :  CreateMenuDevIDIndexDelMenu
'!  Переменные  :  strDevID As String
'!  Описание    :
'! -----------------------------------------------------------
Private Sub CreateMenuDevIDIndexDelMenu(ByVal strDevID As String)

Dim i                                   As Long
Dim ii                                  As Long
Dim DevId_x()                           As String
Dim strName                             As String

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

'! -----------------------------------------------------------
'!  Функция     :  CreateMenuDevIDIndexCopyMenu
'!  Переменные  :  strDevID As String
'!  Описание    :
'! -----------------------------------------------------------
Private Sub CreateMenuDevIDIndexCopyMenu(ByVal strDevID As String)

Dim i                                   As Long
Dim ii                                  As Long
Dim DevId_x()                           As String
Dim strName                             As String

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

'! -----------------------------------------------------------
'!  Функция     :  CreateMenuLngIndex
'!  Переменные  :  Name As String
'!  Описание    :
'! -----------------------------------------------------------
Private Sub CreateMenuLngIndex(ByVal strName As String)

Dim i                                   As Long

    On Error Resume Next

    If Not mnuLang(0).Visible Then
        'если меню еще не создано
        mnuLang(0).Visible = True
        mnuLang(0).Caption = strName
    Else
        Load mnuLang(mnuLang.Count)
        mnuLang(mnuLang.Count - 1).Visible = True

        For i = mnuLang.UBound To mnuLang.LBound Step -1

            If i = mnuLang.LBound Then
                mnuLang(0).Caption = strName
                Exit For
            End If

            mnuLang(i).Caption = mnuLang(i - 1).Caption
        Next
    End If

    On Error GoTo 0

End Sub

'! -----------------------------------------------------------
'!  Функция     :  CreateProgressBar
'!  Переменные  :
'!  Описание    :  Создаем элемент ProgressBar
'! -----------------------------------------------------------
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

'! -----------------------------------------------------------
'!  Функция     :  DeleteUnUsedBase
'!  Переменные  :
'!  Описание    :  Очистка лишних файлов БД
'! -----------------------------------------------------------
Private Sub DeleteUnUsedBase(Optional mbSilent As Boolean = False)

Dim TabCount                            As Long
Dim i                                   As Integer
Dim ii                                  As Integer
Dim strPathDRP                          As String
Dim strPathDevDB                        As String
Dim strFileListTXT_x()                  As String
Dim strFileListDRP_x()                  As String
Dim strFileListDBExists                 As String
Dim strFileListDBNotExists              As String
Dim strDRPFilename                      As String
Dim strFileNameDB                       As String
Dim strFileNameDBHwid                   As String
Dim strFileNameDBIni                    As String
Dim lngFileDBVerIniSize                 As Long
Dim strFileDBVerIniPath                 As String
Dim strFileName2Del                     As String

    DebugMode "DeleteUnUsedBase-Start"

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
                strFileListDRP_x = SearchFoldersInRoot(strPathDRP, "DP*", False, False)
            End If
            'Построение списка txt и ini файлов в каталоге БД
            strFileListTXT_x = SearchFilesInRoot(strPathDevDB, "*DP*.txt;*DP*.ini;*DP*.hwid;*DevDBVersions*.ini", False, False)

            ' Проверка на существование БД
            For ii = LBound(strFileListDRP_x, 2) To UBound(strFileListDRP_x, 2)
                strDRPFilename = FileNameFromPath(strFileListDRP_x(0, ii))

                If CheckExistDB(strPathDevDB, strDRPFilename) Then
                    If InStr(1, strDRPFilename, ".zip", vbTextCompare) Then
                        strFileNameDB = strPathDevDB & Replace$(strDRPFilename, ".zip", ".txt", , , vbTextCompare)
                    End If
                    If InStr(1, strDRPFilename, ".7z", vbTextCompare) Then
                        strFileNameDB = strPathDevDB & Replace$(strDRPFilename, ".7z", ".txt", , , vbTextCompare)
                    End If
                    strFileNameDBHwid = Replace$(strFileNameDB, ".txt", ".hwid", , , vbTextCompare)
                    strFileNameDBIni = Replace$(strFileNameDB, ".txt", ".ini", , , vbTextCompare)
                    strFileListDBExists = AppendStr(strFileListDBExists, strFileNameDB & vbTab & strFileNameDBHwid, vbTab)

                    If PathExists(strFileNameDBIni) Then
                        strFileListDBExists = IIf(LenB(strFileListDBExists) > 0, strFileListDBExists & vbTab, vbNullString) & strFileNameDBIni
                    End If
                End If

            Next
            strFileDBVerIniPath = BackslashAdd2Path(strPathDevDB) & "DevDBVersions.ini"
            strFileListDBExists = IIf(LenB(strFileListDBExists) > 0, strFileListDBExists & vbTab, vbNullString) & strFileDBVerIniPath

            'Строим список удаляемых файлов для несуществующих пакетов
            For ii = LBound(strFileListTXT_x, 2) To UBound(strFileListTXT_x, 2)

                If InStr(1, strFileListDBExists, strFileListTXT_x(0, ii), vbTextCompare) = 0 Then
                    If PathExists(strFileListTXT_x(0, ii)) Then
                        strFileListDBNotExists = IIf(LenB(strFileListDBNotExists) > 0, strFileListDBNotExists & vbNewLine, vbNullString) & Replace$(strFileListTXT_x(0, ii), strAppPath, vbNullString, , , vbTextCompare)

                        'Удаление секции о данном пакете из ini-файла
                        'IniDelAllKeyPrivate FileName_woExt(FileNameFromPath(strFileListTXT_x(0, ii))), strFileDBVerIniPath
                    End If
                End If

            Next
        Next

        ' Вывод итогового сообщения со списком удаляемых файлов и запросом на удаление
        If LenB(strFileListDBNotExists) > 0 Then
            ChangeStatusTextAndDebug strMessages(71)

            If ShowMsbBoxForm(strFileListDBNotExists, strMessages(28), strMessages(29)) = vbYes Then
                strFileListTXT_x = Split(strFileListDBNotExists, vbNewLine)
                'удаление файлов для несуществующих пакетов
                For ii = LBound(strFileListTXT_x) To UBound(strFileListTXT_x)
                    strFileName2Del = PathCollect(strFileListTXT_x(ii))
                    If PathExists(strFileName2Del) Then
                        DeleteFiles strFileName2Del
                        'Удаление секции о данном пакете из ini-файла
                        For i = 0 To TabCount - 1
                            strPathDevDB = arrOSList(i).devIDFolderFull
                            strFileDBVerIniPath = PathCombine(strPathDevDB, "DevDBVersions.ini")
                            'если файл DevDBVersions.ini нулевого размера, то удаляем и его
                            lngFileDBVerIniSize = GetFileSizeByPath(strFileDBVerIniPath)
                            If lngFileDBVerIniSize > 0 Then
                                IniDelAllKeyPrivate FileName_woExt(FileNameFromPath(strFileListTXT_x(ii))), strFileDBVerIniPath
                            Else
                                DebugMode str2VbTab & "DeleteUnUsedBase: Delete - file is zero = 0 bytes: " & strFileDBVerIniPath
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

    DebugMode "DeleteUnUsedBase-End"

End Sub

' Вызов функции для показана формы с сообщением вместо стандартного MsgBox
Private Function ShowMsbBoxForm(strMsgDialog As String, _
                                strMsgFrmCaption As String, _
                                strMsgOKCaption As String) As Long
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

Private Sub EditOrReadDPName(ByVal CurButtonIndex As Long, _
                             Optional ByVal mbRead As Boolean = False)

Dim strDRPFilename                      As String
Dim strDPName                           As String
Dim strDPNameOld                        As String
Dim strDPNameMsg                        As String

    DebugMode str2VbTab & "EditOrReadDPName-Start"
    DebugMode str3VbTab & "EditOrReadDPName: CurButtonIndex=" & CurButtonIndex
    'Считываем текущее имя пакета из файла
    strDPName = vbNullString
    strDRPFilename = FileNameFromPath(acmdPackFiles(CurButtonIndex).Tag)
    strDPNameOld = acmdPackFiles(CurButtonIndex).Caption
    strDPName = IniStringPrivate("DPNames", strDRPFilename, strSysIni)

    ' Если такого значения в файле нет, то ничего не добавляем
    If strDPName = "no_key" Then
        strDPName = strDPNameOld
    End If

    If mbRead Then
        If LenB(strDPName) > 0 Then
            If mbButtonTextUpCase Then
                acmdPackFiles(CurButtonIndex).Caption = UCase$(strDPName)
            Else
                acmdPackFiles(CurButtonIndex).Caption = strDPName
            End If

            ChangeStatusTextAndDebug , str3VbTab & "Change Viewed Name: " & strDRPFilename & " on " & strDPName
        End If

    Else

        If mbIsDriveCDRoom Then
            If Not mbSilentRun Then
                MsgBox strMessages(16), vbInformation, strProductName
            End If

        Else
            ChangeStatusTextAndDebug strMessages(74) & " " & strDRPFilename
            strDPName = InputBox(strMessages(75) & " " & strDRPFilename, strMessages(76), strDPName)

            If LenB(strDPName) = 0 Then
                strDPName = vbNullString
            End If

            If StrComp(strDPNameOld, strDPName) <> 0 Then
                IniWriteStrPrivate "DPNames", strDRPFilename, strDPName, strSysIni
                ChangeStatusTextAndDebug strMessages(77) & " " & strDRPFilename

                If LenB(strDPName) = 0 Then
                    If LenB(strDPName) = 0 Then
                        strDPNameMsg = strDPNameOld
                        strDPName = strDPNameOld
                    Else
                        strDPNameMsg = FileNameFromPath(acmdPackFiles(CurButtonIndex).Tag)
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

    DebugMode str2VbTab & "EditOrReadDPName-End"

End Sub

' Разные сообщения если нет поддерживаемых вкладок, или что-то нет так с пакетами
Private Sub EventOnActivateForm()

Dim lngMsgRet                           As Long

    ' Если пакетов нет вообще, то предлагаем пользователю посетить мой сайт
    If StrComp(acmdPackFiles(0).Container.Name, "frTabPanel", vbTextCompare) = 0 Then
        BlockControlInNoDP False

        With lblNoDPInProgram
            Set .Container = SSTab1

            .AutoSize = True
            .Left = 100

            ' Изменяем положение лабел
            Dim cntUnHideTab            As Long
            Dim miValue1                As Long
            Dim sngNum1                 As Single
            Dim SSTabTabHeight          As Long
            SSTabTabHeight = SSTab1.TabHeight
            cntUnHideTab = FindUnHideTab
            If cntUnHideTab > 0 Then
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

Private Function FindAndInstallPanel(ByVal strArcDRPPath As String, _
                                     ByVal strIniPath As String, _
                                     ByVal strSection As String, _
                                     ByVal lngNumberPanel As Long, _
                                     ByVal strWorkPath As String) As Boolean

Dim lngTagFilesCount                    As Long
Dim lngCommandsCount                    As Long
Dim i                                   As Long
Dim strPrefix                           As String
Dim strPrefixTag                        As String
Dim strPrefixCommand                    As String
Dim strTemp                             As String
Dim strDPSROOT                          As String
Dim strOtherFile                        As String
Dim cmdString                           As String

    'Dim strCommands()    As String
    DebugMode "FindAndInstallPanel-Start"
    DebugMode "FindAndInstallPanel: strIniPath=" & strIniPath
    DebugMode "FindAndInstallPanel: strSection=" & strSection
    DebugMode "FindAndInstallPanel: lngNumberPanel=" & lngNumberPanel
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
    If lngTagFilesCount > 0 Then
        If lngTagFilesCount <> 9999 Then

            ' получаем список таговых файлов
            For i = 1 To lngTagFilesCount
                strTemp = IniStringPrivate(strSection, strPrefixTag & i, strIniPath)

                If strTemp = "no_key" Then
                    GoTo ExitWithFalse
                End If

                'если в пути %DPSROOT% то заменям рабочим каталогом
                strTemp = Replace$(strTemp, "%DPSROOT%\", strDPSROOT, , , vbTextCompare)

                ' Если в пути есть переменные окружения, то заменяем их на нормальный путь
                If InStr(strTemp, Percentage) Then
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
            If lngCommandsCount > 0 Then
                If lngCommandsCount <> 9999 Then

                    'ReDim strCommands(lngCommandsCount) As String
                    ' получаем список Комманд на выполнение
                    For i = 1 To lngCommandsCount
                        strTemp = IniStringPrivate(strSection, strPrefixCommand & i, strIniPath)

                        If strTemp = "no_key" Then
                            GoTo NextCommand
                        End If

                        'если в пути %DPSROOT% то заменям рабочим каталогом
                        strTemp = Replace$(strTemp, "%DPSROOT%\", strDPSROOT, , , vbTextCompare)
                        strTemp = Replace$(strTemp, "%DPSTMP%", strWorkTemp, , , vbTextCompare)
                        '%DPSTMP%
                        strTemp = Replace$(strTemp, "%SystemDrive%\devcon.exe", strDevConExePath, , , vbTextCompare)

                        ' Если в пути есть переменные окружения, то заменяем их на нормальный путь
                        If InStr(strTemp, Percentage) Then
                            strTemp = GetEnviron(strTemp, True)
                        End If

                        'strCommands(i) = strTemp
                        cmdString = strTemp
                        ChangeStatusTextAndDebug strMessages(78) & " '" & strSection & "': " & cmdString

                        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
                            If Not mbSilentRun Then
                                MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
                            End If

                            ChangeStatusTextAndDebug strMessages(79) & " " & strSection, "Error on run : " & cmdString
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
    DebugMode "FindAndInstallPanel-End"
    Exit Function
    ' Аварийный выход
ExitWithFalse:
    FindAndInstallPanel = False

End Function

Private Function FindCheckCount(Optional ByVal mbMsgStatus As Boolean = True) As Long

Dim i                                   As Integer
Dim miCount                             As Integer

    For i = acmdPackFiles.LBound To acmdPackFiles.UBound

        If chkPackFiles(i).Value Then
            miCount = miCount + 1
        End If

    Next
    'cmdRunTask.EnabledCtrl = miCount > 0
    cmdRunTask.Enabled = miCount > 0

    If optRezim_Upd.Value Then
        cmdRunTask.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdRunTask", cmdRunTask.Caption)
    Else
        cmdRunTask.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdRunTask1", cmdRunTask.Caption)
    End If

    If mbLoadAppEnd Then
        If optRezim_Upd.Value Then
            ctlUcStatusBar1.PanelText(1) = strMessages(128)
        Else
            ctlUcStatusBar1.PanelText(1) = strMessages(129)
        End If

        If miCount > 0 Then
            cmdRunTask.Caption = cmdRunTask.Caption & " (" & miCount & ")   "
            'ctlJCbutton1.Caption = cmdRunTask.Caption & " (" & miCount & ")"

            If mbMsgStatus Then
                ChangeStatusTextAndDebug strMessages(104) & " " & miCount, , , False
            End If

        Else

            If mbMsgStatus Then
                ChangeStatusTextAndDebug strMessages(105), , , False
            End If
        End If
    End If

    FindCheckCount = miCount

End Function

'! -----------------------------------------------------------
'!  Функция     :  FindHwidInBaseNew
'!  Переменные  :  strPathDB As String
'!  Описание    :  Поиск вхождения Hwids в БД
'! -----------------------------------------------------------
Private Function FindHwidInBaseNew(ByVal strDevDBPath As String, _
                                   ByVal strPackFileName As String, _
                                   ByVal lngButtonIndex As Long) As String

Dim i                                   As Long
Dim ii                                  As Long
Dim iii                                 As Long
Dim lngCnt                              As Long
Dim strFind                             As String
Dim strFindMachID                       As String
Dim strFindCompatIDTemp                 As String
Dim strFindCompatID_x()                 As String
Dim strFindCompatIDFind                 As String
Dim strFile                             As String
Dim objTextFile                         As TextStream
Dim strFileNameDevDB                    As String
Dim strPathFileNameDevDB                As String
Dim strPathFileNameDevDBHwid            As String
Dim strLineAll                          As String
Dim strAll                              As String
Dim strTemp                             As String
Dim strDevID                            As String
Dim strDevIDOrig                        As String
Dim strDevIDOrig_x()                    As String
Dim strDevPath                          As String
Dim strDevInf                           As String
Dim strDevVer                           As String
Dim strDevVerLocal                      As String
Dim strDevStatus                        As String
Dim strDevName                          As String
Dim strSection                          As String
Dim lngMaxLengthRow1                    As Long
Dim lngMaxLengthRow2                    As Long
Dim lngMaxLengthRow3                    As Long
Dim lngMaxLengthRow4                    As Long
Dim lngMaxLengthRow5                    As Long
Dim lngMaxLengthRow6                    As Long
Dim lngMaxLengthRow9                    As Long
Dim lngMaxLengthRow13                   As Long
Dim lngMaxLengthRowAllLine              As Long
Dim strTTipLocalArr()                   As String
Dim lngTTipLocalArrCount                As Long
Dim miMaxCountArr                       As Long
Dim strPriznakSravnenia                 As String
Dim R                                   As Boolean
Dim R2                                  As Boolean
Dim R3                                  As Boolean
Dim strHwidToDel                        As String
Dim strHwidToDelLine                    As String
Dim lngMatchesCount                     As Long
Dim lngBuffer                           As Long
Dim lngBuffer2                          As Long
Dim lngFileStartFromSymbol              As Long
Dim strFileFullText                     As String
Dim strFileFullTextHwid                 As String
Dim lngDriverScore                      As Long
Dim lngDriverScorePrev                  As Long
Dim strSectionUnsupported               As String
Dim strCatFileExists                    As String

Dim TimeScriptRun                       As Long
Dim TimeScriptFinish                    As Long
Dim strFile_x() As String
Dim strFileFull_x() As String
Dim strResult_x() As String
Dim strResultByTab_x() As String

    DebugMode str4VbTab & "FindHwidInBaseNew-Start"
    DebugMode str5VbTab & "FindHwidInBaseNew: strPackFileName=" & strPackFileName

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
            ' Считываем содержимое всего файла в буфер
            Set objTextFile = objFSO.OpenTextFile(strPathFileNameDevDB, ForReading, False, TristateUseDefault)
            strFileFullText = objTextFile.ReadAll()
            objTextFile.Close
            Erase strFileFull_x
            strFileFull_x = Split(strFileFullText, vbNewLine)
            
            Set objTextFile = objFSO.OpenTextFile(strPathFileNameDevDBHwid, ForReading, False, TristateUseDefault)
            strFileFullTextHwid = objTextFile.ReadAll()
            objTextFile.Close
            Erase strFile_x
            strFile_x = Split(strFileFullTextHwid, vbNewLine)
                                
            lngCnt = UBound(arrHwidsLocal)
            miMaxCountArr = 100
            
            ReDim strTTipLocalArr(12, miMaxCountArr) As String
            lngMaxLengthRow1 = lngTableHwidHeader1
            lngMaxLengthRow2 = lngTableHwidHeader2
            lngMaxLengthRow3 = lngTableHwidHeader3
            lngMaxLengthRow4 = lngTableHwidHeader4
            lngMaxLengthRow5 = lngTableHwidHeader5
            lngMaxLengthRow6 = lngTableHwidHeader6
            lngMaxLengthRow9 = lngTableHwidHeader9
            lngMaxLengthRow13 = lngTableHwidHeader13
            maxSizeRowAllLine = 0
            'i = 0

            'Do While i <= lngCnt
            For i = 0 To lngCnt
                strFind = arrHwidsLocal(i).HWIDCutting
                'Debug.Print strFind
                strFindCompatIDTemp = arrHwidsLocal(i).HWIDCompat
                'If InIDE() Then
                   'If InStr(1, strFind, "FORCED\7x64\HP\E1D62x64.INF", vbTextCompare) Then Stop ' Debug .Assert strInfPath
                'End If
                
                ' Быстрый поиск совпадений в массиве
                lngBuffer = BinarySearch(strFile_x(), strFind)
                
                DebugMode str5VbTab & "FindHwidInBaseNew: PreFind by HWID: " & strFind & " =" & lngBuffer, 2
                'Debug.Print str5VbTab & "FindHwidInBaseNew: PreFind by HWID: " & strFind & " =" & lngBuffer
                lngFileStartFromSymbol = lngBuffer

                If lngBuffer < 0 Then
                    ' заносим HWID в индекс, чтобы лишний раз потом его не проверять
                    Set objHashOutput3 = New Scripting.Dictionary
                    objHashOutput3.CompareMode = TextCompare

                    ' подходящие HWID (т.е обозначение HWID просто может быть записано по другому)
                    If mbMatchingHWID Then
                        strFindMachID = arrHwidsLocal(i).HWIDMatches

                        If LenB(strFindMachID) > 0 Then
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
                    End If

                    ' Поиск по совместимым HWID
                    If mbCompatiblesHWID Then
                        If InStr(strFindCompatIDTemp, "UNKNOWN") = 0 Then
                            If LenB(strFindCompatIDTemp) > 0 Then
                                strFindCompatID_x = Split(strFindCompatIDTemp, " | ")
                            End If

                        Else
                            GoTo NextStrFind
                        End If

                        strFind = vbNullString

                        For iii = LBound(strFindCompatID_x) To UBound(strFindCompatID_x)

                            'Глубина поиска HWID
                            If iii > lngCompatiblesHWIDCount Then
                                Exit For
                            End If

                            strFindCompatIDFind = strFindCompatID_x(iii)
                            'Debug.Print iii & " " & lngCompatiblesHWIDCount & " " & strFindCompatIDFind

                            If Not MatchSpec(strFindCompatIDFind, strExcludeHWID) Then
                                R3 = objHashOutput3.Exists(strFindCompatIDFind)

                                If Not R3 Then
                                    objHashOutput3.Item(strFindCompatIDFind) = "+"
                                    lngBuffer2 = 0
                                    lngBuffer2 = BinarySearch(strFile_x(), strFindCompatIDFind)
                                    DebugMode str5VbTab & "FindHwidInBaseNew: ***PreFind by HWID-Compatibles: " & strFindCompatIDFind & " =" & lngBuffer2, 2
                                    'Debug.Print str5VbTab & "FindHwidInBaseNew: ***PreFind by HWID-Compatibles: " & strFindCompatIDFind & " =" & lngBuffer2
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

'
ExitFromForNext_iii:

                If lngFileStartFromSymbol < 0 Then
                    DebugMode str5VbTab & "FindHwidInBaseNew: !!!ERROR lngFileStartFromSymbol=0 " & strPackFileName & vbBackslash & BackslashAdd2Path(strDevPath) & strDevInf & " by HWID=" & strFind, 1
                    GoTo NextStrFind
                End If
                
                Erase strResult_x
                strResult_x = Filter(strFileFull_x(), strFind & vbTab, True, vbBinaryCompare)

                lngMatchesCount = UBound(strResult_x)

                If lngMatchesCount >= 0 Then
                    DebugMode str5VbTab & "FindHwidInBaseNew: !!!Find " & lngMatchesCount & "Match in: " & strPackFileName & vbBackslash & BackslashAdd2Path(strDevPath) & strDevInf & " by HWID=" & strFind, 2
                    'ii = 0

                    'Do While ii <= lngMatchesCount
                    For ii = 0 To lngMatchesCount
                        strResultByTab_x = Split(strResult_x(ii), vbTab)
                        
                        ' Получаем имя секции файла inf для дальнейшего анализа
                        strDevPath = strResultByTab_x(1)
                        strSection = strResultByTab_x(3)
                        ' получение списка секций несовместимых ОС
                        strSectionUnsupported = strResultByTab_x(5)

                        ' Если драйвер несовместим с текущей ОС (вкладкой), то пропускаем его (анализ имени секции manufactured)
                        If Not CompatibleDriver4OS(strSection, strPackFileName, strDevPath, strSectionUnsupported) Then
                            DebugMode str5VbTab & "FindHwidInBaseNew: !!! SKIP. Driver is not compatible for this OS - IniSection: " & strSection, 1
                            GoTo NextlngMatchesCount
                        End If

                        strDevID = strResultByTab_x(0)
                        strDevInf = strResultByTab_x(2)
                        strCatFileExists = strResultByTab_x(6)

                        If mbCalcDriverScore Then
                            ' Проверка и присвоение баллов драйверов
                            ' Если до этого оценок не было, то добавляем в базу
                            DebugMode str5VbTab & "FindHwidInBaseNew: ***Driver find in : " & PathCombine(strPackFileName & vbBackslash & strDevPath, strDevInf) & " Has Score=" & lngDriverScore, 1

                            If arrHwidsLocal(i).DRVScore = 0 Then
                                arrHwidsLocal(i).DRVScore = lngDriverScore
                            Else
                                lngDriverScorePrev = arrHwidsLocal(i).DRVScore

                                If lngDriverScore > lngDriverScorePrev Then
                                    DebugMode str5VbTab & "FindHwidInBaseNew: ***Driver is WORSE than found previously: ScoredPrev=" & lngDriverScorePrev, 1
                                    GoTo NextlngMatchesCount
                                Else
                                    arrHwidsLocal(i).DRVScore = lngDriverScore
                                    DebugMode str5VbTab & "FindHwidInBaseNew: ***Added! Driver is BETTER OR EQUAL than found previously: ScoredPrev=" & lngDriverScorePrev, 1
                                End If
                            End If
                        End If

                        strDevVer = strResultByTab_x(4)

                        ' если необходимо конвертировать дату в формат dd/mm/yyyy
                        If mbDateFormatRus Then
                            ConvertVerByDate strDevVer
                        End If

                        strDevVerLocal = arrHwidsLocal(i).VerLocal

                        If LenB(strDevVerLocal) = 0 Then
                            strDevVerLocal = "unknown"
                        End If

                        strDevName = strResultByTab_x(7)

                        If arrHwidsLocal(i).Status = 0 Then
                            mbStatusHwid = False
                            If InStr(strDevVerLocal, "unknown") = 0 Then
                                If LenB(strDevVerLocal) > 0 Then
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

                            Select Case strPriznakSravnenia
                                    ' В БД новее
                                Case ">"
                                    mbStatusNewer = True
                                    mbStatusOlder = False

                                    ' В БД старее
                                Case "<"

                                    If Not mbStatusOlder Then
                                        If Not mbStatusNewer Then
                                            mbStatusOlder = True
                                        End If
                                    End If

                                    ' Дрова равны
                            End Select

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
                        arrHwidsLocal(i).DRVExist = 1
                        arrHwidsLocal(i).DPsList = AppendStr(arrHwidsLocal(i).DPsList, strPackFileName, " | ")

                        ' Если записей в массиве становится больше чем объявлено, то увеличиваем размерность массива
                        If lngTTipLocalArrCount = miMaxCountArr Then
                            miMaxCountArr = 2 * miMaxCountArr
                            ReDim Preserve strTTipLocalArr(12, miMaxCountArr)
                        End If

                        ' Заносим данные во временный массив
                        strTTipLocalArr(0, lngTTipLocalArrCount) = strDevID
                        strTTipLocalArr(1, lngTTipLocalArrCount) = strDevPath
                        strTTipLocalArr(2, lngTTipLocalArrCount) = strDevInf
                        strTTipLocalArr(3, lngTTipLocalArrCount) = strDevVer
                        strTTipLocalArr(4, lngTTipLocalArrCount) = strDevVerLocal
                        strTTipLocalArr(5, lngTTipLocalArrCount) = strDevStatus
                        strTTipLocalArr(6, lngTTipLocalArrCount) = strDevName
                        strTTipLocalArr(7, lngTTipLocalArrCount) = strPriznakSravnenia
                        strTTipLocalArr(8, lngTTipLocalArrCount) = strSection
                        strTTipLocalArr(9, lngTTipLocalArrCount) = strDevIDOrig
                        strTTipLocalArr(10, lngTTipLocalArrCount) = lngDriverScore
                        strTTipLocalArr(11, lngTTipLocalArrCount) = strSectionUnsupported
                        strTTipLocalArr(12, lngTTipLocalArrCount) = strCatFileExists
                        lngTTipLocalArrCount = lngTTipLocalArrCount + 1

                        If mbFirstStart Then
                            ' Заносим данные в глобальный массив массив
                            ' Если записей в массиве становится больше чем объявлено, то увеличиваем размерность массива
                            If lngDriversArrCount = lngMaxDriversArrCount Then
                                lngMaxDriversArrCount = 2 * lngMaxDriversArrCount
                                ReDim Preserve arrDriversList(13, lngMaxDriversArrCount)
                            End If
                            arrDriversList(0, lngDriversArrCount) = strDevID
                            arrDriversList(1, lngDriversArrCount) = strDevPath
                            arrDriversList(2, lngDriversArrCount) = strDevInf
                            arrDriversList(3, lngDriversArrCount) = strDevVer
                            arrDriversList(4, lngDriversArrCount) = strDevVerLocal
                            arrDriversList(5, lngDriversArrCount) = strDevStatus
                            arrDriversList(6, lngDriversArrCount) = strDevName
                            arrDriversList(7, lngDriversArrCount) = strPriznakSravnenia
                            arrDriversList(8, lngDriversArrCount) = strSection
                            arrDriversList(9, lngDriversArrCount) = strDevIDOrig
                            arrDriversList(10, lngDriversArrCount) = lngDriverScore
                            arrDriversList(11, lngDriversArrCount) = strSectionUnsupported
                            arrDriversList(12, lngDriversArrCount) = strCatFileExists
                            arrDriversList(13, lngDriversArrCount) = strPackFileName
                            lngDriversArrCount = lngDriversArrCount + 1
                        End If

                        'Устанавливаем ширину колонок в таблице
                        If Len(strDevID) > lngMaxLengthRow1 Then
                            lngMaxLengthRow1 = Len(strDevID)
                        End If

                        If Len(strDevPath) > lngMaxLengthRow2 Then
                            lngMaxLengthRow2 = Len(strDevPath)
                        End If

                        If Len(strDevInf) > lngMaxLengthRow3 Then
                            lngMaxLengthRow3 = Len(strDevInf)
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
NextlngMatchesCount:
                    '    ii = ii + 1
                    'Loop
                    Next ii
                Else
                    DebugMode str5VbTab & "FindHwidInBaseNew: !!!ERROR Driver NOT find by Regexp in : " & strPackFileName & vbBackslash & BackslashAdd2Path(strDevPath) & strDevInf & " by HWID=" & strFind, 2
                End If
NextStrFind:
                'i = i + 1
                Set objHashOutput3 = Nothing
            Next i
            'Loop

            If lngTTipLocalArrCount > 0 Then
                ' Объеявляем индексы, чтобы убрать повторяющие ся строки в подсказках
                Set objHashOutput = New Scripting.Dictionary
                objHashOutput.CompareMode = TextCompare
                Set objHashOutput2 = New Scripting.Dictionary
                objHashOutput2.CompareMode = TextCompare

                ReDim Preserve strTTipLocalArr(12, lngTTipLocalArrCount - 1)
                'i = LBound(strTTipLocalArr, 2)
                'ii = UBound(strTTipLocalArr, 2)

                For i = LBound(strTTipLocalArr, 2) To UBound(strTTipLocalArr, 2)
                'Do While i <= ii
                    'strDevID
                    strTemp = strTTipLocalArr(0, i)
                    strTTipLocalArr(0, i) = strTemp & Space$(lngMaxLengthRow1 - Len(strTemp) + 1) & "| "
                    'strDevPath
                    strTemp = strTTipLocalArr(1, i)
                    strTTipLocalArr(1, i) = strTemp & Space$(lngMaxLengthRow2 - Len(strTemp) + 1) & "| "
                    'strDevInf
                    strTemp = strTTipLocalArr(2, i)
                    strTTipLocalArr(2, i) = strTemp & Space$(lngMaxLengthRow3 - Len(strTemp) + 1) & "| "
                    'strDevVer
                    strTemp = strTTipLocalArr(3, i)
                    strTTipLocalArr(3, i) = strTemp & Space$(lngMaxLengthRow4 - Len(strTemp) + 1) & "| "
                    'strDevVerLocal
                    strTemp = strTTipLocalArr(4, i)
                    strTTipLocalArr(4, i) = strTemp & Space$(lngMaxLengthRow5 - Len(strTemp) + 1) & "| "
                    ' strPriznakSravnenia
                    strTemp = strTTipLocalArr(7, i)
                    strTTipLocalArr(7, i) = strTemp & Space$(lngMaxLengthRow9 - Len(strTemp) + 1) & "| "
                    'strDevStatus & strDevName
                    strTemp = strTTipLocalArr(5, i)
                    strTTipLocalArr(5, i) = strTemp & Space$(lngMaxLengthRow6 - Len(strTemp) + 1) & "| "
                    ' Секция
                    strTemp = strTTipLocalArr(8, i)
                    strTTipLocalArr(8, i) = strTemp & Space$(lngMaxLengthRow13 - Len(strTemp) + 1) & "|"
                    ' Итоговый
                    strLineAll = strTTipLocalArr(0, i) & strTTipLocalArr(1, i) & strTTipLocalArr(2, i) & strTTipLocalArr(3, i) & strTTipLocalArr(7, i) & strTTipLocalArr(4, i) & strTTipLocalArr(5, i) & strTTipLocalArr(6, i)
                    R = objHashOutput.Exists(strLineAll)

                    If Not R Then
                        objHashOutput.Item(strLineAll) = "+"
                        strAll = AppendStr(strAll, strLineAll, vbNewLine)
                    End If

                    ' Заполняем массив для удаления драйверов по HWID
                    strHwidToDelLine = strTTipLocalArr(9, i)
                    R2 = objHashOutput2.Exists(strHwidToDelLine)

                    If Not R2 Then
                        objHashOutput2.Item(strHwidToDelLine) = "+"
                        strHwidToDel = AppendStr(strHwidToDel, strHwidToDelLine & vbTab & strTTipLocalArr(6, i), ";")
                    End If

                    ' Подсчитываем максимальную длину строки в подсказке
                    If Len(strLineAll) > lngMaxLengthRowAllLine Then
                        lngMaxLengthRowAllLine = Len(strLineAll)
                    End If

                    'i = i + 1
                'Loop
                Next i
                
                Set objHashOutput = Nothing
                Set objHashOutput2 = Nothing
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

            lngSizeRow3 = lngMaxLengthRow3
            If lngSizeRow3Max < lngSizeRow3 Then
                lngSizeRow3Max = lngSizeRow3
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
    arrTTipSize(lngButtonIndex) = maxSizeRowAllLine & (";" & lngSizeRow1 & ";" & lngSizeRow2 & ";" & lngSizeRow3 & ";" & lngSizeRow4 & ";" & lngSizeRow9 & ";" & lngSizeRow5 & ";" & lngSizeRow6)

ExitFromSub:
    TimeScriptFinish = GetTickCount
    DebugMode str4VbTab & "FindHwidInBaseNew-Time to Find by HWID - " & strPackFileName & ": " & CalculateTime(TimeScriptRun, TimeScriptFinish, True), 1
    Exit Function
    
End Function

Private Function FindNoDBCount() As Long

Dim miCount                             As Integer
Dim i                                   As Integer

    For i = acmdPackFiles.LBound To acmdPackFiles.UBound

        If Not (acmdPackFiles(i).Picture Is Nothing) Then
            If acmdPackFiles(i).Picture = imgNoDB.Picture Then
                miCount = miCount + 1
            End If
        End If

    Next
    FindNoDBCount = miCount

End Function

Private Function FindUnHideTab() As Integer

Dim miCount                             As Integer
Dim i                                   As Integer

    miCount = -1

    For i = 0 To SSTab1.Tabs - 1

        If SSTab1.TabVisible(i) Then
            miCount = miCount + 1
        End If

    Next
    FindUnHideTab = miCount

End Function

'! -----------------------------------------------------------
'!  Функция     :  Form_Activate
'!  Переменные  :
'!  Описание    :  Событие активации формы
'! -----------------------------------------------------------
Private Sub Form_Activate()

Dim lStart                              As Long
Dim lEnd                                As Long
Dim cntFindUnHideTab                    As Integer

    If mbFirstStart Then
        If mbStartMaximazed Or mbChangeResolution Then
            Me.WindowState = vbMaximized
            DoEvents
        End If

        ' Создаем элемент ProgressBar
        CreateProgressNew

        Sleep 300
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
            DebugMode "Time to Collect INFO from Reestr: =" & CalculateTime(lStart, lEnd, True)

            With ctlProgressBar1
                .Value = 250
                .SetTaskBarProgressValue 250, 1000
            End With
            ChangeFrmMainCaption 250
            
            ChangeStatusTextAndDebug strMessages(80)
        Else
            ChangeStatusTextAndDebug strMessages(57) & vbNewLine & strMessages(4)
            Unload Me
        End If

        ' Назначить имена для вкладок и установить текущую на основании версии ОС
        SetTabsNameAndCurrTab False
        ' Загрузить все кнопки
        LoadButton
        SaveHWIDs2File
        ' Вывести в лог список всех драйверов
        If lngDriversArrCount > 0 Then
            PutAllDrivers2Log
        End If

        SetTabsNameAndCurrTab True
        DoEvents
        BlockControl True
        FindCheckCount
        frTabPanel.Visible = True

        'SSTab1.Visible = True
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
        ChangeStatusTextAndDebug strMessages(59) & " " & dtAllTimeProg, "End Start Operation" & " StartTime is: " & dtAllTimeProg
        ' Иконки меню
        'If mbExMenu Then
        'ExMenuEnable
        'End If
        Me.Refresh
        Sleep 400

        If mbRunWithParam Then
            ChangeStatusTextAndDebug strMessages(60), "Program start in silentMode"
            frmSilent.Show vbModal, Me

            If mbSilentRun Then

                ' Создаем точку восстановления
                If mbCreateRestorePoint Then
                    CreateRestorePoint
                End If

                If Not mbNoSupportedOS Then
                    '"Начинается автоматическая установка"
                    SilentInstall
                    ' после установки закрываем программу
                    Unload Me
                End If

            Else

                ' Создаем точку восстановления
                If mbCreateRestorePoint Then
                    If MsgBox(strMessages(115) & vbNewLine & strMessages(120), vbQuestion + vbYesNo, strProductName) = vbYes Then
                        CreateRestorePoint
                    End If
                End If
            End If

        Else
            ' Разные сообщения если нет поддерживаемых вкладок, или что-то нет так с пакетами
            EventOnActivateForm

            ' Создаем точку восстановления
            If mbCreateRestorePoint Then
                If MsgBox(strMessages(115) & vbNewLine & strMessages(120), vbQuestion + vbYesNo, strProductName) = vbYes Then
                    CreateRestorePoint
                End If
            End If

            ' Проверка обновлений при старте, только если не тихий режим установки
            If mbUpdateCheck Then
                ctlUcStatusBar1.PanelText(1) = strMessages(145)
                ChangeStatusTextAndDebug strMessages(58)
                CheckUpd
            Else
                ShowUpdateToolTip
            End If
        End If

    End If

    mbFirstStart = False
    mbLoadAppEnd = True

End Sub

'! -----------------------------------------------------------
'!  Функция     :  Form_KeyDown
'!  Переменные  :  KeyCode As Integer, Shift As Integer
'!  Описание    :  обработка нажатий клавиш клавиатуры
'! -----------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Выход из программы по "Escape"
    If Not mbFirstStart And KeyCode = vbKeyEscape Then
        If Not mbCheckUpdNotEnd Then
'            If MsgBox(strMessages(34), vbQuestion + vbYesNo, strProductName) = vbYes Then
'                Unload Me
'                Exit Sub
'            End If
        End If
    ' Нажата кнопка "Ctrl"
    ElseIf Shift = 2 Then

        Select Case KeyCode

            Case 65
                ' Ctrl+A (Выделение всех пакетов для установки)
                CheckAllButton True

            Case 90
                ' Ctrl+Z (Отмена выделения всех)
                CheckAllButton False

            Case 83
                ' Ctrl+S (Выделение всех пакетов на вкладке)
                SelectAllOnTabDP True

            Case 78
                ' Ctrl+N (Выделение всех пакетов с новыми драйверами)
                SelectRecommendedDP True

            Case 81
                ' Ctrl+Q (Выделение пакетов с не установленными)
                SelectNotInstalledDP True

            Case 73
                ' Ctrl+I (Установка выделенных пакетов)
                InsOrUpdSelectedDP True

            Case 85
                ' Ctrl+U (Обновление БД выделенных пакетов)
                InsOrUpdSelectedDP False

            Case 9

                ' CTRL+Tab (Переключение по вкладкам)
                If SSTab1.Tabs > 0 Then
                    SelectNextTab
                End If

            Case 19

                ' CTRL+Break (Прерывание групповой обработки)
                If cmdBreakUpdateDB.Visible Then
                    cmdBreakUpdateDB_Click
                End If

        End Select
    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  Form_Load
'!  Переменные  :
'!  Описание    :
'! -----------------------------------------------------------
Private Sub Form_Load()

'cmdViewAllDevice.SetPopupMenu mnuContextMenu
'cmdViewAllDevice.SetPopupMenuRBT mnuContextMenu2

Dim i                                   As Long
Dim ii                                  As Long

    DebugMode "MainForm Show"

    SetupVisualStyles Me

    With Me
        ' изменяем иконки формы и приложения
        ' Icon for Exe-file
        SetIcon .hWnd, "APPICON", True
        SetIcon .hWnd, "FRMMAIN", False
        ' Смена заголовка формы
        strFormName = .Name
        ChangeFrmMainCaption
        ' Разворачиваем форму на весь экран
        .Width = MainFormWidth
        .Height = MainFormHeight
        ' Центрируем форму на экране
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With


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
        .PanelWidth(2) = (MainFormWidth \ Screen.TwipsPerPixelX) - .PanelWidth(1)
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
    'SSTab1.Visible = False
    frTabPanel.Visible = False
    mnuContextMenu.Visible = False
    mnuContextMenu2.Visible = False
    mnuContextMenu3.Visible = False
    BlockControl False
    CheckMenuUtilsPath
    frTabPanel.Top = 3100
    frTabPanel.Left = 75
    lblOsInfo.Left = 75

    With acmdPackFiles(0)
        .Left = miButtonLeft
        .Top = miButtonTop
        .Width = miButtonWidth
        .Height = miButtonHeight
        .CheckExist = True
    End With

    With chkPackFiles(0)
        .Width = 200
        .Height = 200
        .Left = miButtonLeft + miButtonWidth - 225
        .Top = miButtonTop + 30
    End With

    ' Устанавливаем шрифт кнопок
    SetButtonProperties acmdPackFiles(0)
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

    DebugMode "LoadTabList"
    DebugMode "TabsPerRow: " & SSTab1.TabsPerRow
    DebugMode "TabsCount: " & SSTab1.Tabs

    ' Загрузка меню утилит
    If arrUtilsList(0, 1) <> "List_Empty" Then
        DebugMode "CreateUtilsList: " & UBound(arrUtilsList)

        For i = UBound(arrUtilsList) To 0 Step -1
            CreateMenuIndex arrUtilsList(i, 0)
        Next
    End If

    ' Загрузка меню языков и локализация приложения
    mnuMainLang.Visible = mbMultiLanguage

    If mbMultiLanguage Then
        DebugMode "CreateLangList: " & UBound(arrLanguage)

        For i = UBound(arrLanguage, 2) To 1 Step -1
            CreateMenuLngIndex arrLanguage(2, i)
        Next
        Localise strPCLangCurrentPath

        For ii = mnuLang.LBound To mnuLang.UBound
            mnuLang(ii).Checked = arrLanguage(1, ii + 1) = strPCLangCurrentPath
        Next
        mnuLangStart.Checked = Not mbAutoLanguage
    End If

    DebugMode "OsInfo: " & lblOsInfo.Caption
    DebugMode "PCModel: " & lblPCInfo.Caption

    ' Выставляем шрифт
    FontCharsetChange

    ' Изменяем параметры Всплывающей подсказки для кнопок
    With TT
        .Font.Name = strMainForm_FontName
        .Font.Size = lngMainForm_FontSize
        .MaxTipWidth = Me.Width
        .SetDelayTime TipDelayTimeInitial, 400
        .SetDelayTime TipDelayTimeShow, 15000
        .Title = strTTipTextTitle
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
    cmdRunTask.Enabled = False
    'загрузка меню кнопки CmdRunTask
    LoadCmdRunTask

    'заполнение списка на выделение
    LoadListChecked
    mbFirstStart = True

    If mbIsWin64 Then
        If PathExists(PathCollect("Tools\SIV\SIV64X.exe")) Then
            lblOsInfo.ToolTipText = "View system info using System Information Viewer"
        End If

    Else

        If PathExists(PathCollect("Tools\SIV\SIV32X.exe")) Then
            lblOsInfo.ToolTipText = "View system info using System Information Viewer"
        End If
    End If

    mnuAutoInfoAfterDelDRV.Checked = mbAutoInfoAfterDelDRV

End Sub

'! -----------------------------------------------------------
'!  Функция     :  Form_QueryUnload
'!  Переменные  :  Cancel As Integer, UnloadMode As Integer
'!  Описание    :  Корректная выгрузка формы
'! -----------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Проверяем закончена ли проверка обновления, если нет то прерываем выход из программы, иначе программа вылетит
    If mbCheckUpdNotEnd Then
        Cancel = UnloadMode = vbFormControlMenu Or vbFormCode
        Exit Sub
    End If

End Sub

Public Sub UnloadAllForms(Optional FormToIgnore As String = vbNullString)
Dim F                                   As Form

    For Each F In Forms
        If Not F Is Nothing Then
            If StrComp(F.Name, FormToIgnore, vbTextCompare) <> 0 Then
                Unload F
                Set F = Nothing
            End If
        End If
    Next F
End Sub

'! -----------------------------------------------------------
'!  Функция     :  Form_Resize
'!  Переменные  :
'!  Описание    :  Изменение размеров формы
'! -----------------------------------------------------------
Public Sub Form_Resize()

Dim OptWidth                            As Long
Dim OptWidthDelta                       As Long
Dim ImgWidth                            As Long
Dim imgWidthDelta                       As Long
Dim miDeltafrmMainWidth                 As Long
Dim miDeltafrmMainHeight                As Long
Dim cntFindUnHideTab                    As Integer

    On Error Resume Next

    ' если форма не свернута, то изменяем размеры
    If Me.WindowState <> vbMinimized Then

        ' если форма не максимизирована, то изменяем размеры формы
        If OsCurrVersionStruct.VerFull >= "6.0" Then
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

        If Me.Width < MainFormWidthMin Then
            Me.Width = MainFormWidthMin
            Me.Enabled = False
            Me.Enabled = True
            Exit Sub
        End If

        If Me.Height < MainFormHeightMin Then
            Me.Height = MainFormHeightMin
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

            With SSTab1
                .Height = frTabPanel.Height - 20
                .Width = frTabPanel.Width - 20
            End With

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
        lblOsInfo.Width = frInfo.Width - 200
        lblPCInfo.Width = frInfo.Width - 200
        cmdViewAllDevice.Width = optRezim_Upd.Left + optRezim_Upd.Width - cmdViewAllDevice.Left
        ' Удаление иконки в трее если есть
        SetTrayIcon NIM_DELETE, Me.hWnd, 0&, vbNullString

        With lblNoDPInProgram
            .Left = 100

            ' Изменяем положение лабел
            Dim cntUnHideTab            As Long
            Dim miValue1                As Long
            Dim sngNum1                 As Single
            Dim SSTabTabHeight          As Long
            SSTabTabHeight = SSTab1.TabHeight
            cntUnHideTab = FindUnHideTab
            If cntUnHideTab > 0 Then
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

Private Sub Form_Unload(Cancel As Integer)
' Проверяем закончена ли проверка обновления, если нет то прерываем выход из программы, иначе программа вылетит
    If mbCheckUpdNotEnd Then
        Cancel = True
        Exit Sub
    End If

    ' Выгружаем из памяти формы
    UnloadAllForms strFormName

    ' Удаление временных файлов если есть и если опция включена
    If mbDelTmpAfterClose Then
        ChangeStatusTextAndDebug strMessages(81), , , , strMessages(130)
        'Чистим если только не перезапуск программы
        If Not mbRestartProgram Then
            Me.Hide
            DelTemp
        End If
    End If

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

        If StrComp(FileNameFromPath(strSysIni), "Settings_DIA_TMP.ini", vbTextCompare) = 0 Then
            DeleteFiles strSysIni
        End If
    End If

    ' Выгружаем из памяти другие компоненты
    'Unload Me
    'Set frmMain = Nothing
    
End Sub

Private Sub frDescriptionIco_MouseMove(Button As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)

    If Button = vbRightButton Then
        OpenContextMenu Me, Me.mnuContextMenu2
    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  FRMStateSave
'!  Переменные  :
'!  Описание    :  Запись положения форм в ini-шку
'! -----------------------------------------------------------
Private Sub FRMStateSave()

Dim miHeight                            As Long
Dim miWidth                             As Long
Dim miWindowState                       As Long

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

Private Sub GroupInstallDP()

Dim ButtIndex                           As Long
Dim miCheckDPCount                      As Long
Dim i                                   As Long
Dim strPackFileName                     As String
Dim strPathDRP                          As String
Dim strPathDevDB                        As String
Dim strPackFileName_woExt               As String
Dim ArchTempPath                        As String
Dim DPInstExitCode                      As Long
Dim strDevPathShort                     As String
Dim miCheckDPNumber                     As Long
Dim strPhysXPath                        As String
Dim strLangPath                         As String
Dim strRuntimes                         As String
Dim ReadExitCodeString                  As String
Dim miPbInterval                        As Long
Dim miPbNext                            As Long
Dim lngFindCheckCountTemp               As Long

    DebugMode "GroupInstallDP-Start"
    ButtIndex = chkPackFiles.UBound
    miCheckDPCount = FindCheckCount
    BlockControl False

    If miCheckDPCount > 0 Then
        ReDim arrCheckDP(1, miCheckDPCount - 1)

        If ButtIndex > 0 Then
            miCheckDPNumber = 0

            ' Составляем список выделенных пакетов
            For i = 0 To ButtIndex

                ' Если галка стоит на кнопке, то обрабатываем эту кнопку
                If chkPackFiles(i).Value Then
                    If chkPackFiles(i).Left > 0 Then
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
                    If chkPackFiles(i).Left > 0 Then
                        ' Индексы отмеченых кнопок
                        arrCheckDP(0, 0) = 0
                        miCheckDPNumber = 1
                    End If
                End If
            Else
                If Not mbSilentRun Then
                    MsgBox strMessages(12), vbInformation, strProductName
                End If

                DebugMode "GroupInstallDP-DpPack is not Exist"
                Exit Sub
            End If
        Else

            If Not mbSilentRun Then
                MsgBox strMessages(12), vbInformation, strProductName
            End If

            DebugMode "GroupInstallDP-DpPack is not Exist"
            Exit Sub
        End If

        ' Получаем список извлекаемых масок
        ' если выборочная установка, то получаем список каталогов для распаковки
        If mbooSelectInstall Then
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
                ChangeStatusTextAndDebug strMessages(82)
                Exit Sub
            End If

        Else

            ' иначе список строится сам
            For i = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)
                Dim strTemp_x() As String
                Dim strTempLine_x() As String
                Dim i_arr As Long
                
                strPathDRPList = vbNullString
                strTemp_x = Split(arrTTip(arrCheckDP(0, i)), vbNewLine)
                
                For i_arr = LBound(strTemp_x) To UBound(strTemp_x)
                    strTempLine_x = Split(strTemp_x(i_arr), " | ")
    
                    If LenB(Trim$(strTemp_x(i_arr))) Then
                        strDevPathShort = Trim$(strTempLine_x(1))
                        ' Если данного пути нет в списке, то добавляем
                        If InStr(1, strPathDRPList, strDevPathShort, vbTextCompare) = 0 Then
                            strPathDRPList = AppendStr(strPathDRPList, strDevPathShort, " ")
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
                
                If .ShowFolder = True Then
                    ArchTempPath = .FileName
                End If
            End With

            If LenB(ArchTempPath) = 0 Then
                ChangeStatusTextAndDebug strMessages(132)
                '# if user cancel #
                Exit Sub
            End If

            DebugMode "StartBackUp: Destination=" & ArchTempPath
        End If

        mbBreakUpdateDBAll = False
        ' Отображаем ProgressBar
        CreateProgressNew
        cmdBreakUpdateDB.Visible = True
        ' Начальные пременные прогрессбара
        lngFindCheckCountTemp = FindCheckCount

        If lngFindCheckCountTemp > 0 Then
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
                FlatBorderButton .hWnd
                .Refresh

                ' Прерываем процесс распаковки
                If mbBreakUpdateDBAll Then
                    MsgBox strMessages(27) & vbNewLine & .Tag, vbInformation, strProductName
                    Exit For
                End If

                strPackFileName = .Tag
                strPackFileName_woExt = FileName_woExt(strPackFileName)

                If UnPackDPFile(strPathDRP, strPackFileName, arrCheckDP(1, i), ArchTempPath) = False Then
                    If Not mbSilentRun Then
                        MsgBox strMessages(13) & vbNewLine & strPackFileName, vbInformation, strProductName
                    End If
                End If

                If chkPackFiles(arrCheckDP(0, i)).Value Then
                    chkPackFiles(arrCheckDP(0, i)).Value = False
                End If

                FlatBorderButton .hWnd, False
                .Refresh
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
            If LenB(arrOSList(SSTab1.Tab).PathPhysX) > 0 Then
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
            If LenB(arrOSList(SSTab1.Tab).PathLanguages) > 0 Then
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
            If LenB(arrOSList(SSTab1.Tab).PathRuntimes) > 0 Then
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
                        For i = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)
                            strPackFileName = acmdPackFiles(arrCheckDP(0, i)).Tag
                            strPackFileName_woExt = FileName_woExt(strPackFileName)
                            ArchTempPath = strWorkTempBackSL & strPackFileName_woExt
                            WorkWithFinish strPathDRP, strPackFileName, ArchTempPath, arrCheckDP(1, i)
                        Next

                        ' Обновление подсказки
                        For i = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)
                            strPackFileName = acmdPackFiles(arrCheckDP(0, i)).Tag
                            ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, CInt(arrCheckDP(0, i))
                        Next
                    End If
                End If
            End If

            ChangeStatusTextAndDebug strMessages(85) & " " & ReadExitCodeString, "Install from : " & strPackFileName & " finish."

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
            ChangeStatusTextAndDebug strMessages(125) & " " & ArchTempPath

            If MsgBox(strMessages(125) & str2vbNewLine & strMessages(133), vbYesNo, strProductName) = vbYes Then
                RunUtilsShell ArchTempPath, False
            End If
        End If

        mbUnpackAdditionalFile = False
    Else

        If Not mbSilentRun Then
            MsgBox strMessages(14), vbInformation, strProductName
        End If

        DebugMode "GroupInstallDP-DpPack is not Checked"
        ChangeStatusTextAndDebug strMessages(14)
    End If

EndedSub:
    BlockControl True
    DebugMode "GroupInstallDP-End"
    FindCheckCount False
    mbBreakUpdateDBAll = False
    
    ChangeFrmMainCaption
    pbProgressBar.Visible = False
    ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone

End Sub

' Запуск процесс установки(или обновления БД) выделенных пакетов драйверов
Private Sub InsOrUpdSelectedDP(ByVal mbInstallMode As Boolean)

    If cmdRunTask.Enabled Then
        If mbInstallMode Then
            If optRezim_Upd.Value Then
                SelectStartMode 1, False
            End If

            mbGroupTask = True
            mbooSelectInstall = False
            Sleep 200
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

Private Sub lblOsInfoChange()

Dim str64bit                            As String
Dim lblOsInfoCaption                    As String

    If mbIsWin64 Then
        str64bit = " x64 Edition"
    Else
        str64bit = " x86 Edition"
    End If

    lblOsInfoCaption = LocaliseString(strPCLangCurrentPath, strFormName, "lblOsInfo", lblOsInfo.Caption)
    'lblOsInfo.Caption = lblOsInfoCaption & " " & OSInfoWMI(0) & " " & " (" & OSInfoWMI(4) & "." & OSInfoWMI(1) & " " & OSInfoWMI(2) & ")" & str64bit
    lblOsInfo.Caption = lblOsInfoCaption & " " & OSInfo.Name & " " & " (" & OSInfo.VerFullwBuild & " " & OSInfo.ServicePack & ")" & str64bit

End Sub

Private Sub lblOsInfo_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    If mnuUtils_SIV.Visible Then mnuUtils_SIV_Click

End Sub

Private Sub PutAllDrivers2Log()
Dim i                                   As Long
Dim strTTipTextHeaders                  As String
Dim strTemp                             As String
Const strTableHwidHeaderDP              As String = "Drivers in DriverPack"
Dim strLineAll                          As String

    DebugMode "===============================List of all found a matched driver==================================="
    'Формируем шапку для подсказки
    strTTipTextHeaders = strTTipTextDrv2Install & vbNewLine & String$(maxSizeRowAllLineMax, "-") & vbNewLine & _
                         UCase$( _
                         strTableHwidHeader1 & Space$(lngSizeRow1Max - lngTableHwidHeader1 + 1) & "| " & _
                         strTableHwidHeaderDP & Space$(lngSizeRowDPMax - Len(strTableHwidHeaderDP) + 1) & "| " & _
                         strTableHwidHeader2 & Space$(lngSizeRow2Max - lngTableHwidHeader2 + 1) & "| " & _
                         strTableHwidHeader3 & Space$(lngSizeRow3Max - lngTableHwidHeader3 + 1) & "| " & _
                         strTableHwidHeader4 & Space$(lngSizeRow4Max - lngTableHwidHeader4 + 1) & "| " & _
                         strTableHwidHeader9 & Space$(lngSizeRow9Max - lngTableHwidHeader9 + 1) & "| " & _
                         strTableHwidHeader5 & Space$(lngSizeRow5Max - lngTableHwidHeader5 + 1) & "| " & _
                         strTableHwidHeader6 & Space$(lngSizeRow6Max - lngTableHwidHeader6 + 1) & "| " & _
                         strTableHwidHeader7 _
                         ) & vbNewLine & String$(maxSizeRowAllLineMax, "-") & vbNewLine
    DebugMode strTTipTextHeaders

    ReDim Preserve arrDriversList(13, lngDriversArrCount - 1)
    QuickSortMDArray arrDriversList, 1, 0

    For i = 0 To UBound(arrDriversList, 2)
        'strDevID
        strTemp = arrDriversList(0, i)
        arrDriversList(0, i) = strTemp & Space$(lngSizeRow1Max - Len(strTemp) + 1) & "| "
        'strDevPath
        strTemp = arrDriversList(1, i)
        arrDriversList(1, i) = strTemp & Space$(lngSizeRow2Max - Len(strTemp) + 1) & "| "
        'strDevInf
        strTemp = arrDriversList(2, i)
        arrDriversList(2, i) = strTemp & Space$(lngSizeRow3Max - Len(strTemp) + 1) & "| "
        'strDevVer
        strTemp = arrDriversList(3, i)
        arrDriversList(3, i) = strTemp & Space$(lngSizeRow4Max - Len(strTemp) + 1) & "| "
        'strDevVerLocal
        strTemp = arrDriversList(4, i)
        arrDriversList(4, i) = strTemp & Space$(lngSizeRow5Max - Len(strTemp) + 1) & "| "
        ' strPriznakSravnenia
        strTemp = arrDriversList(7, i)
        arrDriversList(7, i) = strTemp & Space$(lngSizeRow9Max - Len(strTemp) + 1) & "| "
        'strDevStatus & strDevName
        strTemp = arrDriversList(5, i)
        arrDriversList(5, i) = strTemp & Space$(lngSizeRow6Max - Len(strTemp) + 1) & "| "
        ' Секция
        strTemp = arrDriversList(8, i)
        arrDriversList(8, i) = strTemp & Space$(lngSizeRow13Max - Len(strTemp) + 1) & "|"
        ' Имя DP
        strTemp = arrDriversList(13, i)
        arrDriversList(13, i) = strTemp & Space$(lngSizeRowDPMax - Len(strTemp) + 1) & "|"
        ' Итоговый
        strLineAll = arrDriversList(0, i) & arrDriversList(13, i) & arrDriversList(1, i) & (arrDriversList(2, i) & arrDriversList(3, i) & arrDriversList(7, i)) & (arrDriversList(4, i) & arrDriversList(5, i) & arrDriversList(6, i))
        DebugMode strLineAll
    Next
    DebugMode "===================================================================================================="
End Sub

'! -----------------------------------------------------------
'!  Функция     :  LoadButton
'!  Переменные  :
'!  Описание    :  Загрузка кнопок при старте программы
'! -----------------------------------------------------------
Private Sub LoadButton()

Dim i                                   As Long
Dim cnt                                 As Long
Dim pbStart                             As Long
Dim pbDelta                             As Long
Dim strPathFolderDRP                    As String
Dim strPathFolderDB                     As String

    On Error Resume Next

    DebugMode "LoadButton-Start"
    mbNextTab = False
    frTabPanel.Visible = False
    lngCntBtn = 0
    cnt = UBound(arrOSList)
    
    With ctlProgressBar1
        pbStart = .Value
        .SetTaskBarProgressState PrbTaskBarStateInProgress
        .SetTaskBarProgressValue pbStart, 1000
    End With

    If cnt > 0 Then
        pbDelta = (1000 - pbStart) / (cnt + 1)
    Else
        pbDelta = 1000 - pbStart
    End If

    i = 0

    ' массив со списокм драйверов для будущего использования для каждой кнопки
    lngDriversArrCount = 0
    lngMaxDriversArrCount = 100
    ReDim Preserve arrDriversList(13, lngMaxDriversArrCount) As String

    For i = 0 To cnt
        strPathFolderDRP = arrOSList(i).drpFolderFull
        strPathFolderDB = arrOSList(i).devIDFolderFull
        ChangeStatusTextAndDebug strMessages(69) & " " & strPathFolderDRP
        DebugMode vbTab & "Analize Folder DRP: " & strPathFolderDRP
        DebugMode vbTab & "Analize Folder DB: " & strPathFolderDB
        pbProgressBar.Refresh

        If Not arrOSList(i).DPFolderNotExist Then
            ' Запуск процедуры создания кнопок на вкладке
            CreateButtonsonSSTab strPathFolderDRP, strPathFolderDB, i, pbDelta
        Else
            
            SSTab1.TabEnabled(i) = False

            If mbTabHide Then
                SSTab1.TabVisible(i) = False
            End If

            ' грузим вкладки , но делаем скрытыми
            If i > 0 Then
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
        ChangeStatusTextAndDebug strMessages(86), "Create Buttons: True"
    Else

        If acmdPackFiles.Count <= 1 Then
            ChangeStatusTextAndDebug strMessages(87), "Create Buttons: False"
            mnuRezimBaseDrvUpdateALL.Enabled = False
        End If

        SSTab1.Enabled = False
    End If
    
    DebugMode "LoadButton-End"

End Sub

Private Sub LoadCmdRunTask()

    With cmdRunTask
        .SetPopupMenu mnuContextMenu3
        .DropDownSeparator = True
        .DropDownSymbol = 6
    End With

End Sub

' Изменение описания кнопки cmdViewAllDevice
Public Sub LoadCmdViewAllDeviceCaption()
    lngNotFinedDriversInDP = CalculateUnknownDrivers

    If lngNotFinedDriversInDP > 0 Then
        cmdViewAllDevice.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdViewAllDevice", cmdViewAllDevice.Caption) & vbNewLine & strMessages(122) & " " & lngNotFinedDriversInDP
        'cmdViewAllDevice.TextColor = vbRed
    Else
        cmdViewAllDevice.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdViewAllDevice", cmdViewAllDevice.Caption)

        'cmdViewAllDevice.TextColor = cmdRunTask.TextColor
    End If

End Sub

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
    'LoadIconImage2Btn cmdRunTask, "BTN_RUNTASK", strPathImageMainWork
    LoadIconImage2BtnJC cmdRunTask, "BTN_RUNTASK", strPathImageMainWork
    LoadIconImage2BtnJC cmdBreakUpdateDB, "BTN_BREAK_UPDATE", strPathImageMainWork
    LoadIconImage2BtnJC cmdViewAllDevice, "BTN_VIEW_SEARCH", strPathImageMainWork
    LoadIconImage2BtnJC cmdCheck, "BTN_CHECKMARK", strPathImageMainWork
    '--------------------- Группы
    LoadIconImage2FrameJC frRezim, "FRAME_GROUP", strPathImageMainWork
    DebugMode "LoadIconImage-End"

End Sub

'заполнение списка на выделение
Private Sub LoadListChecked()

    cmbCheckButton.Clear
    ' Режимы выделения
    strCmbChkBtnListElement1 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement1", "Все на текущей вкладке")
    strCmbChkBtnListElement2 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement2", "Сброс отметок")
    strCmbChkBtnListElement3 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement3", "Все")
    strCmbChkBtnListElement4 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement4", "Все новые")
    strCmbChkBtnListElement5 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement5", "Неустановленные")
    strCmbChkBtnListElement6 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbCheckButtonListElement6", "Рекомендуемые")

    If optRezim_Upd.Value Then

        With cmbCheckButton
            .AddItem strCmbChkBtnListElement1, 0
            .AddItem strCmbChkBtnListElement2, 1
            .AddItem strCmbChkBtnListElement3, 2
            .AddItem strCmbChkBtnListElement4, 3
            .ListIndex = 3
            ' Подсчитываем кол-во пакетов не имеющих БД, и если их нет то ставим "Все новые"
            If FindNoDBCount = 0 Then .ListIndex = 2
        End With

        'cmbCheckButton
    ElseIf optRezim_Ust.Value Then

        With cmbCheckButton
            .AddItem strCmbChkBtnListElement2, 0
            .AddItem strCmbChkBtnListElement5, 1
            .AddItem strCmbChkBtnListElement6, 2
            .AddItem strCmbChkBtnListElement1, 3
            .ListIndex = 1
        End With

        'CMBCHECKBUTTON
    Else

        With cmbCheckButton
            .AddItem strCmbChkBtnListElement2, 0
            .AddItem strCmbChkBtnListElement5, 1
            .AddItem strCmbChkBtnListElement6, 2
            .AddItem strCmbChkBtnListElement1, 3
            .ListIndex = 1
        End With

        'CMBCHECKBUTTON
    End If

End Sub

Private Sub LoadSSTab2Desc()

Dim i                                   As Long

    SetTabPropertiesTabDrivers

    With SSTab2

        For i = .LBound To .UBound
            .Item(i).TabCaption(0) = strSSTabTypeDPTab1
            .Item(i).TabCaption(1) = strSSTabTypeDPTab2
            .Item(i).TabCaption(2) = strSSTabTypeDPTab3
            .Item(i).TabCaption(3) = strSSTabTypeDPTab4
            .Item(i).TabCaption(4) = strSSTabTypeDPTab5
        Next
    End With

    'SSTab2
End Sub

Private Sub Localise(ByVal StrPathFile As String)

' изменяем шрифт
    FontCharsetChange
    'Frame
    frRezim.Caption = LocaliseString(StrPathFile, strFormName, "frRezim", frRezim.Caption)
    frDescriptionIco.Caption = LocaliseString(StrPathFile, strFormName, "frDescriptionIco", frDescriptionIco.Caption)
    frRunChecked.Caption = LocaliseString(StrPathFile, strFormName, "frRunChecked", frRunChecked.Caption)
    frCheck.Caption = LocaliseString(StrPathFile, strFormName, "frCheck", frCheck.Caption)
    frInfo.Caption = LocaliseString(StrPathFile, strFormName, "frInfo", frInfo.Caption)
    ' Описание режимов
    optRezim_Intellect.Caption = LocaliseString(StrPathFile, strFormName, "RezimIntellect", optRezim_Intellect.Caption)
    optRezim_Ust.Caption = LocaliseString(StrPathFile, strFormName, "RezimUst", optRezim_Ust.Caption)
    optRezim_Upd.Caption = LocaliseString(StrPathFile, strFormName, "RezimUpd", optRezim_Upd.Caption)
    optRezim_Intellect.ToolTipText = LocaliseString(StrPathFile, strFormName, "RezimIntellectTip", optRezim_Intellect.ToolTipText)
    optRezim_Ust.ToolTipText = LocaliseString(StrPathFile, strFormName, "RezimUstTip", optRezim_Ust.ToolTipText)
    optRezim_Upd.ToolTipText = LocaliseString(StrPathFile, strFormName, "RezimUpdTip", optRezim_Upd.ToolTipText)
    ' Меню
    mnuRezim.Caption = LocaliseString(StrPathFile, strFormName, "mnuRezim", mnuRezim.Caption)
    mnuRezimBaseDrvUpdateALL.Caption = LocaliseString(StrPathFile, strFormName, "mnuRezimBaseDrvUpdateALL", mnuRezimBaseDrvUpdateALL.Caption)
    mnuRezimBaseDrvUpdateNew.Caption = LocaliseString(StrPathFile, strFormName, "mnuRezimBaseDrvUpdateNew", mnuRezimBaseDrvUpdateNew.Caption)
    mnuRezimBaseDrvClean.Caption = LocaliseString(StrPathFile, strFormName, "mnuRezimBaseDrvClean", mnuRezimBaseDrvClean.Caption)
    mnuService.Caption = LocaliseString(StrPathFile, strFormName, "mnuService", mnuService.Caption)
    mnuShowHwidsTxt.Caption = LocaliseString(StrPathFile, strFormName, "mnuShowHwidsTxt", mnuShowHwidsTxt.Caption)
    mnuShowHwidsXLS.Caption = LocaliseString(StrPathFile, strFormName, "mnuShowHwidsXLS", mnuShowHwidsXLS.Caption)
    mnuShowHwidsAll.Caption = LocaliseString(StrPathFile, strFormName, "mnuShowHwidsAll", mnuShowHwidsAll.Caption)
    mnuUpdateStatusAll.Caption = LocaliseString(StrPathFile, strFormName, "mnuUpdateStatusAll", mnuUpdateStatusAll.Caption)
    mnuReCollectHWID.Caption = LocaliseString(StrPathFile, strFormName, "mnuReCollectHWID", mnuReCollectHWID.Caption)
    mnuAutoInfoAfterDelDRV.Caption = LocaliseString(StrPathFile, strFormName, "mnuAutoInfoAfterDelDRV", mnuAutoInfoAfterDelDRV.Caption)
    mnuRunSilentMode.Caption = LocaliseString(StrPathFile, strFormName, "mnuRunSilentMode", mnuRunSilentMode.Caption)
    mnuViewDPInstLog.Caption = LocaliseString(StrPathFile, strFormName, "mnuViewDPInstLog", mnuViewDPInstLog.Caption)
    mnuOptions.Caption = LocaliseString(StrPathFile, strFormName, "mnuOptions", mnuOptions.Caption)
    mnuMainUtils.Caption = LocaliseString(StrPathFile, strFormName, "mnuMainUtils", mnuMainUtils.Caption)
    mnuUtils_devmgmt.Caption = LocaliseString(StrPathFile, strFormName, "mnuUtils_devmgmt", mnuUtils_devmgmt.Caption)
    mnuMainAbout.Caption = LocaliseString(StrPathFile, strFormName, "mnuMainAbout", mnuMainAbout.Caption)
    mnuLinks.Caption = LocaliseString(StrPathFile, strFormName, "mnuLinks", mnuLinks.Caption)
    mnuHistory.Caption = LocaliseString(StrPathFile, strFormName, "mnuHistory", mnuHistory.Caption)
    mnuHelp.Caption = LocaliseString(StrPathFile, strFormName, "mnuHelp", mnuHelp.Caption)
    mnuHomePage.Caption = LocaliseString(StrPathFile, strFormName, "mnuHomePage", mnuHomePage.Caption)
    mnuHomePage1.Caption = LocaliseString(StrPathFile, strFormName, "mnuHomePage1", mnuHomePage1.Caption)
    mnuDriverPacks.Caption = LocaliseString(StrPathFile, strFormName, "mnuDriverPacks", mnuDriverPacks.Caption)
    mnuCreateRestorePoint.Caption = LocaliseString(StrPathFile, strFormName, "mnuCreateRestorePoint", mnuCreateRestorePoint.Caption)
    mnuUpdateStatusTab.Caption = LocaliseString(StrPathFile, strFormName, "mnuUpdateStatusTab", mnuUpdateStatusTab.Caption)
    mnuReCollectHWIDTab.Caption = LocaliseString(StrPathFile, strFormName, "mnuReCollectHWIDTab", mnuReCollectHWIDTab.Caption)
    mnuDriverPacksOnMySite.Caption = LocaliseString(StrPathFile, strFormName, "mnuDriverPacksOnMySite", mnuDriverPacksOnMySite.Caption)
    mnuCheckUpd.Caption = LocaliseString(StrPathFile, strFormName, "mnuCheckUpd", mnuCheckUpd.Caption)
    mnuDonate.Caption = LocaliseString(StrPathFile, strFormName, "mnuDonate", mnuDonate.Caption)
    mnuLicence.Caption = LocaliseString(StrPathFile, strFormName, "mnuLicence", mnuLicence.Caption)
    mnuAbout.Caption = LocaliseString(StrPathFile, strFormName, "mnuAbout", mnuAbout.Caption)
    mnuModulesVersion.Caption = LocaliseString(StrPathFile, strFormName, "mnuModulesVersion", mnuUtils_devmgmt.Caption)
    mnuContextXLS.Caption = LocaliseString(StrPathFile, strFormName, "mnuContextXLS", mnuContextXLS.Caption)
    mnuContextTxt.Caption = LocaliseString(StrPathFile, strFormName, "mnuContextTxt", mnuContextTxt.Caption)
    mnuContextToolTip.Caption = LocaliseString(StrPathFile, strFormName, "mnuContextToolTip", mnuContextToolTip.Caption)
    mnuContextUpdStatus.Caption = LocaliseString(StrPathFile, strFormName, "mnuContextUpdStatus", mnuContextUpdStatus.Caption)
    mnuContextEditDPName.Caption = LocaliseString(StrPathFile, strFormName, "mnuContextEditDPName", mnuContextEditDPName.Caption)
    mnuContextDeleteDRP.Caption = LocaliseString(StrPathFile, strFormName, "mnuContextDeleteDRP", mnuContextDeleteDRP.Caption)
    mnuContextTestDRP.Caption = LocaliseString(StrPathFile, strFormName, "mnuContextTestDRP", mnuContextTestDRP.Caption)
    mnuContextDeleteDevIDs.Caption = LocaliseString(StrPathFile, strFormName, "mnuContextDeleteDevIDs", mnuContextDeleteDevIDs.Caption)
    mnuContextLegendIco.Caption = LocaliseString(StrPathFile, strFormName, "mnuContextLegendIco", mnuContextLegendIco.Caption)
    mnuMainLang.Caption = LocaliseString(StrPathFile, strFormName, "mnuMainLang", mnuMainLang.Caption)
    mnuLangStart.Caption = LocaliseString(StrPathFile, strFormName, "mnuLangStart", mnuLangStart.Caption)
    strContextInstall1 = LocaliseString(StrPathFile, strFormName, "mnuContextInstall1", "Обычная установка")
    strContextInstall2 = LocaliseString(StrPathFile, strFormName, "mnuContextInstall2", "Выборочная установка")
    strContextInstall3 = LocaliseString(StrPathFile, strFormName, "mnuContextInstall3", "Распаковать драйвера в каталог - Все подобранные")
    strContextInstall4 = LocaliseString(StrPathFile, strFormName, "mnuContextInstall4", "Распаковать драйвера в каталог - Выбрать...")
    mnuContextInstall(0).Caption = LocaliseString(StrPathFile, strFormName, "mnuContextInstall1", mnuContextInstall(0).Caption)
    mnuContextInstall(2).Caption = LocaliseString(StrPathFile, strFormName, "mnuContextInstall2", mnuContextInstall(1).Caption)
    mnuContextInstall(4).Caption = LocaliseString(StrPathFile, strFormName, "mnuContextInstall3", mnuContextInstall(3).Caption)
    mnuContextInstall(5).Caption = LocaliseString(StrPathFile, strFormName, "mnuContextInstall4", mnuContextInstall(4).Caption)
    mnuCreateBackUp.Caption = LocaliseString(StrPathFile, strFormName, "mnuCreateBackUp", mnuCreateBackUp.Caption)
    mnuContextCopyHWIDs.Caption = LocaliseString(StrPathFile, strFormName, "mnuContextCopyHWIDs", mnuContextCopyHWIDs.Caption)
    mnuContextCopyHWIDDesc.Caption = "HWID" & vbTab & "Device Name"
    mnuDelDuplicateOldDP.Caption = LocaliseString(StrPathFile, strFormName, "mnuDelDuplicateOldDP", mnuDelDuplicateOldDP.Caption)

    mnuLoadOtherPC.Caption = LocaliseString(StrPathFile, strFormName, "mnuLoadOtherPC", mnuLoadOtherPC.Caption)
    mnuSaveInfoPC.Caption = LocaliseString(StrPathFile, strFormName, "mnuSaveInfoPC", mnuSaveInfoPC.Caption)
    'Кнопки
    cmdRunTask.Caption = LocaliseString(StrPathFile, strFormName, "cmdRunTask", cmdRunTask.Caption)
    cmdCheck.Caption = LocaliseString(StrPathFile, strFormName, "cmdCheck", cmdCheck.Caption)
    cmdBreakUpdateDB.Caption = LocaliseString(StrPathFile, strFormName, "cmdBreakUpdateDB", cmdBreakUpdateDB.Caption)
    cmdViewAllDevice.Caption = LocaliseString(StrPathFile, strFormName, "cmdViewAllDevice", cmdViewAllDevice.Caption)
    ' Лейблы
    lblPCInfo.Caption = LocaliseString(StrPathFile, strFormName, "lblPCInfo", lblPCInfo.Caption) & " " & strCompModel
    lblNoDP4Mode.Caption = LocaliseString(StrPathFile, strFormName, "lblNoDP4Mode", lblNoDP4Mode.Caption)
    lblNoDPInProgram.Caption = LocaliseString(StrPathFile, strFormName, "lblNoDPInProgram", lblNoDPInProgram.Caption)
    ' Другие параметры
    strTableHwidHeader1 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader1", "-HWID-")
    strTableHwidHeader2 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader2", "-Путь-")
    strTableHwidHeader3 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader3", "-Файл-")
    strTableHwidHeader4 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader4", "-Версия(БД)-")
    strTableHwidHeader5 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader5", "-Версия(PC)-")
    strTableHwidHeader6 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader6", "-Статус-")
    strTableHwidHeader7 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader7", "-Наименование устройства-")
    strTableHwidHeader8 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader8", "-Пакет драйверов-")
    strTableHwidHeader9 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader9", "-!-")
    strTableHwidHeader10 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader10", "-Производитель-")
    strTableHwidHeader11 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader11", "-Совместимый HWID-")
    strTableHwidHeader12 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader12", "-Код устройства-")
    strTableHwidHeader13 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader13", "-Секция-")
    strTableHwidHeader14 = LocaliseString(StrPathFile, strFormName, "TableHwidHeader14", "Найден в пакете")
    strTTipTextTitle = LocaliseString(StrPathFile, strFormName, "ToolTipTextTitle", "Файл пакета драйверов:")
    strTTipTextFileSize = LocaliseString(StrPathFile, strFormName, "ToolTipTextFileSize", "Размер файла:")
    strTTipTextClassDRV = LocaliseString(StrPathFile, strFormName, "ToolTipTextClassDRV", "Класс драйверов:")
    strTTipTextDrv2Install = LocaliseString(StrPathFile, strFormName, "ToolTipTextDrv2Install", "ДРАЙВЕРА ДОСТУПНЫЕ ДЛЯ УСТАНОВКИ:")
    strTTipTextDrv4UnsupOS = LocaliseString(StrPathFile, strFormName, "ToolTipTextDrv4UnsupportedOS", "ВНИМАНИЕ! ДРАЙВЕРА ДЛЯ ДРУГОЙ ОС." & vbNewLine & "ОБАБОТКА ВКЛАДКИ ВЫКЛЮЧЕНА В НАСТРОЙКАХ")
    strTTipTextTitleStatus = LocaliseString(StrPathFile, strFormName, "ToolTipTextTitleStatus", "Подробное описание:")
    strSSTabTypeDPTab1 = LocaliseString(StrPathFile, strFormName, "SSTabTypeDPTab1", "Все драйверпаки")
    strSSTabTypeDPTab2 = LocaliseString(StrPathFile, strFormName, "SSTabTypeDPTab2", "Доступно обновление")
    strSSTabTypeDPTab3 = LocaliseString(StrPathFile, strFormName, "SSTabTypeDPTab3", "Неустановленные")
    strSSTabTypeDPTab4 = LocaliseString(StrPathFile, strFormName, "SSTabTypeDPTab4", "Установленные")
    strSSTabTypeDPTab5 = LocaliseString(StrPathFile, strFormName, "SSTabTypeDPTab5", "БД не создана")

    ' Прописываем как константу длину названия колонок
    lngTableHwidHeader1 = Len(strTableHwidHeader1)
    lngTableHwidHeader2 = Len(strTableHwidHeader2)
    lngTableHwidHeader3 = Len(strTableHwidHeader3)
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
    ' Перегрузка ToolTipStatus
    ToolTipStatusLoad
    ' Изменение меню кнопок
    ChangeMenuCaption
    ' Изменение кнопки RunTask
    LoadCmdRunTask
    ' Изменение SSTab2
    LoadSSTab2Desc
    ' Перегрузка сообщений
    LocaliseMessage strPCLangCurrentPath

    If mbDP_Is_aFolder Then
        frRezim.Caption = frRezim.Caption & " " & strMessages(150)
    End If

    ' Установка текста панели
    ctlUcStatusBar1.PanelText(1) = strMessages(127)

    ' Если это не началный запуск программы, то изменяем еще и эти параметры
    If Not mbFirstStart Then
        ' изменение caption кнопки CmdViewAll
        LoadCmdViewAllDeviceCaption
        ' Перезагрузка всплывающих подсказок для кнопок с драйверами
        Me.Font.Name = strMainForm_FontName
        Me.Font.Size = lngMainForm_FontSize
        ToolTipBtnReLoad
    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuAbout_Click
'!  Переменные  :
'!  Описание    :  Меню - О программе
'! -----------------------------------------------------------
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuAutoInfoAfterDelDRV_Click()
    mnuAutoInfoAfterDelDRV.Checked = Not mnuAutoInfoAfterDelDRV.Checked
    mbAutoInfoAfterDelDRV = Not mbAutoInfoAfterDelDRV

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuCheckUpd_Click
'!  Переменные  :
'!  Описание    :  Меню - Проверить обновление
'! -----------------------------------------------------------
Private Sub mnuCheckUpd_Click()
    ctlUcStatusBar1.PanelText(1) = strMessages(145)
    ChangeStatusTextAndDebug strMessages(58)
    CheckUpd False
End Sub

Private Sub mnuContextDeleteDevID_Click(Index As Integer)

Dim strValue                            As String
Dim strValueDevID                       As String
Dim strValueDevID_x()                   As String
Dim mbDeleteDriverByHwidTemp            As Boolean

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

Private Sub mnuContextDeleteDRP_Click()

Dim i                                   As Long
Dim strPathDRP                          As String
Dim strPathDB                           As String
Dim strFullPathDRP                      As String
Dim strFullPathDB                       As String
Dim strFullPathDBIni                    As String

    If mbIsDriveCDRoom Then
        MsgBox strMessages(16), vbInformation, strProductName
    Else
        i = SSTab1.Tab
        strPathDRP = arrOSList(i).drpFolderFull
        strPathDB = arrOSList(i).devIDFolderFull
        strFullPathDRP = PathCombine(strPathDRP, acmdPackFiles(CurrentSelButtonIndex).Tag)
        strFullPathDB = PathCombine(strPathDB, FileNameFromPath(strCurSelButtonPath))
        strFullPathDBIni = Replace$(strFullPathDB, ".txt", "*.ini", , , vbTextCompare)

        If MsgBox(strMessages(17) & " '" & acmdPackFiles(CurrentSelButtonIndex).Tag & "' ?", vbQuestion + vbYesNo, strProductName) = vbYes Then
            If PathExists(strFullPathDRP) Then
                If Not PathIsAFolder(strFullPathDRP) Then
                    DebugMode "Delete file: " & strFullPathDRP
                    DeleteFiles strFullPathDRP
                End If
            End If

            If PathExists(strFullPathDB) Then
                If Not PathIsAFolder(strFullPathDB) Then
                    DebugMode "Delete file: " & strFullPathDB
                    DeleteFiles strFullPathDB
                    'Удаление секции о данном пакете из ini-файла
                    IniDelAllKeyPrivate FileName_woExt(FileNameFromPath(strCurSelButtonPath)), PathCombine(strPathDB, "DevDBVersions.ini")
                End If
            End If

            If PathExists(strFullPathDBIni) Then
                If Not PathIsAFolder(strFullPathDBIni) Then
                    DebugMode "Delete file: " & strFullPathDBIni
                    DeleteFiles strFullPathDBIni
                End If
            End If

            acmdPackFiles(CurrentSelButtonIndex).Visible = False
            chkPackFiles(CurrentSelButtonIndex).Visible = False
            chkPackFiles(CurrentSelButtonIndex).Value = False
            ChangeStatusTextAndDebug strMessages(88) & " " & strFullPathDRP
        End If
    End If

End Sub

Private Sub mnuContextEditDPName_Click()

    If Not FileisReadOnly(strSysIni) Then
        EditOrReadDPName CurrentSelButtonIndex
    End If

End Sub

Private Sub mnuContextInstall_Click(Index As Integer)
    mbGroupTask = True
    mbOnlyUnpackDP = False

    Select Case Index

        Case 0
            mbooSelectInstall = False
            mbOnlyUnpackDP = False

        Case 2
            mbooSelectInstall = True
            mbOnlyUnpackDP = False

        Case 4
            mbooSelectInstall = False
            mbOnlyUnpackDP = True

        Case 5
            mbooSelectInstall = True
            mbOnlyUnpackDP = True

    End Select

    GroupInstallDP
    mbGroupTask = False
    BlockControl True

End Sub

Private Sub mnuContextLegendIco_Click()
    frmLegendIco.Show vbModal, Me

End Sub

Private Sub mnuContextTestDRP_Click()

Dim cmdString                           As String
Dim strPackFileName                     As String
Dim strPathDRP                          As String

    strPackFileName = acmdPackFiles(CurrentSelButtonIndex).Tag
    strPathDRP = arrOSList(SSTab1.Tab).drpFolderFull
    cmdString = Kavichki & strArh7zExePATH & Kavichki & " t " & Kavichki & strPathDRP & strPackFileName & Kavichki & " -r"
    ChangeStatusTextAndDebug strMessages(109) & " " & strPackFileName

    BlockControl False
    If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
        MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
    Else

        ' Архиватор отработал на все 100%? Если нет то сообщаем
        If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
            ChangeStatusTextAndDebug strMessages(13) & " " & strPackFileName
            MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
        Else
            ChangeStatusTextAndDebug strMessages(110) & " " & strPackFileName
            MsgBox strMessages(110) & " " & strPackFileName, vbInformation, strProductName
        End If
    End If

    BlockControl True

End Sub

Private Sub mnuContextToolTip_Click()
    mbooSelectInstall = False

    If IsFormLoaded("frmListHwid") = False Then
        frmListHwid.Show vbModal, Me
    Else
        frmListHwid.FormLoadDefaultParam
        frmListHwid.FormLoadAction
        frmListHwid.Show vbModal, Me
    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuContextTxt_Click
'!  Переменные  :
'!  Описание    :  Меню - Файл БД в текстовом виде
'! -----------------------------------------------------------
Private Sub mnuContextTxt_Click()
    RunUtilsShell Kavichki & strCurSelButtonPath & Kavichki, False

End Sub

Private Sub mnuContextUpdStatus_Click()

Dim strPackFileName                     As String
Dim strPathDRP                          As String
Dim strPathDevDB                        As String

    strPathDRP = arrOSList(SSTab1.Tab).drpFolderFull
    strPathDevDB = arrOSList(SSTab1.Tab).devIDFolderFull
    strPackFileName = acmdPackFiles(CurrentSelButtonIndex).Tag
    ' Обновление подсказки
    ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, CInt(CurrentSelButtonIndex)

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuContextXLS_Click
'!  Переменные  :
'!  Описание    :  Меню - Файл БД в Excel
'! -----------------------------------------------------------
Private Sub mnuContextXLS_Click()

Dim strCurSelButtonPathTemp             As String

    strCurSelButtonPathTemp = strWorkTempBackSL & FileNameFromPath(strCurSelButtonPath)
    ' Копируем файл БД во временный каталог
    CopyFileTo strCurSelButtonPath, strCurSelButtonPathTemp
    ' Открываем в Excel
    OpenTxtInExcel strCurSelButtonPathTemp

End Sub

Private Sub mnuCreateBackUp_Click()

Dim lngMsgRet                           As Long

    lngMsgRet = MsgBox(strMessages(123), vbYesNo + vbQuestion, strProductName)

    Select Case lngMsgRet

        Case vbYes
            mnuHomePage1_Click

    End Select

End Sub

Private Sub mnuCreateRestorePoint_Click()

    If MsgBox(strMessages(115), vbQuestion + vbYesNo, strProductName) = vbYes Then
        CreateRestorePoint
    End If

End Sub

Private Sub mnuDonate_Click()
    frmDonate.Show vbModal, Me

End Sub

Private Sub mnuDriverPacks_Click()
    RunUtilsShell Kavichki & "http://driverpacks.net/driverpacks" & Kavichki, False

End Sub

Private Sub mnuDriverPacksOnMySite_Click()
    RunUtilsShell Kavichki & "http://adia-project.net/forum/index.php?topic=789.0" & Kavichki, False

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuHelp_Click
'!  Переменные  :
'!  Описание    :  Меню - Помощь
'! -----------------------------------------------------------
Private Sub mnuHelp_Click()

Dim cmdString                           As String
Dim strFilePathTemp                     As String

    strFilePathTemp = strAppPathBackSL & "Tools\Docs\" & strPCLangCurrentID & "\Help.html"

    If PathExists(strFilePathTemp) = False Then
        strFilePathTemp = strAppPathBackSL & "Tools\Docs\0409\Help.html"
    End If

    cmdString = Kavichki & strFilePathTemp & Kavichki
    RunUtilsShell cmdString, False

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuHistory_Click
'!  Переменные  :
'!  Описание    :  Меню - История изменений
'! -----------------------------------------------------------
Private Sub mnuHistory_Click()

Dim cmdString                           As String
Dim strFilePathTemp                     As String

    strFilePathTemp = strAppPathBackSL & "Tools\Docs\" & strPCLangCurrentID & "\history.txt"

    If PathExists(strFilePathTemp) = False Then
        strFilePathTemp = strAppPathBackSL & "Tools\Docs\0409\history.txt"
    End If

    cmdString = Kavichki & strFilePathTemp & Kavichki
    RunUtilsShell cmdString, False

End Sub

Private Sub mnuHomePage1_Click()
    RunUtilsShell Kavichki & "http://www.adia-project.net" & Kavichki, False

End Sub

Private Sub mnuHomePage_Click()
    RunUtilsShell Kavichki & "http://forum.oszone.net/thread-139908.html" & Kavichki, False

End Sub

Private Sub mnuLang_Click(Index As Integer)

Dim i                                   As Long
Dim ii                                  As Long
Dim strPathLng                          As String
Dim strPCLangCurrentIDTemp              As String
Dim strPCLangCurrentID_x()              As String

    i = Index + 1

    For ii = mnuLang.LBound To mnuLang.UBound
        mnuLang(ii).Checked = ii = Index
    Next
    strPathLng = arrLanguage(1, i)
    strPCLangCurrentPath = strPathLng
    strPCLangCurrentIDTemp = arrLanguage(3, i)
    lngDialog_Charset = GetCharsetFromLng(CLng(arrLanguage(6, i)))

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
        .Name = strMainForm_FontName
        .Size = lngMainForm_FontSize
        .Charset = lngDialog_Charset
    End With

    ChangeStatusTextAndDebug strMessages(142) & " " & arrLanguage(2, i), , , False
    FindCheckCount False
    If mbNoSupportedOS Then
        SelectStartMode 3, False
        BlockControl True
        BlockControlEx False
    End If

End Sub

Private Sub mnuLangStart_Click()
    mnuLangStart.Checked = Not mnuLangStart.Checked

End Sub

Private Sub mnuLicence_Click()
    frmLicence.Show vbModal, Me

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuLinks_Click
'!  Переменные  :
'!  Описание    :  Меню - Ссылки
'! -----------------------------------------------------------
Private Sub mnuLinks_Click()

Dim cmdString                           As String
Dim strFilePathTemp                     As String

    strFilePathTemp = strAppPathBackSL & "Tools\Docs\" & strPCLangCurrentID & "\Links.html"

    If PathExists(strFilePathTemp) = False Then
        strFilePathTemp = strAppPathBackSL & "Tools\Docs\0409\Links.html"
    End If

    cmdString = Kavichki & strFilePathTemp & Kavichki
    RunUtilsShell cmdString, False

End Sub

Private Sub mnuLoadOtherPC_Click()
    frmEmulate.Show vbModal, Me
End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuModulesVersion_Click
'!  Переменные  :
'!  Описание    :  Меню - Версии модулей
'! -----------------------------------------------------------
Private Sub mnuModulesVersion_Click()
    VerModules

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuOptions_Click
'!  Переменные  :
'!  Описание    :  Меню - Настройки
'! -----------------------------------------------------------
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
        ShellExecute Me.hWnd, "open", App.EXEName, vbNullString, strAppPath, SW_SHOWNORMAL
        Unload Me
    End If

End Sub

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

'! -----------------------------------------------------------
'!  Функция     :  mnuRezimBaseDrvClean_Click
'!  Переменные  :
'!  Описание    :  Меню - Очистка лишних файлов БД
'! -----------------------------------------------------------
Private Sub mnuRezimBaseDrvClean_Click()
    DeleteUnUsedBase

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuRezimBaseDrvUpdateALL_Click
'!  Переменные  :
'!  Описание    :  Меню - Обновление всех баз поочередно
'! -----------------------------------------------------------
Private Sub mnuRezimBaseDrvUpdateALL_Click()
    BaseUpdateOrRunTask False

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuRezimBaseDrvUpdateNew_Click
'!  Переменные  :
'!  Описание    :  Меню - Обновление новых баз поочередно
'! -----------------------------------------------------------
Private Sub mnuRezimBaseDrvUpdateNew_Click()

    If FindNoDBCount > 0 Then
        SilentCheckNoDB
        ' возвращаяем обратно стартовый режим
        SelectStartMode
    Else
        ChangeStatusTextAndDebug strMessages(68)
    End If

End Sub

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

Private Sub mnuSaveInfoPC_Click()
Dim strFilePathTo As String
Dim mbErrCopy As Boolean
    
    With New CommonDialog
        .Filter = "Text Files (*.TXT)|*.TXT"
        .DefaultExt = ".txt"
        .InitDir = GetSpecialFolderPath(CSIDL_DESKTOPDIRECTORY)
        .FileName = ExpandFileNamebyEnvironment("hwids_%PCMODEL%_" & strOsCurrentVersion & "_%OSBIT%")
        '.DialogTitle = "Select File"
        If .ShowSave = True Then
            strFilePathTo = .FileName
        End If
    End With

    If LenB(strFilePathTo) > 0 Then
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

Private Sub mnuService_Click()
    mnuViewDPInstLog.Enabled = PathExists(strWinDir & "DPINST.LOG")
End Sub

Private Sub mnuShowHwidsAll_Click()
    If IsFormLoaded("frmListHwidAll") = False Then
        frmListHwidAll.Show vbModal, Me
    Else
        frmListHwidAll.FormLoadDefaultParam
        frmListHwidAll.FormLoadAction
        frmListHwidAll.Show vbModal, Me
    End If

End Sub

Private Sub mnuShowHwidsAllBase_Click()

'frmListHwidAllBase.Show vbModal, Me
End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuShowHwidsTxt_Click
'!  Переменные  :
'!  Описание    :
'! -----------------------------------------------------------
Private Sub mnuShowHwidsTxt_Click()

    If PathExists(strHwidsTxtPathView) = False Then
        RunDevconView
    End If

    RunUtilsShell strHwidsTxtPathView, False

End Sub

Private Sub mnuShowHwidsXLS_Click()
    OpenTxtInExcel strResultHwidsTxtPath

End Sub

Private Sub mnuUpdateStatusAll_Click()
    UpdateStatusButtonAll
    ' Обновить список неизвестных дров и описание для кнопки
    LoadCmdViewAllDeviceCaption

End Sub

Private Sub mnuUpdateStatusTab_Click()
    UpdateStatusButtonTAB
    ' Обновить список неизвестных дров и описание для кнопки
    LoadCmdViewAllDeviceCaption

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuUtils_Click
'!  Переменные  :  Index As Integer
'!  Описание    :  Запуск дополнительной утилиты
'! -----------------------------------------------------------
Private Sub mnuUtils_Click(Index As Integer)

Dim i                                   As Long
Dim PathExe                             As String
Dim PathExe64                           As String
Dim Params                              As String
Dim cmdString                           As String

    i = Index
    PathExe = PathCollect(arrUtilsList(i, 1))
    PathExe64 = PathCollect(arrUtilsList(i, 2))

    If mbIsWin64 Then
        If LenB(PathExe64) > 0 Then
            PathExe = PathExe64
        End If
    End If

    Params = arrUtilsList(i, 3)

    If LenB(Params) = 0 Then
        cmdString = Kavichki & PathExe & Kavichki
    Else
        cmdString = Kavichki & PathExe & Kavichki & " " & Params
    End If

    RunUtilsShell cmdString, False

End Sub

Private Sub mnuUtils_DevManView_Click()

    If mbIsWin64 Then
        RunUtilsShell strDevManView_Path64
    Else
        RunUtilsShell strDevManView_Path
    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  mnuUtils_devmgmt_Click
'!  Переменные  :
'!  Описание    :  Запуск диспетчера устройств
'! -----------------------------------------------------------
Private Sub mnuUtils_devmgmt_Click()
    RunUtilsShell "devmgmt.msc", False

End Sub

Private Sub mnuUtils_DoubleDriver_Click()
    RunUtilsShell strDoubleDriver_Path

End Sub

Private Sub mnuUtils_SIV_Click()

    If mbIsWin64 Then
        RunUtilsShell strSIV_Path64
    Else
        RunUtilsShell strSIV_Path
    End If

End Sub

Private Sub mnuUtils_UDI_Click()
    RunUtilsShell strUDI_Path

End Sub

Private Sub mnuUtils_UnknownDevices_Click()
    RunUtilsShell strUnknownDevices_Path, True, True

End Sub

Private Sub mnuViewDPInstLog_Click()

Dim cmdString                           As String
Dim strLogPath                          As String
Dim strLogPathNew                       As String

    strLogPath = strWinDir & "DPINST.LOG"
    strLogPathNew = strWorkTempBackSL & "DPINST.LOG.TXT"

    If PathExists(strLogPath) Then
        CopyFileTo strLogPath, strLogPathNew
        cmdString = Kavichki & strLogPathNew & Kavichki
        RunUtilsShell cmdString, False
    Else
        DebugMode "cmdString - File not exist: " & strLogPath
    End If

End Sub

' Разные сообщения если нет поддерживаемых вкладок, или что-то нет так с пакетами
Private Sub NoSupportOSorNoDevBD()

Dim lngCnt                              As Long

    'Если нет поддерживаемых вкладок для вашей ОС, то
    If mbNoSupportedOS Then
        SelectStartMode 3, False
        BlockControl True
        BlockControlEx False
        ChangeStatusTextAndDebug strMessages(111)
        MsgBox strMessages(111) & vbNewLine & Replace$(optRezim_Upd.Caption, vbNewLine, " "), vbInformation, strProductName
        mbSilentRun = False
        mbRunWithParam = False
    End If

    ' Если есть несовместимый(ые) пакеты драйверов, то выводим сообщение
    If mbNotSupportedDevDB Then
        MsgBox strMessages(112), vbInformation & vbApplicationModal, strProductName
    End If

    ' Подсчитываем кол-во пакетов не имеющих БД, и выводим сообщение
    lngCnt = FindNoDBCount

    If lngCnt > 0 Then
        If MsgBox(lngCnt & " " & strMessages(103), vbYesNo + vbQuestion, strProductName) = vbYes Then
            ' Выставляем вкладку для которых нет БД
            SSTab2(SSTab1.Tab).Tab = 4
            DoEvents
            SelectStartMode 3, False
            ' собственно запуск создания БД
            SilentCheckNoDB
            ' возвращаяем обратно стартовый режим
            SelectStartMode
        End If
    End If

End Sub

Private Sub OpenTxtInExcel(ByVal strPathTxt As String)

Dim ExcelApp                            As Object
Dim ExcelDoc                            As Object

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

        'EXCELAPP
    End If

End Sub

Private Sub optRezim_Intellect_Click()

Dim ButtIndex                           As Long
Dim strSSTabCurrentOSListTemp           As String
Dim i                                   As Integer
Dim i_i                                 As Integer
Dim cntFindUnHideTab                    As Integer

    If Not mbFirstStart Then
        ButtIndex = acmdPackFiles.UBound

        For i = 0 To ButtIndex

            If ButtIndex > 0 Then

                With acmdPackFiles(i)
                    If Not (.Picture Is Nothing) Then
                        If .Picture = imgNo.Picture Or .Picture = imgNoDB.Picture Then
                            If .EnabledCtrl Then
                                .EnabledCtrl = False
                                chkPackFiles(i).Enabled = False
                            End If
    
                            .MenuExist = False
                        Else
                            .MenuExist = True
                        End If
                    End If
                End With
            Else

                With acmdPackFiles(0)

                    If .Visible Then
                        If Not (.Picture Is Nothing) Then
                            If .Picture = imgNo.Picture Or .Picture = imgNoDB.Picture Then
                                If .EnabledCtrl Then
                                    .EnabledCtrl = False
                                    chkPackFiles(0).Enabled = False
                                End If
    
                                .MenuExist = False
                            Else
                                .MenuExist = True
                            End If
                        End If
                    End If
                End With
            End If

        Next
    End If

    If mbTabBlock Then
        strSSTabCurrentOSListTemp = strSSTabCurrentOSList & " "

        For i = 0 To SSTab1.Tabs - 1

            If InStr(strSSTabCurrentOSListTemp, i & " ") = 0 Then
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

    'SSTab1
    cmdRunTask.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdRunTask1", cmdRunTask.Caption)
    cmdRunTask.SetPopupMenu mnuContextMenu3
    cmdRunTask.DropDownSeparator = True
    cmdRunTask.DropDownSymbol = 6
    'заполнение списка на выделение
    FindCheckCount
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
            If lngStartModeTab2 > 0 Then
                ' Если вкладка активна, то выставляем начальную
                If SSTab2(i_i).TabEnabled(lngStartModeTab2) = True Then
                    SSTab2(i_i).Tab = lngStartModeTab2
                Else
                    SSTab2(i_i).Tab = 0
                End If
            End If
        Next
    End If

End Sub

Private Sub optRezim_Upd_Click()

Dim i                                   As Integer
Dim i_i                                 As Integer
Dim cntFindUnHideTab                    As Integer

    If Not mbFirstStart Then

        For i = 0 To acmdPackFiles.UBound

            If Not acmdPackFiles(i).EnabledCtrl Then
                acmdPackFiles(i).EnabledCtrl = True
                chkPackFiles(i).Enabled = True
            End If

            acmdPackFiles(i).MenuExist = False
        Next
    End If

    If mbTabBlock Then

        For i = 0 To SSTab1.Tabs - 1

            If Not arrOSList(i).DPFolderNotExist Then
                If arrOSList(i).CntBtn = 0 Then
                    SSTab1.TabEnabled(i) = False
                Else
                    If Not SSTab1.TabVisible(i) Then SSTab1.TabVisible(i) = True
                    If Not SSTab1.TabEnabled(i) Then SSTab1.TabEnabled(i) = True
                End If

            Else

                If mbTabHide Then
                    SSTab1.TabVisible(i) = False
                End If
            End If

        Next
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

    cmdRunTask.UnsetPopupMenu
    cmdRunTask.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdRunTask", cmdRunTask.Caption)
    cmdRunTask.DropDownSeparator = False
    cmdRunTask.DropDownSymbol = 0
    'заполнение списка на выделение
    FindCheckCount
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
    If SSTab2(SSTab1.Tab).Tab > 0 Then
        If SSTab2(SSTab1.Tab).Tab < 4 Then
            For i_i = SSTab2.LBound To SSTab2.UBound
                SSTab2(i_i).Tab = 0
            Next
        End If
    End If

    mbSet2UpdateFromTab4 = False

End Sub

Private Sub optRezim_Ust_Click()

Dim ButtIndex                           As Integer
Dim i                                   As Integer
Dim i_i                                 As Integer
Dim strSSTabCurrentOSListTemp           As String
Dim cntFindUnHideTab                    As Integer

    If Not mbFirstStart Then
        ButtIndex = acmdPackFiles.UBound

        For i = 0 To ButtIndex

            If ButtIndex > 0 Then

                With acmdPackFiles(i)

                    If .Picture = imgNoDB.Picture Then
                        If .EnabledCtrl Then
                            .EnabledCtrl = False
                            chkPackFiles(i).Enabled = False
                        End If

                    Else

                        If Not .EnabledCtrl Then
                            .EnabledCtrl = True
                            chkPackFiles(i).Enabled = True
                        End If
                    End If

                    .MenuExist = False
                End With
            End If

        Next
    End If

    If mbTabBlock Then
        strSSTabCurrentOSListTemp = strSSTabCurrentOSList & " "

        For i = 0 To SSTab1.Tabs - 1

            If InStr(strSSTabCurrentOSListTemp, i & " ") = 0 Then
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

    cmdRunTask.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "cmdRunTask1", cmdRunTask.Caption)
    cmdRunTask.UnsetPopupMenu
    cmdRunTask.DropDownSeparator = False
    cmdRunTask.DropDownSymbol = 0
    'заполнение списка на выделение
    FindCheckCount
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
            If lngStartModeTab2 > 0 Then
                ' Если вкладка активна, то выставляем начальную
                If SSTab2(i_i).TabEnabled(lngStartModeTab2) = True Then
                    SSTab2(i_i).Tab = lngStartModeTab2
                Else
                    SSTab2(i_i).Tab = 0
                End If
            End If
        Next
    End If

End Sub

Private Sub pbProgressBar_Resize()
    cmdBreakUpdateDB.Left = (pbProgressBar.Width - cmdBreakUpdateDB.Width) / 2
    cmdBreakUpdateDB.Top = (pbProgressBar.Height - cmdBreakUpdateDB.Height - 25) / 2
End Sub

Private Function ReadExitCode(ByVal lngCode As Long) As String

Dim strResultText                       As String
Dim strCode                             As String
Dim strCodeWW                           As String
Dim strCodeXX                           As String
Dim strCodeYY                           As String
Dim strCodeZZ                           As String
Dim mbCodeNotInstall                    As Boolean
Dim strCodeNotInstallCount              As Long
Dim mbCodeInstall                       As Boolean
Dim strCodeInstallCount                 As Long
Dim strCodeReadyToInstallCount          As Long
Dim mbReadyToInstall                    As Boolean
Dim mbCodeReboot                        As Boolean

    DebugMode str2VbTab & "ReadExitCode-Start"
    DebugMode str2VbTab & "ReadExitCode: lngCode=" & CStr(lngCode)
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
                If strCode = "0" Then
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
    If LenB(strCodeXX) > 0 Then
        strCodeNotInstallCount = CLng("&H" & Trim$(strCodeXX))
    End If

    If LenB(strCodeYY) > 0 Then
        strCodeReadyToInstallCount = CLng("&H" & Trim$(strCodeYY))
    End If

    If LenB(strCodeZZ) > 0 Then
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
        strResultText = IIf(LenB(strResultText) > 0, strResultText & " ", vbNullString) & "Install: " & strCodeInstallCount
    End If

    If mbCodeNotInstall Then
        strResultText = IIf(LenB(strResultText) > 0, strResultText & " ", vbNullString) & "NotInstall: " & strCodeNotInstallCount
    End If

    If mbReadyToInstall Then
        strResultText = IIf(LenB(strResultText) > 0, strResultText & " ", vbNullString) & "ReadyToInstall: " & strCodeReadyToInstallCount
    End If

    If mbCodeReboot Then
        strResultText = IIf(LenB(strResultText) > 0, strResultText & " ", vbNullString) & "Need Reboot"
    End If

    If LenB(strResultText) > 0 Then
        ReadExitCode = "(" & strResultText & ")"
    Else
        ReadExitCode = vbNullString
    End If

    DebugMode str2VbTab & "ReadExitCode: strResultText=" & strResultText
    DebugMode str2VbTab & "ReadExitCode-End"

End Function

Private Sub ReadOrSaveToolTip(ByVal strPathDevDB As String, _
                              ByVal strPathDRP As String, _
                              ByVal strPackFileName As String, _
                              ByVal Index As Long, _
                              Optional ByVal mbSaveToolTip As Boolean = False, _
                              Optional ByVal mbReloadToolTip As Boolean = False)

Dim strTTipText                         As String
Dim strTTipTextTemp                     As String
Dim strClassesName                      As String
Dim strTTipTextHeadersTemp              As String
Dim strPackFileNameFull                 As String
Dim strFinishIniPath                    As String
Dim strTTipTextOnlyDrivers              As String
Dim strTTipSizeHeader_x()               As String
Dim maxLengthRow1                       As String
Dim maxLengthRow2                       As String
Dim maxLengthRow3                       As String
Dim maxLengthRow4                       As String
Dim maxLengthRow5                       As String
Dim maxLengthRow6                       As String
Dim maxLengthRow9                       As String
Dim TimeScriptRun                       As Long
Dim TimeScriptFinish                    As Long

    DebugMode str2VbTab & "ReadOrSaveToolTip-Start"
    TimeScriptRun = GetTickCount
    ' Небольшое прерывание для большего отклика от приложения
    DoEvents
    'Считываем класс пакета из файла
    strClassesName = vbNullString

    If LenB(strPackFileName) > 0 Then
        If mbReadClasses Then
            strFinishIniPath = BackslashAdd2Path(strPathDevDB) & FileName_woExt(strPackFileName) & ".ini"
            strClassesName = IniStringPrivate("DriverPack", "classes", strFinishIniPath)

            ' Если такого значения в файле нет, то ничего не добавляем
            If strClassesName = "no_key" Then
                strClassesName = vbNullString
            End If
        End If

        ' Всплывающая подсказка
        strPackFileNameFull = PathCombine(strPathDRP, strPackFileName)

        If LenB(strClassesName) > 0 Then
            If Not mbDP_Is_aFolder Then
                strTTipTextHeadersTemp = strPathDRP & str2vbNewLine & strPackFileName & vbNewLine & strTTipTextFileSize & " " & FileSizeApi(strPackFileNameFull) & vbNewLine & strTTipTextClassDRV & " " & strClassesName
            Else
                strTTipTextHeadersTemp = strPathDRP & str2vbNewLine & strPackFileName & vbNewLine & strTTipTextFileSize & " " & FolderSizeApi(strPackFileNameFull, True) & vbNewLine & strTTipTextClassDRV & " " & strClassesName
            End If
        Else
            If Not mbDP_Is_aFolder Then
                strTTipTextHeadersTemp = strPathDRP & str2vbNewLine & strPackFileName & vbNewLine & strTTipTextFileSize & " " & FileSizeApi(strPackFileNameFull)
            Else
                strTTipTextHeadersTemp = strPathDRP & str2vbNewLine & strPackFileName & vbNewLine & strTTipTextFileSize & " " & FolderSizeApi(strPackFileNameFull, True)
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

        If LenB(strTTipTextTemp) > 0 Then
            If strTTipTextTemp <> "Unsupported" And InStr(strTTipTextTemp, "|") Then
                'Формируем шапку для подсказки
                If mbReloadToolTip Then
                    strTTipSizeHeader_x = Split(arrTTipSize(Index), ";")
                    maxLengthRow1 = lngTableHwidHeader1
                    maxLengthRow2 = lngTableHwidHeader2
                    maxLengthRow3 = lngTableHwidHeader3
                    maxLengthRow4 = lngTableHwidHeader4
                    maxLengthRow9 = lngTableHwidHeader9
                    maxLengthRow5 = lngTableHwidHeader5
                    maxLengthRow6 = lngTableHwidHeader6

                    maxSizeRowAllLine = strTTipSizeHeader_x(0)
                    lngSizeRow1 = strTTipSizeHeader_x(1)
                    lngSizeRow2 = strTTipSizeHeader_x(2)
                    lngSizeRow3 = strTTipSizeHeader_x(3)
                    lngSizeRow4 = strTTipSizeHeader_x(4)
                    lngSizeRow9 = strTTipSizeHeader_x(5)
                    lngSizeRow5 = strTTipSizeHeader_x(6)
                    lngSizeRow6 = strTTipSizeHeader_x(7)

                    If lngSizeRow1 < maxLengthRow1 Then
                        lngSizeRow1 = maxLengthRow1
                    End If
                    If lngSizeRow2 < maxLengthRow2 Then
                        lngSizeRow2 = maxLengthRow2
                    End If
                    If lngSizeRow3 < maxLengthRow3 Then
                        lngSizeRow3 = maxLengthRow3
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

                strTTipTextHeaders = strTTipTextHeadersTemp & str2vbNewLine & strTTipTextDrv2Install & vbNewLine & _
                                     String$(maxSizeRowAllLine, "-") & vbNewLine & _
                                     UCase$(strTableHwidHeader1 & Space$(lngSizeRow1 - lngTableHwidHeader1 + 1) & "| " & _
                                            strTableHwidHeader2 & Space$(lngSizeRow2 - lngTableHwidHeader2 + 1) & "| " & _
                                            strTableHwidHeader3 & Space$(lngSizeRow3 - lngTableHwidHeader3 + 1) & "| " & _
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
        End If

        ' Сохраняем подсказку в массив, если требуется
        If mbSaveToolTip Then
            If mbFirstStart Then
                ReDim Preserve arrTTip(Index)
                'arrTTip(index) = strTTipText
                arrTTip(Index) = strTTipTextOnlyDrivers
            Else
                DebugMode str2VbTab & "ReadOrSaveToolTip: ToolTipArrIndex=" & Index & ":" & UBound(arrTTip)
                DebugMode str2VbTab & "ReadOrSaveToolTip: strTTipText=" & strTTipText
                arrTTip(Index) = strTTipText
            End If
        End If

        TT.Tools.Add acmdPackFiles(Index).hWnd, , strTTipText, True

        TimeScriptFinish = GetTickCount
        DebugMode str2VbTab & "ReadOrSaveToolTip - End - Time to Read Driverpack's - " & strPackFileName & ": " & CalculateTime(TimeScriptRun, TimeScriptFinish, True), 1
    Else
        DebugMode str2VbTab & "ReadOrSaveToolTip: Empty Driverpack's Name"
    End If


End Sub

Private Function RunDPInst(ByVal strWorkPath As String) As Long

Dim cmdString                           As String

    DebugMode "RunDPInst-Start"
    DebugMode "RunDPInst: strWorkPath" & strWorkPath
    RunDPInst = 0
    cmdString = Kavichki & strDPInstExePath & Kavichki & " " & CollectCmdString & "/PATH " & Kavichki & strWorkPath & Kavichki
    ChangeStatusTextAndDebug strMessages(93)

    If RunAndWaitNew(cmdString, PathNameFromPath(strDPInstExePath), vbNormalFocus) = False Then
        If Not mbSilentRun Then
            MsgBox strMessages(21) & str2vbNewLine & cmdString, vbInformation, strProductName
        End If

        ChangeStatusTextAndDebug strMessages(21) & " " & cmdString, "Error on run : " & cmdString
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
                ChangeStatusTextAndDebug strMessages(96) & " " & cmdString
            End If
        End If
    End If

    DebugMode "RunDPInst-End"

End Function

' Выделение все пакеты драйверов на текущей вкладке
Private Sub SelectAllOnTabDP(Optional ByVal mbIntellectMode As Boolean = True)

    If SSTab1.Enabled Then
        'MsgBox "Выбираем нужный режим установки"
        Sleep 100

        If mbIntellectMode Then
            SelectStartMode 1, False
        Else
            SelectStartMode 2, False
        End If

        cmbCheckButton.ListIndex = 3
        cmbCheckButton.Refresh
        DoEvents
        Sleep 200
        cmdCheck_Click
    End If

End Sub

Private Sub SelectNextTab()

Dim lng2Tab                             As Long

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

' Выделение пакетов c неустановленными драйверами
Private Sub SelectNotInstalledDP(Optional ByVal mbIntellectMode As Boolean = True)

    If SSTab1.Enabled Then
        'MsgBox "Выбираем нужный режим установки"
        Sleep 100

        If mbIntellectMode Then
            SelectStartMode 1, False
        Else
            SelectStartMode 2, False
        End If

        cmbCheckButton.ListIndex = 1
        cmbCheckButton.Refresh
        DoEvents
        Sleep 200
        cmdCheck_Click
    End If

End Sub

' Выделение пакетов рекомендованных к установке
Private Sub SelectRecommendedDP(Optional ByVal mbIntellectMode As Boolean = True)

    If SSTab1.Enabled Then
        'MsgBox "Выбираем нужный режим установки"
        Sleep 100

        If mbIntellectMode Then
            SelectStartMode 1, False
        Else
            SelectStartMode 2, False
        End If

        'MsgBox "Выбираем всё рекомендованное для установки"
        cmbCheckButton.ListIndex = 2
        cmbCheckButton.Refresh
        DoEvents
        Sleep 200
        cmdCheck_Click
    End If

End Sub

Private Function SetFirstEnableTab() As Long

Dim i                                   As Long

    For i = 0 To SSTab1.Tabs - 1

        If SSTab1.TabVisible(i) Then
            If SSTab1.TabEnabled(i) Then
                SetFirstEnableTab = i
                Exit For
            End If
        End If

    Next

End Function

Private Sub SetScrollFramePos(ByVal sgnNum As Single, ByVal LngValue As Long, ByVal lngCntTab As Long)
Dim i                                   As Integer
Dim SSTabHeight                         As Long
Dim SSTabTabHeight                      As Long
Dim miValue3                            As Long
Dim lngControlHeight                    As Long
Dim lngControlWidth                     As Long
    
    SSTabTabHeight = SSTab1.TabHeight
    SSTabHeight = SSTab1.Height
    miValue3 = frRunChecked.Left + frRunChecked.Width - 50
    
    For i = SSTab2.LBound To SSTab2.UBound

        With SSTab2(i)

            If Not (SSTab2.Item(i) Is Nothing) Then
                '.Visible = False

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

Private Sub SetStartScrollFramePos(ByVal miUnHideTabTemp As Integer)

Dim cntUnHideTab                        As Long
Dim miValue1                            As Long
Dim miValue2                            As Long
Dim sngNum1                             As Single
Dim sngNum2                             As Single

    If mbTabHide Then
        cntUnHideTab = miUnHideTabTemp + 1
        sngNum1 = cntUnHideTab / lngOSCountPerRow
        miValue1 = Round(sngNum1, 0)

        If cntUnHideTab > 0 Then
            SetScrollFramePos sngNum1, miValue1, cntUnHideTab
        End If

    Else
        sngNum2 = lngOSCount / lngOSCountPerRow
        miValue2 = Round(sngNum2, 0)
        
        SetScrollFramePos sngNum2, miValue2, lngOSCount

    End If

End Sub

Private Sub SetTabProperties()

    With SSTab1
        .Font.Name = strDialogTab_FontName
        .Font.Size = miDialogTab_FontSize
        .Font.Underline = mbDialogTab_Underline
        .Font.Strikethrough = mbDialogTab_Strikethru
        .Font.Bold = mbDialogTab_Bold
        .Font.Italic = mbDialogTab_Italic
        .ForeColor = lngDialogTab_Color
        .Font.Charset = lngDialog_Charset
    End With

End Sub

Private Sub SetTabPropertiesTabDrivers()

'Сохранение визуально заданых свойств шрифтов в переменных
    If mbFirstStart Then
        With SSTab2(0)
            .Font.Name = strDialogTab2_FontName
            .Font.Size = miDialogTab2_FontSize
            .Font.Underline = mbDialogTab2_Underline
            .Font.Strikethrough = mbDialogTab2_Strikethru
            .Font.Bold = mbDialogTab2_Bold
            .Font.Italic = mbDialogTab2_Italic
            .ForeColor = lngDialogTab2_Color
            .Font.Charset = lngDialog_Charset
        End With
    Else
        Dim i                           As Long

        With SSTab2
            For i = .LBound To .UBound
                With .Item(i)
                    .Font.Name = strDialogTab2_FontName
                    .Font.Size = miDialogTab2_FontSize
                    .Font.Underline = mbDialogTab2_Underline
                    .Font.Strikethrough = mbDialogTab2_Strikethru
                    .Font.Bold = mbDialogTab2_Bold
                    .Font.Italic = mbDialogTab2_Italic
                    .ForeColor = lngDialogTab2_Color
                    .Font.Charset = lngDialog_Charset
                End With
            Next
        End With
    End If
End Sub

'! -----------------------------------------------------------
'!  Функция     :  SetTabsNameAndCurrTab
'!  Переменные  :
'!  Описание    :  Назначить имена для вкладок и установить текущую на основании версии ОС
'! -----------------------------------------------------------
Private Sub SetTabsNameAndCurrTab(ByVal mbSecondStart As Boolean)

Dim i                                   As Long
Dim i_i                                 As Long
Dim miTabIndex                          As Long
Dim miFirstTabIndex                     As Long
Dim miSymbol                            As Long
Dim strTabIndex                         As String
Dim strTabIndex_x()                     As String
Dim strTabIndexTemp                     As String
Dim StrTabName                          As String
Dim str_x64                             As String
Dim lngSupportedOS                      As Long

    DebugMode "SetTabsNameAndCurrTab-Start"
    lngSupportedOS = 0

    For i = 0 To UBound(arrOSList)
        StrTabName = arrOSList(i).Name
        str_x64 = arrOSList(i).is64bit

        If InStr(arrOSList(i).Ver, strOsCurrentVersion) Then

            ' Если в списке есть ОС x64
            If str_x64 = 1 Then
                If InStr(StrTabName, "64") = 0 Then
                    StrTabName = StrTabName & " x64"
                End If
            End If

            If str_x64 = "2" Or str_x64 = "3" Then
                miTabIndex = i
                strTabIndex = IIf(LenB(strTabIndex) > 0, strTabIndex & " ", vbNullString) & CStr(miTabIndex)
                lngSupportedOS = lngSupportedOS + 1
            Else

                If CBool(str_x64) = mbIsWin64 Then
                    miTabIndex = i
                    strTabIndex = IIf(LenB(strTabIndex) > 0, strTabIndex & " ", vbNullString) & CStr(miTabIndex)
                    lngSupportedOS = lngSupportedOS + 1
                End If
            End If
        End If

        SSTab1.TabCaption(i) = StrTabName
    Next

    'Если среди вкладок не найдено поддержки вашей ОС
    If lngSupportedOS = 0 Then
        mbNoSupportedOS = True
    End If

    miSymbol = InStr(strTabIndex, " ")

    If miSymbol > 0 Then
        strTabIndexTemp = Trim$(Left$(strTabIndex, miSymbol))
        miFirstTabIndex = CInt(strTabIndexTemp)
    Else
        miFirstTabIndex = miTabIndex
    End If

    If mbSecondStart Then
        strTabIndex_x = Split(strTabIndex, " ")

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
    DebugMode vbTab & "SetTabsNameAndCurrTab: SetCurrentTabOSList=" & strTabIndex
    DebugMode vbTab & "SetTabsNameAndCurrTab: SetCurrentTab=" & miFirstTabIndex
    DebugMode "SetTabsNameAndCurrTab-End"

End Sub

Private Sub SetTabsWidth(ByVal miUnHideTabTemp As Integer)

Dim cntUnHideTab                        As Integer
Dim miValue                             As Integer

    If mbTabHide Then
        cntUnHideTab = miUnHideTabTemp + 1
        miValue = frRunChecked.Left + frRunChecked.Width - 50

        With SSTab1

            If cntUnHideTab > 0 Then
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

'Сценарий запуска тихой установки
Private Sub SilentCheckNoDB()
    Sleep 200
    SelectStartMode 3, False
    
    'Выбираем всё рекомендованное для установки
    cmbCheckButton.ListIndex = 3
    cmbCheckButton.Refresh
    DoEvents
    Sleep 200
    cmdCheck_Click
    
    'Собственно запускаем сам процесс создания БД
    mbGroupTask = True
    mbooSelectInstall = False
    Sleep 200
    cmdRunTask_Click
    FindNoDBCount
    mbGroupTask = False

End Sub

'Сценарий запуска тихой установки
Private Sub SilentInstall()
'Команда для программы DPInst работать в тихом режиме
    mbDpInstQuietInstall = True
    ' для работы в тихом режиме надо обязательно отключать promt
    mbDpInstPromptIfDriverIsNotBetter = False
    ' Включение отладочного режима
    mbDebugEnable = True
    DebugMode "SilentInstall-Start"
    'MsgBox "Выбираем нужный режим установки"
    DebugMode vbTab & "SilentInstall: SelectMode: " & strSilentSelectMode

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
    mbooSelectInstall = False
    Sleep 200
    GroupInstallDP
    mbGroupTask = False
    DebugMode "SilentInstall-End"

End Sub

'! -----------------------------------------------------------
'!  Функция     :  SSTab1_Click
'!  Переменные  :  PreviousTab As Integer
'!  Описание    :
'! -----------------------------------------------------------
Private Sub SSTab1_Click(PreviousTab As Integer)
    TabStopParam

    If acmdPackFiles(0).Visible Then
        If acmdPackFiles.UBound > 1 Then
            mbNextTab = True
        End If
    End If

    If ctlScrollControl1.UBound >= SSTab1.Tab Then
        If arrOSList(SSTab1.Tab).CntBtn > 0 Then
            ctlScrollControl1(SSTab1.Tab).Refresh
        End If
    End If

    If optRezim_Upd.Value Then
        ' если активна вкладка c 1-3, то тогда в этом режиме переставляем на 0
        If SSTab2(SSTab1.Tab).Tab > 0 Then
            If SSTab2(SSTab1.Tab).Tab < 4 Then
                SSTab2(SSTab1.Tab).Tab = 0
            End If
        End If
    End If
End Sub

' Нажатие кнопки на SStab2
Private Sub SSTab2_Click(Index As Integer, PreviousTab As Integer)
    If SSTab2(Index).Tab = 0 Then
        If PreviousTab > 0 Then
            ctlScrollControl1(Index).Visible = False
        End If
    End If
    
    StartReOrderBtnOnTab2 Index, PreviousTab
    
    If SSTab2(Index).Tab = 0 Then
        If PreviousTab > 0 Then
            If ctlScrollControl1(Index).Visible = False Then
                ctlScrollControl1(Index).Visible = True
            End If
        End If
    End If
End Sub

' Запуск перестроение кнопок на активной вкладке
Private Sub StartReOrderBtnOnTab2(ByVal miIndex As Integer, ByVal miPrevTab As Integer)
Dim lngCntBtnTab                        As Long
Dim lngCntBtnPrevious                   As Long

    If Not mbFirstStart Then
        lblNoDP4Mode.Visible = False
        lngCntBtnTab = arrOSList(miIndex).CntBtn - 1
        
        If lngCntBtnTab >= 0 Then
            If miIndex > 0 Then
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
            Sleep 100
    
            Select Case SSTab2(miIndex).Tab
                
                ' Построение пакетов со всеми драйверами (возврат всех кнопок на место)
                Case 0
                    If miPrevTab > 0 Then
                        ReOrderBtnOnTab2 miIndex, SSTab2(miIndex).Tab, lngCntBtnPrevious, lngCntBtnTab, ctlScrollControl1(miIndex)
                    End If
    
                ' Построение пакетов с новыми драйверами
                Case 1
                    ReOrderBtnOnTab2 miIndex, SSTab2(miIndex).Tab, lngCntBtnPrevious, lngCntBtnTab, ctlScrollControlTab1(miIndex)
    
                ' Построение пакетов с неустановленными драйверами
                Case 2
                    ReOrderBtnOnTab2 miIndex, SSTab2(miIndex).Tab, lngCntBtnPrevious, lngCntBtnTab, ctlScrollControlTab2(miIndex)
                    
                ' Построение пакетов с установленными драйверами
                Case 3
                    ReOrderBtnOnTab2 miIndex, SSTab2(miIndex).Tab, lngCntBtnPrevious, lngCntBtnTab, ctlScrollControlTab3(miIndex)
    
                ' Построение пакетов с "БД не создана"
                Case 4
                    ' Переключаемся в режим создания БД
                    mbSet2UpdateFromTab4 = True
                    SelectStartMode 3, False
                    
                    ReOrderBtnOnTab2 miIndex, SSTab2(miIndex).Tab, lngCntBtnPrevious, lngCntBtnTab, ctlScrollControlTab4(miIndex)
    
                    mbSet2UpdateFromTab4 = False
                    
            End Select
        End If
    End If
End Sub

' Запуск перестроение кнопок на определенной вкладке
Private Sub ReOrderBtnOnTab2(ByVal lngTab2Index As Long, ByVal lngTab2Tab As Long, ByVal lngBtnPrevCnt As Long, ByVal lngBtnTabCnt As Long, objScrollControl As Object)
Dim i                                 As Long
Dim lngStartPosLeft                   As Long
Dim lngStartPosTop                    As Long
Dim lngNextPosLeft                    As Long
Dim lngNextPosTop                     As Long
Dim lngMaxLeftPos                     As Long
Dim lngDeltaPosLeft                   As Long
Dim lngDeltaPosTop                    As Long
Dim lngBtnPrevNum                     As Long
Dim lngNoDP4ModeCnt                   As Long

    lngStartPosLeft = miButtonLeft
    lngStartPosTop = miButtonTop
    lngBtnPrevNum = 0
    lngNoDP4ModeCnt = 0
    
    'Debug.Print objScrollControl.Index
    objScrollControl.Visible = False
    For i = lngBtnPrevCnt To lngBtnTabCnt

        If Not (acmdPackFiles(i).Picture Is Nothing) Then
            Select Case lngTab2Tab
                Case 1
                    If acmdPackFiles(i).Picture = imgOkNew.Picture Or acmdPackFiles(i).Picture = imgOkAttentionNew.Picture Then
                        GoTo MoveBtn
                    Else
                        GoTo NextBtn
                    End If
                Case 2
                    If acmdPackFiles(i).Picture = imgOkAttention.Picture Or acmdPackFiles(i).Picture = imgOkAttentionOLD.Picture Or acmdPackFiles(i).Picture = imgOkAttentionNew.Picture Then
                        GoTo MoveBtn
                    Else
                        GoTo NextBtn
                    End If
                Case 3
                    If acmdPackFiles(i).Picture = imgOK.Picture Or acmdPackFiles(i).Picture = imgOkAttentionOLD.Picture Or acmdPackFiles(i).Picture = imgOkAttentionNew.Picture Or acmdPackFiles(i).Picture = imgOkNew.Picture Or acmdPackFiles(i).Picture = imgOkOld.Picture Then
                        GoTo MoveBtn
                    Else
                        GoTo NextBtn
                    End If
                Case 4
                    If acmdPackFiles(i).Picture = imgNoDB.Picture Then
                        GoTo MoveBtn
                    Else
                        GoTo NextBtn
                    End If
            End Select

MoveBtn:
            ' Собственно перемещаем кнопки на другую вкладку
            Set acmdPackFiles(i).Container = objScrollControl
            Set chkPackFiles(i).Container = objScrollControl
            
            ' положения кнопок
            If i = 0 Then
                lngNextPosLeft = lngStartPosLeft
                lngNextPosTop = lngStartPosTop
            Else

                If lngBtnPrevNum > 0 Then
                    lngDeltaPosLeft = acmdPackFiles(lngBtnPrevNum).Left + miButtonWidth + miBtn2BtnLeft - lngStartPosLeft
                Else
                    ' Если первая кнопка подходит, то расчитываем следующее положение исходя из нее
                    If lngTab2Tab > 0 Then
                        If InStr(1, acmdPackFiles(0).Container.Name, "ctlScrollControlTab", vbTextCompare) Then
                            lngDeltaPosLeft = acmdPackFiles(0).Left + miButtonWidth + miBtn2BtnLeft - lngStartPosLeft
                        End If
                    Else
                        lngDeltaPosLeft = acmdPackFiles(0).Left + miButtonWidth + miBtn2BtnLeft - lngStartPosLeft
                    End If
                End If

                lngNextPosLeft = lngStartPosLeft + lngDeltaPosLeft
                lngMaxLeftPos = lngNextPosLeft + miButtonWidth + 25

                If lngMaxLeftPos > objScrollControl.Width Then
                    ' Если по горизонтали кнопка не входит, то перешагиваем
                    lngDeltaPosLeft = 0
                    lngDeltaPosTop = lngDeltaPosTop + miButtonHeight + miBtn2BtnTop
                    lngNextPosLeft = lngStartPosLeft
                    lngNextPosTop = lngStartPosTop + lngDeltaPosTop
                Else
                    lngNextPosTop = lngStartPosTop + lngDeltaPosTop
                End If
            End If
            
            ' Перемещение кнопок и checkbox по расчитанным ранее параметрам
            acmdPackFiles(i).Move lngNextPosLeft, lngNextPosTop
            chkPackFiles(i).Move (lngNextPosLeft + 50), (lngNextPosTop + (miButtonHeight - chkPackFiles(i).Height) / 2)
            chkPackFiles(i).ZOrder 0
            
            ' Увеличиваем счетчики
            lngBtnPrevNum = i
            lngNoDP4ModeCnt = lngNoDP4ModeCnt + 1
NextBtn:
        End If
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
    'objScrollControl.SetFocus

End Sub

Private Sub TabInstBlockOnUpdate(ByVal mbBlock As Boolean)

Dim i                                   As Long

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

Private Sub TabStopParam()

Dim i                                   As Long
Dim lngCntBtnTab                        As Long
Dim lngCntBtnPrevious                   As Long
Dim lngSSTab1Tab                        As Long

    With SSTab1

        lngSSTab1Tab = .Tab

        If lngSSTab1Tab > 0 Then
            lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 1).CntBtn

            If lngCntBtnPrevious = 0 Then
                If lngSSTab1Tab > 1 Then
                    lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 2).CntBtn
                    If lngCntBtnPrevious = 0 Then
                        If lngSSTab1Tab > 2 Then
                            lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 3).CntBtn
                            If lngCntBtnPrevious = 0 Then
                                If lngSSTab1Tab > 3 Then
                                    lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 4).CntBtn
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With

    lngCntBtnTab = arrOSList(lngSSTab1Tab).CntBtn - 1

    For i = acmdPackFiles.LBound To acmdPackFiles.UBound

        With acmdPackFiles(i)
            .TabStop = i >= lngCntBtnPrevious And i <= lngCntBtnTab
        End With

    Next

End Sub

'! -----------------------------------------------------------
'!  Функция     :  ToolTipStatusLoad
'!  Переменные  :
'!  Описание    :  Загрузка статусных соощений
'! -----------------------------------------------------------
Private Sub ToolTipStatusLoad()

Dim arrTTipStatusIconTemp()            As String

    ReDim arrTTipStatusIcon(8) As String
    ReDim arrTTipStatusIconTemp(8) As String

    DebugMode "ToolTipStatusLoad-Start"
    arrTTipStatusIconTemp(0) = "В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере." & str2vbNewLine & "Такие же драйвера (тех же версий) уже установлены на вашем компьютере." & str2vbNewLine & "Ваши действия:" & vbNewLine & "Никаких действий не требуется. " & str2vbNewLine & "Примечание:" & vbNewLine & "Если в колонке статус для устроуства стоит '0', то это означает:" & vbNewLine & " * - устройство блокировано;" & vbNewLine & " * - драйвер для данного устройства не активен (см. сведения в диспетчере устройств)"
    arrTTipStatusIconTemp(1) = "В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере." & vbNewLine & _
                               vbNewLine & _
                               "На вашем компьютере эти драйвера не установлены." & vbNewLine & _
                               vbNewLine & _
                               "Ваши действия:" & vbNewLine & _
                               "Переключите программу в один из режимов установки драйверов и нажмите на эту кнопку - это приведет к установке необходимых драйверов из пакета и соответствующему изменению вида кнопки." & vbNewLine & _
                               vbNewLine & _
                               "Примечания:" & vbNewLine & _
                               "1. В некоторых случаях драйвера, из обнаруженных в пакете драйверов, могут не подойти к вашему оборудованию. Сравнение идентификаторов устройств (HWID) происходит по основной части HWID без учета подклассов устройств (SUBSYS|REV|MI)." & vbNewLine & _
                               "2. Если в колонке статус для устроуства стоит '0', то это означает:" & vbNewLine & _
                               " * - драйвер для данного устройства не установлен;" & vbNewLine & _
                               " * - устройство блокировано;" & vbNewLine & _
                               " * - драйвер для данного устройства не активен (см. сведения в диспетчере устройств)"
    arrTTipStatusIconTemp(2) = "В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере, но более новые, чем те, что уже установлены." & str2vbNewLine & "Ваши действия:" & vbNewLine & "Переключите программу в один из режимов установки драйверов и нажмите на эту кнопку - это приведет к установке более новых драйверов из пакета и соответствующему изменению вида кнопки." & str2vbNewLine & "Примечание:" & vbNewLine & "В некоторых случаях драйвера, из обнаруженных в пакете драйверов, могут не подойти к вашему оборудованию. Сравнение идентификаторов устройств (HWID) происходит по основной части HWID без учета подклассов устройств (SUBSYS|REV|MI)."
    arrTTipStatusIconTemp(3) = "В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере, но более старые, чем те, что уже установлены." & str2vbNewLine & "Ваши действия:" & vbNewLine & "Ничего делать не надо. Можете поискать в сети более свежие драйвера и обновить (заменить) данный пакет в программе."
    arrTTipStatusIconTemp(4) = "1. В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере, но более новые, чем те, что уже установлены." & vbNewLine & _
                               vbNewLine & _
                               "Ваши действия:" & vbNewLine & _
                               "Переключите программу в один из режимов установки драйверов и нажмите на эту кнопку - это приведет к установке более новых драйверов из пакета и соответствующему изменению вида кнопки." & vbNewLine & _
                               vbNewLine & _
                               "2. В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере." & vbNewLine & _
                               vbNewLine & _
                               "На вашем компьютере эти драйвера не установлены." & vbNewLine & _
                               vbNewLine & _
                               "Ваши действия:" & vbNewLine & _
                               "Переключите программу в один из режимов установки драйверов и нажмите на эту кнопку - это приведет к установке необходимых драйверов из пакета и соответствующему изменению вида кнопки." & vbNewLine & _
                               vbNewLine & _
                               "Примечания:" & vbNewLine & _
                               "1. В некоторых случаях драйвера, из обнаруженных в пакете драйверов, могут не подойти к вашему оборудованию. Сравнение идентификаторов устройств (HWID) происходит по основной части HWID без учета подклассов устройств (SUBSYS|REV|MI)." & vbNewLine & _
                               "2. Если в колонке статус для устроуства стоит '0', то это означает:" & vbNewLine & _
                               " * - драйвер для данного устройства не установлен;" & vbNewLine & _
                               " * - устройство блокировано;" & vbNewLine & _
    " * - драйвер для данного устройства не активен (см. сведения в диспетчере устройств)"
    arrTTipStatusIconTemp(5) = "1. В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере, но более старые, чем те, что уже установлены." & vbNewLine & _
                               vbNewLine & _
                               "Ваши действия:" & vbNewLine & _
                               "Ничего делать не надо. Можете поискать в сети более свежие драйвера и обновить (заменить) данный пакет в программе." & vbNewLine & _
                               vbNewLine & _
                               "2. В этом пакете драйверов программы есть драйвера, подходящие к устройству, обнаруженному на вашем компьютере." & vbNewLine & _
                               vbNewLine & _
                               "На вашем компьютере эти драйвера не установлены." & vbNewLine & _
                               vbNewLine & _
                               "Ваши действия:" & vbNewLine & _
                               "Переключите программу в один из режимов установки драйверов и нажмите на эту кнопку - это приведет к установке необходимых драйверов из пакета и соответствующему изменению вида кнопки." & vbNewLine & _
                               vbNewLine & _
                               "Примечания:" & vbNewLine & _
                               "1. В некоторых случаях драйвера, из обнаруженных в пакете драйверов, могут не подойти к вашему оборудованию. Сравнение идентификаторов устройств (HWID) происходит по основной части HWID без учета подклассов устройств (SUBSYS|REV|MI)." & vbNewLine & _
                               "2. Если в колонке статус для устроуства стоит '0', то это означает:" & vbNewLine & _
                               " * - драйвер для данного устройства не установлен;" & vbNewLine & _
                               " * - устройство блокировано;" & vbNewLine & _
                               " * - драйвер для данного устройства не активен (см. сведения в диспетчере устройств)"
    arrTTipStatusIconTemp(6) = "Драйвера из этого пакета программы не имеют отношения к вашему компьютеру." & str2vbNewLine & "Ваши действия:" & vbNewLine & "Ничего делать не надо. Этот пакет драйверов пригодится вам как-нибудь в другой раз - при замене устройств или на другом компьютере."
    arrTTipStatusIconTemp(7) = "Программа не может определить, что находится в этом пакете драйверов." & str2vbNewLine & "Ваши действия:" & vbNewLine & "Переключите программу в режим 'Создание или обновление базы данных драйверов', нажмите на эту кнопку - таким образом сведения о драйверах из пакета будут добавлены в базу данных программы и вид кнопки изменится. О дальнейших действиях читайте в пояснении к соответствующему значку."
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
        If .Tools.Count > 0 Then
            .Tools.Clear
        End If
        .Font.Name = strMainForm_FontName
        .Font.Size = lngMainForm_FontSize
        .MaxTipWidth = frDescriptionIco.Width
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

    DebugMode "ToolTipStatusLoad-End"

End Sub

' Перезагрузка всплывающих подсказок для кнопок с драйверами
Private Sub ToolTipBtnReLoad()
    DebugMode str2VbTab & "ReloadToolTip-Start"
    
    'Если подсказки уже созданы, то очистка
    If TT.Tools.Count > 0 Then
        TT.Tools.Clear
        TT.Title = strTTipTextTitle
    End If

    ' Обновляем всплывающую подсказку
    UpdateStatusButtonAll True
    DebugMode str2VbTab & "ReloadToolTip-End"
End Sub

'! -----------------------------------------------------------
'!  Функция     :  UnPackDPFile
'!  Переменные  :  strPathDRP As String, strPackFileName As String, StrMaskFile As String
'!  Описание    :  Извлечение файлов из архива
'! -----------------------------------------------------------
Private Function UnPackDPFile(ByVal strPathDRP As String, _
                              ByVal strPackFileName As String, _
                              ByVal strMaskFile As String, _
                              ByVal strDest4OnlyUnpack As String) As Boolean

Dim WorkDir                             As String
Dim strPackFileName_woExt               As String
Dim cmdString                           As String
Dim ArchTempPath                        As String
Dim strPhysXPath                        As String
Dim strLangPath                         As String
Dim strRuntimes                         As String
Dim strClassesName                      As String
Dim strFinishIniPath                    As String
Dim ret                                 As Long
Dim strMaskFile_x()                     As String
Dim i                                   As Long
Dim strMaskFile_x_TEMP                  As String
Dim strMaskFile_x_TEMPTo                As String
Dim strMaskFile_xx()                    As String

    DebugMode "UnPackDPFile-Start"
    DebugMode "UnPackDPFile: strMaskFile=" & strMaskFile

    If Not mbOnlyUnpackDP Then
        strPackFileName_woExt = FileName_woExt(strPackFileName)

        'Рабочий каталог
        If mbGroupTask Then
            WorkDir = strWorkTempBackSL & "GroupInstall\"
            ArchTempPath = strWorkTempBackSL & "GroupInstall"
        Else
            WorkDir = BackslashAdd2Path(strWorkTempBackSL & strPackFileName_woExt)
            ArchTempPath = strWorkTempBackSL & strPackFileName_woExt

            If PathExists(WorkDir) Then
                DelRecursiveFolder (WorkDir)
            End If
        End If

    Else
        ArchTempPath = strDest4OnlyUnpack
    End If

    If Not mbDP_Is_aFolder Then
        cmdString = Kavichki & strArh7zExePATH & Kavichki & " x -yo" & Kavichki & ArchTempPath & Kavichki & " -r " & Kavichki & strPathDRP & strPackFileName & Kavichki & " " & strMaskFile
        ChangeStatusTextAndDebug strMessages(97) & " " & strPackFileName, "Extract: " & cmdString
        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
            If Not mbSilentRun Then
                MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
            End If

            UnPackDPFile = False
            ChangeStatusTextAndDebug strMessages(13) & " " & strPackFileName, "Error on run : " & cmdString
        Else

            'Распаковка дополнительных файлов
            ' Если класс пакета считывается при запуске программы, то
            ' Архиватор отработал на все 100%? Если нет то сообщаем
            If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
                ChangeStatusTextAndDebug strMessages(13) & " " & strPackFileName

                If Not mbSilentRun Then
                    MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                End If

            Else
                strClassesName = vbNullString

                If mbReadClasses Then
                    strFinishIniPath = PathCombine(arrOSList(SSTab1.Tab).devIDFolderFull, FileName_woExt(strPackFileName) & ".ini")
                    strClassesName = IniStringPrivate("DriverPack", "classes", strFinishIniPath)

                    ' Если такого значения в файле нет, то ничего не добавляем
                    If strClassesName = "no_key" Then
                        strClassesName = vbNullString
                    End If

                    If LenB(strClassesName) > 0 Then

                        ' Если класс пакета определен как Display, то
                        If StrComp(strClassesName, "Display", vbTextCompare) = 0 Then
                            If Not mbGroupTask Then

                                ' Распаковка strPhysXPath
                                If LenB(arrOSList(SSTab1.Tab).PathPhysX) > 0 Then
                                    strPhysXPath = PathCollect(arrOSList(SSTab1.Tab).PathPhysX)
                                    UnPackDPFileAdd strPhysXPath, strPathDRP, ArchTempPath
                                End If

                                ' Распаковка strLangPath
                                If LenB(arrOSList(SSTab1.Tab).PathLanguages) > 0 Then
                                    strLangPath = PathCollect(arrOSList(SSTab1.Tab).PathLanguages)
                                    UnPackDPFileAdd strLangPath, strPathDRP, ArchTempPath
                                End If

                                ' Распаковка strRuntimes
                                If LenB(arrOSList(SSTab1.Tab).PathRuntimes) > 0 Then
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

        ChangeStatusTextAndDebug strMessages(149) & " " & strPackFileName, "Copy: " & strMaskFile
        If PathExists(WorkDir) = False Then
            CreateNewDirectory WorkDir
        End If
        If InStr(strMaskFile, " ") Then
            strMaskFile_x = Split(strMaskFile, " ")
            For i = LBound(strMaskFile_x) To UBound(strMaskFile_x)
                strMaskFile_x_TEMP = BacklashDelFromPath(strMaskFile_x(i))
                strMaskFile_xx = Split(strMaskFile_x_TEMP, vbBackslash)
                If UBound(strMaskFile_xx) > 1 Then
                    strMaskFile_x_TEMPTo = Left$(strMaskFile_x_TEMP, InStrRev(strMaskFile_x_TEMP, vbBackslash) - 1)
                End If
                ret = ret + CopyFolderByShell(BackslashAdd2Path(strPathDRP & strPackFileName) & strMaskFile_x_TEMP, BackslashAdd2Path(ArchTempPath) & strMaskFile_x_TEMPTo)
            Next
        Else
            strMaskFile_x_TEMP = BacklashDelFromPath(strMaskFile)
            strMaskFile_xx = Split(strMaskFile_x_TEMP, vbBackslash)
            If UBound(strMaskFile_xx) > 1 Then
                strMaskFile_x_TEMPTo = Left$(strMaskFile_x_TEMP, InStrRev(strMaskFile_x_TEMP, vbBackslash) - 1)
            End If
            ret = CopyFolderByShell(BackslashAdd2Path(strPathDRP & strPackFileName) & strMaskFile, BackslashAdd2Path(ArchTempPath) & strMaskFile_x_TEMPTo)
        End If
        UnPackDPFile = Not Abs(ret)
        DebugMode "UnPackDPFile-Copy files: " & UnPackDPFile
    End If
    DebugMode "UnPackDPFile-End"

End Function

Private Sub UnPackDPFileAdd(ByVal strPathAddFile As String, _
                            ByVal strPathDRP As String, _
                            ByVal strArchTempPath As String)

Dim cmdString                           As String

    If InStr(strPathAddFile, vbBackslash) = 0 Then
        strPathAddFile = BackslashAdd2Path(strPathDRP) & strPathAddFile
    End If

    If PathExists(strPathAddFile) = False Then
        If Not PathIsAFolder(strPathAddFile) Then
            strPathAddFile = SearchFilesInRoot(PathNameFromPath(strPathAddFile), FileNameFromPath(strPathAddFile), False, True)
        End If
    End If

    If PathExists(strPathAddFile) Then
        If Not PathIsAFolder(strPathAddFile) Then
            cmdString = Kavichki & strArh7zExePATH & Kavichki & " x -yo" & Kavichki & strArchTempPath & Kavichki & " -r " & Kavichki & strPathAddFile & Kavichki & " *.*"
            ChangeStatusTextAndDebug strMessages(98) & " " & strPathAddFile, "Extract: " & cmdString

            If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
                If Not mbSilentRun Then
                    MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                End If

                ChangeStatusTextAndDebug strMessages(13) & " " & strPathAddFile, "Error on run : " & cmdString
            Else

                ' Архиватор отработал на все 100%? Если нет то сообщаем
                If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
                    ChangeStatusTextAndDebug strMessages(13) & " " & strPathAddFile

                    If Not mbSilentRun Then
                        MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Function UnpackOtherFile(ByVal strArcDRPPath As String, _
                                 ByVal strWorkDir As String, _
                                 ByVal strMaskFile As String) As Boolean

Dim cmdString                           As String

    DebugMode "UnpackOtherFile-Start"
    DebugMode "UnpackOtherFile: strArcDRPPath=" & strArcDRPPath
    DebugMode "UnpackOtherFile: strMaskFile=" & strMaskFile
    cmdString = Kavichki & strArh7zExePATH & Kavichki & " x -yo" & Kavichki & strWorkDir & Kavichki & " -r " & Kavichki & strArcDRPPath & Kavichki & " " & strMaskFile
    ChangeStatusTextAndDebug strMessages(99) & " " & strArcDRPPath, "Extract: " & cmdString
    UnpackOtherFile = True

    If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
        If Not mbSilentRun Then
            MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
        End If

        ChangeStatusTextAndDebug strMessages(13) & " " & cmdString, "Error on run : " & cmdString
        UnpackOtherFile = False
    Else

        ' Архиватор отработал на все 100%? Если нет то сообщаем
        If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
            ChangeStatusTextAndDebug strMessages(13) & " " & FileNameFromPath(strArcDRPPath)

            If Not mbSilentRun Then
                MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
            End If

            UnpackOtherFile = False
        End If
    End If

    DebugMode "UnpackOtherFile-End"

End Function

'! -----------------------------------------------------------
'!  Функция     :  UpdateStatusButtonAll
'!  Переменные  :
'!  Описание    :  Обновление всех статусов
'! -----------------------------------------------------------
Public Sub UpdateStatusButtonAll(Optional mbReloadTT As Boolean = False)

Dim ButtIndex                           As Long
Dim ButtCount                           As Long
Dim i                                   As Integer
Dim i_Tab                               As Integer
Dim TimeScriptRun                       As Long
Dim TimeScriptFinish                    As Long
Dim AllTimeScriptRun                    As String
Dim miPbInterval                        As Long
Dim miPbNext                            As Long
Dim mbDpNoDBExist                       As Boolean
Dim lngSStabStart                       As Long
Dim strPackFileName                     As String
Dim strPathDRP                          As String
Dim strPathDevDB                        As String
Dim lngTabN                             As Long
Dim lngNumButtOnTab                     As Long

    DebugMode "StatusUpdateAll-Start"
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
        If LenB(chkPackFiles(0).Tag) > 0 Then
            i_Tab = chkPackFiles(0).Tag
        End If
    End If

    BlockControl False
    Sleep 100
    DoEvents
    SSTab1.Tab = i_Tab
    TimeScriptRun = 0
    AllTimeScriptRun = vbNullString
    TimeScriptRun = GetTickCount
    ButtIndex = acmdPackFiles.UBound
    ButtCount = acmdPackFiles.Count
    ' Отображаем ProgressBar
    CreateProgressNew

    If ButtIndex > 0 Then
        ' В цикле обрабатываем обновление
        miPbInterval = 1000 / ButtCount
        miPbNext = 0

        For i = 0 To ButtIndex
            lngTabN = SSTab1.Tab

            'If LenB(arrOSList(lngTabN).CntBtn) > 0 Then
            lngNumButtOnTab = arrOSList(SSTab1.Tab).CntBtn

            Do While i >= lngNumButtOnTab
                lngTabN = lngTabN + 1
                SSTab1.Tab = lngTabN
                DoEvents
                Sleep 100

                'If LenB(arrOSList(SSTab1.Tab).CntBtn) > 0 Then
                lngNumButtOnTab = arrOSList(SSTab1.Tab).CntBtn
                'Else
                'lngNumButtOnTab = 0
                'End If

            Loop
            'Else
            '                lngTabN = lngTabN + 1
            '                SSTab1.Tab = lngTabN
            '                DoEvents
            '                Sleep 100
            'End If

            mbDpNoDBExist = True
            strPathDRP = arrOSList(lngTabN).drpFolderFull
            strPathDevDB = arrOSList(lngTabN).devIDFolderFull

            With acmdPackFiles(i)
                If Not mbReloadTT Then
                    ' Кнопка выглядит нажатой
                    Set .Picture = imgUpdBD.Picture
                    FlatBorderButton .hWnd
                    .Refresh
                    strPackFileName = .Tag
                    ChangeStatusTextAndDebug "(" & i + 1 & " " & strMessages(124) & " " & ButtCount & "): " & strMessages(89) & " " & strPackFileName
                    ' Обновление подсказки
                    ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, i
                    ' Кнопка выглядит отжатой
                    FlatBorderButton .hWnd, False
                    .Refresh
                Else
                    strPackFileName = .Tag
                    ' Обновление подсказки
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
        ChangeStatusTextAndDebug strMessages(67) & " " & AllTimeScriptRun
    Else
        ChangeStatusTextAndDebug strMessages(68)
    End If

    pbProgressBar.Visible = False
    ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
    
    ChangeFrmMainCaption
    BlockControl True
    
TheEnd:
    SSTab1.Tab = lngSStabStart
    DebugMode "StatusUpdateAll-End"

End Sub

'! -----------------------------------------------------------
'!  Функция     :  UpdateStatusButtonTAB
'!  Переменные  :
'!  Описание    :  Обновление всех статусов
'! -----------------------------------------------------------
Public Sub UpdateStatusButtonTAB()

Dim i                                   As Integer
Dim TimeScriptRun                       As Long
Dim TimeScriptFinish                    As Long
Dim AllTimeScriptRun                    As String
Dim miPbInterval                        As Long
Dim miPbNext                            As Long
Dim mbDpNoDBExist                       As Boolean
Dim strPackFileName                     As String
Dim strPathDRP                          As String
Dim strPathDevDB                        As String
Dim lngCntBtnTab                        As Long
Dim lngCntBtnPrevious                   As Long
Dim lngSSTab1Tab                        As Long
Dim lngCurrBtn                          As Long
Dim lngSummBtn                          As Long

    DebugMode "UpdateStatusButtonTAB-Start"
    BlockControl False
    ctlUcStatusBar1.PanelText(1) = strMessages(127)
    
    Sleep 100
    DoEvents
    AllTimeScriptRun = vbNullString
    TimeScriptRun = GetTickCount
    ' Отображаем ProgressBar
    CreateProgressNew

    With SSTab1

        lngSSTab1Tab = .Tab

        If lngSSTab1Tab > 0 Then
            lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 1).CntBtn

            If lngCntBtnPrevious = 0 Then
                If lngSSTab1Tab > 1 Then
                    lngCntBtnPrevious = arrOSList(lngSSTab1Tab - 2).CntBtn
                End If
            End If
        End If
    End With

    lngCntBtnTab = arrOSList(lngSSTab1Tab).CntBtn - 1


    If lngCntBtnTab > 0 Then
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
                Set .Picture = imgUpdBD.Picture
                FlatBorderButton .hWnd
                .Refresh
                strPackFileName = .Tag
                ChangeStatusTextAndDebug "(" & lngCurrBtn & " " & strMessages(124) & " " & lngSummBtn & "): " & strMessages(89) & " " & strPackFileName
                ' Обновление подсказки
                ReadOrSaveToolTip strPathDevDB, strPathDRP, strPackFileName, i
                ' Кнопка выглядит отжатой
                FlatBorderButton .hWnd, False
                .Refresh
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
        ChangeStatusTextAndDebug strMessages(67) & " " & AllTimeScriptRun
    Else
        ChangeStatusTextAndDebug strMessages(68)
    End If

    ChangeFrmMainCaption
    pbProgressBar.Visible = False
    ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
    
    BlockControl True
TheEnd:
    DebugMode "UpdateStatusButtonTAB-End"

End Sub

'! -----------------------------------------------------------
'!  Функция     :  VerModules
'!  Переменные  :
'!  Описание    :  Отображение версий модулей
'! -----------------------------------------------------------
Private Sub VerModules()
    MsgBox strMessages(35) & vbNewLine & "DPinst.exe (x86)" & vbTab & objFSO.GetFileVersion(strDPInstExePath) & vbNewLine & "DPinst.exe (x64)" & vbTab & objFSO.GetFileVersion(strDPInstExePath64) & vbNewLine & "DevCon.exe (x86)" & vbTab & objFSO.GetFileVersion(strDevConExePath) & vbNewLine & "7za.exe (x86)" & vbTab & objFSO.GetFileVersion(strArh7zExePATH), vbInformation, strProductName

End Sub

Private Sub WorkWithFinish(ByVal strPathDRP As String, _
                           ByVal strPackFileName As String, _
                           ByVal strWorkPath As String, _
                           ByVal strPathDRPList As String)

Dim StrPathDRPList_x()                  As String
Dim strSectionName                      As String
Dim strFinishIniPath                    As String
Dim lngEXCCount                         As Long
Dim i                                   As Long
Dim ii                                  As Long

    DebugMode "WorkWithFinish-Start"

    If mbLoadFinishFile Then
        If strPathDRPList <> ALL_FILES Then
            StrPathDRPList_x = Split(strPathDRPList, " ")

            For ii = LBound(StrPathDRPList_x) To UBound(StrPathDRPList_x)
                strSectionName = FileNameFromPath(BacklashDelFromPath(StrPathDRPList_x(ii)))
                ChangeStatusTextAndDebug strMessages(100) & " '" & strSectionName & "'"
                strFinishIniPath = PathCombine(arrOSList(SSTab1.Tab).devIDFolderFull, FileName_woExt(strPackFileName) & ".ini")

                If PathExists(strFinishIniPath) Then
                    If Not PathIsAFolder(strFinishIniPath) Then
                        lngEXCCount = IniLongPrivate(strSectionName, "exc_count", strFinishIniPath)

                        ' Если такого значения в файле нет, то ничего не добавляем
                        If lngEXCCount = "9999" Then
                            lngEXCCount = 0
                        End If

                        If lngEXCCount > 0 Then

                            For i = 1 To lngEXCCount
                                FindAndInstallPanel strPathDRP & strPackFileName, strFinishIniPath, strSectionName, i, strWorkPath
                            Next
                        End If
                    End If
                End If

            Next
        End If
    End If

    DebugMode "WorkWithFinish-End"

End Sub

' Выбор стартового режима работы программы
Private Sub SelectStartMode(Optional miModeTemp As Long = 0, _
                            Optional mbTab2 As Boolean = True)

Dim i_i                                 As Long
Dim miMode                              As Long

    ' Если указан параметр miModeTemp значит это переклбчени вкладок не при старте программы
    If miModeTemp > 0 Then
        miMode = miModeTemp
    Else
        miMode = miStartMode
    End If

    DebugMode "Start Rezim: " & miMode

    ' Режим при старте
    Select Case miMode

        Case 1

            If optRezim_Intellect.Enabled Then
                optRezim_Upd.Value = False
                optRezim_Intellect.Value = False
                optRezim_Intellect.Value = True
                optRezim_Intellect_Click
            Else
                optRezim_Ust.Value = False
                optRezim_Intellect.Value = False
                optRezim_Upd.Value = True
                optRezim_Upd_Click
            End If

        Case 2

            If optRezim_Ust.Enabled Then
                optRezim_Upd.Value = False
                optRezim_Intellect.Value = False
                optRezim_Ust.Value = True
                optRezim_Ust_Click
            Else
                optRezim_Ust.Value = False
                optRezim_Intellect.Value = False
                optRezim_Upd.Value = True
                optRezim_Upd_Click
            End If

        Case 3
            optRezim_Ust.Value = False
            optRezim_Intellect.Value = False
            optRezim_Upd.Value = True
            optRezim_Upd_Click

    End Select

    ' выставляем вторую вкладку только при старте программы
    If mbTab2 Then
        If miMode <> 3 Then
            If lngStartModeTab2 > 0 Then
                For i_i = SSTab2.LBound To SSTab2.UBound
                    ' Если вкладка активна, то выставляем начальную
                    If SSTab2(i_i).TabEnabled(lngStartModeTab2) = True Then
                        SSTab2(i_i).Tab = lngStartModeTab2
                    Else
                        SSTab2(i_i).Tab = 0
                    End If
                Next
            End If
        End If
    End If

End Sub

' Функция проверяет есть ли искомый текст в источнике посредством RegEXP
Private Function CheckExistbyRegExp(ByVal strSourceText As String, _
                                    ByVal strSearchText As String, _
                                    Optional ByVal mbGetText As Boolean, _
                                    Optional ByRef strFindText As String) As Boolean

Dim objRegExpCheck                      As RegExp
Dim objMatchesCheck                     As MatchCollection

    Set objRegExpCheck = New RegExp

    With objRegExpCheck
        .Pattern = strSearchText
        .IgnoreCase = True
        Set objMatchesCheck = .Execute(strSourceText)
    End With

    CheckExistbyRegExp = objMatchesCheck.Count

    If mbGetText Then
        If CheckExistbyRegExp Then
            strFindText = Trim$(objMatchesCheck.Item(0).Value)
        End If
    End If

    ' Очистка переменных
    Set objRegExpCheck = Nothing
    Set objMatchesCheck = Nothing

End Function

Private Sub mnuContextCopyHWID2Clipboard_Click(Index As Integer)

Dim strValue                            As String
Dim strValueDevID                       As String
Dim strValueDevID_x()                   As String

    strValue = mnuContextDeleteDevID(Index).Caption
    strValueDevID = Left$(strValue, InStr(strValue, vbTab) - 1)

    If InStr(strValueDevID, vbBackslash) Then
        strValueDevID_x = Split(strValueDevID, vbBackslash)
        strValueDevID = strValueDevID_x(0) & vbBackslash & strValueDevID_x(1)
    End If

    ' Копируем текст в клипбоард
    CBSetText strValueDevID

End Sub

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

Private Sub DelDuplicateOldDP()

Dim ButtIndex                           As Long
Dim i                                   As Long
Dim ii                                  As Long
Dim strPackFileName()                   As String
Dim strPackFileNames                    As String
Dim strPackFileName_woVersion           As String
Dim strPackFileNameTemp                 As String
Dim lngVersionPosition                  As Long
Dim strPackFileName_Ext                 As String
Dim objRegExp                           As RegExp
Dim objMatch                            As Match
Dim objMatches                          As MatchCollection
Dim strVerDP_1                          As String
Dim strVerDP_2                          As String
Dim strVerDP_1_1                        As String
Dim strVerDP_2_1                        As String
Dim strDPName_1                         As String
Dim strDPName_2                         As String
Dim strVerDP_Main                       As String
Dim strResult                           As String
Dim strResult1                          As String
Dim strResult2                          As String
Dim strPackFileName2Del                 As String
Dim strPackFileName2DelTemp             As String
Dim strPackFileName2Del_x()             As String
Dim lngMsgRet                           As Long
Dim lngStrLen1                          As Long
Dim lngStrLen2                          As Long

    ButtIndex = acmdPackFiles.UBound
    ReDim strPackFileName(ButtIndex, 2)

    If ButtIndex > 0 Then

        For i = 0 To ButtIndex
            strPackFileName(i, 0) = acmdPackFiles(i).Tag
            strPackFileName(i, 1) = i
            If LenB(strPackFileNames) > 0 Then
                strPackFileNames = strPackFileNames & ";" & acmdPackFiles(i).Tag
            Else
                strPackFileNames = acmdPackFiles(i).Tag
            End If
        Next
    End If

    For i = LBound(strPackFileName, 1) To UBound(strPackFileName, 1)
        strPackFileNameTemp = strPackFileName(i, 0)
        If InStr(strPackFileNameTemp, vbBackslash) Then
            strPackFileNameTemp = FileNameFromPath(strPackFileName(i, 0))
        End If
        lngVersionPosition = InStrRev(strPackFileNameTemp, "_", , vbTextCompare)
        If lngVersionPosition > 0 Then
            strPackFileName_woVersion = Left$(strPackFileNameTemp, lngVersionPosition)
            strPackFileName_Ext = ExtFromFileName(strPackFileNameTemp)

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
                            Set objMatch = .Item(ii)
                            strVerDP_1 = Trim$(objMatch.SubMatches(1))
                            strDPName_1 = Trim$(objMatch.SubMatches(0))
                            strVerDP_1_1 = Trim$(objMatch.SubMatches(2))
                            Set objMatch = Nothing
                        Else
                            strVerDP_1 = strVerDP_Main
                            strDPName_1 = strDPName_2
                        End If

                        Set objMatch = .Item(ii + 1)
                        strVerDP_2 = Trim$(objMatch.SubMatches(1))
                        strDPName_2 = Trim$(objMatch.SubMatches(0))

                        strVerDP_2_1 = Trim$(objMatch.SubMatches(2))

                        lngStrLen1 = Len(strVerDP_1)
                        lngStrLen2 = Len(strVerDP_2)

                        If lngStrLen1 > lngStrLen2 Then
                            strResult1 = CompareByVersion(Left$(strVerDP_1, lngStrLen2), strVerDP_2)
                            If strResult1 = "=" Then
                                strResult = strResult1
                            Else
                                strResult = strResult1
                            End If
                        ElseIf lngStrLen1 < lngStrLen2 Then
                            strResult1 = CompareByVersion(strVerDP_1, Left$(strVerDP_2, lngStrLen1))
                            If strResult1 = "=" Then
                                strResult = strResult1
                            Else
                                strResult = strResult1
                            End If
                        Else
                            strResult = CompareByVersion(strVerDP_1, strVerDP_2)
                            If strResult = "=" Then
                                If LenB(strVerDP_1_1) > 0 And LenB(strVerDP_1_1) > 0 Then
                                    strResult2 = CompareByVersion(strVerDP_1_1, strVerDP_2_1)
                                End If
                                strResult = strResult2
                            End If
                        End If

                        If strResult = ">" Then
                            strVerDP_Main = strVerDP_1
                            strPackFileName2DelTemp = strDPName_2
                        ElseIf strResult = "<" Then
                            strVerDP_Main = strVerDP_2
                            strPackFileName2DelTemp = strDPName_1
                        End If

                        If LenB(strPackFileName2Del) > 0 Then
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
    If LenB(strPackFileName2Del) > 0 Then
        If ShowMsbBoxForm(strPackFileName2Del, strMessages(139), strMessages(29)) = vbYes Then
            strPackFileName2Del_x = Split(strPackFileName2Del, vbNewLine)
            For i = LBound(strPackFileName2Del_x) To UBound(strPackFileName2Del_x)
                strPackFileName2DelTemp = strPackFileName2Del_x(i)
                For ii = 0 To ButtIndex
                    If StrComp(strPackFileName2DelTemp, acmdPackFiles(ii).Tag, vbTextCompare) = 0 Then
                        CurrentSelButtonIndex = ii
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

Private Sub mnuDelDuplicateOldDP_Click()
    DelDuplicateOldDP

    If mbRestartProgram Then
        ShellExecute Me.hWnd, "open", App.EXEName, vbNullString, strAppPath, SW_SHOWNORMAL
        Unload Me
        End
    End If

End Sub

Private Sub FontCharsetChange()
' Выставляем шрифт
    With Me.Font
        .Name = strMainForm_FontName
        .Size = lngMainForm_FontSize
        .Charset = lngDialog_Charset
    End With

    frCheck.Font.Charset = lngDialog_Charset
    frDescriptionIco.Font.Charset = lngDialog_Charset
    frInfo.Font.Charset = lngDialog_Charset
    frRezim.Font.Charset = lngDialog_Charset
    frRunChecked.Font.Charset = lngDialog_Charset
    frTabPanel.Font.Charset = lngDialog_Charset
    ctlUcStatusBar1.Font.Charset = lngDialog_Charset
    
    SetButtonProperties cmdRunTask
    SetButtonProperties cmdBreakUpdateDB

End Sub

Private Function IsFormLoaded(FormName As String) As Boolean
Dim i                                   As Integer
    
    For i = 0 To Forms.Count - 1
        If Forms(i).Name = FormName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next i
    IsFormLoaded = False
End Function

