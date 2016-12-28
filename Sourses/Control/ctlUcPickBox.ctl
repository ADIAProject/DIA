VERSION 5.00
Begin VB.UserControl ctlUcPickBox 
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ScaleHeight     =   2055
   ScaleWidth      =   2175
   ToolboxBitmap   =   "ctlUcPickBox.ctx":0000
   Begin VB.CommandButton cmdDrop 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   720
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Click Here to View Selected Files."
      Top             =   720
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.CommandButton cmdPick 
      Caption         =   "..."
      Height          =   275
      Left            =   1155
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   275
   End
   Begin VB.ComboBox cmbMultiSel 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox pbPick 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      Picture         =   "ctlUcPickBox.ctx":0312
      ScaleHeight     =   285
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pbDrop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      Picture         =   "ctlUcPickBox.ctx":0717
      ScaleHeight     =   285
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      MaxLength       =   65384
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Locate Folder..."
      Top             =   0
      Width           =   1455
   End
   Begin VB.Shape ShapeBorder 
      BorderColor     =   &H00B99D7F&
      Height          =   735
      Left            =   1560
      Top             =   0
      Width           =   495
   End
   Begin VB.Image imMetallicDrop 
      Height          =   285
      Index           =   3
      Left            =   1800
      Picture         =   "ctlUcPickBox.ctx":0B19
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imHomeSteadDrop 
      Height          =   285
      Index           =   3
      Left            =   1800
      Picture         =   "ctlUcPickBox.ctx":0EDE
      Top             =   1380
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imBlueDrop 
      Height          =   285
      Index           =   3
      Left            =   1800
      Picture         =   "ctlUcPickBox.ctx":12A3
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imMetallicDrop 
      Height          =   285
      Index           =   2
      Left            =   1560
      Picture         =   "ctlUcPickBox.ctx":1668
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imHomeSteadDrop 
      Height          =   285
      Index           =   2
      Left            =   1560
      Picture         =   "ctlUcPickBox.ctx":1A56
      Top             =   1380
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imBlueDrop 
      Height          =   285
      Index           =   2
      Left            =   1560
      Picture         =   "ctlUcPickBox.ctx":1E49
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imMetallicDrop 
      Height          =   285
      Index           =   1
      Left            =   1320
      Picture         =   "ctlUcPickBox.ctx":2276
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imHomeSteadDrop 
      Height          =   285
      Index           =   1
      Left            =   1320
      Picture         =   "ctlUcPickBox.ctx":2656
      Top             =   1380
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imBlueDrop 
      Height          =   285
      Index           =   1
      Left            =   1320
      Picture         =   "ctlUcPickBox.ctx":2A3D
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imMetallicDrop 
      Height          =   285
      Index           =   0
      Left            =   1080
      Picture         =   "ctlUcPickBox.ctx":2E5A
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imHomeSteadDrop 
      Height          =   285
      Index           =   0
      Left            =   1080
      Picture         =   "ctlUcPickBox.ctx":324B
      Top             =   1380
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imBlueDrop 
      Height          =   285
      Index           =   0
      Left            =   1080
      Picture         =   "ctlUcPickBox.ctx":364F
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imMetallic 
      Height          =   285
      Index           =   3
      Left            =   720
      Picture         =   "ctlUcPickBox.ctx":3A51
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imHomeStead 
      Height          =   285
      Index           =   3
      Left            =   720
      Picture         =   "ctlUcPickBox.ctx":3E15
      Top             =   1380
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imBlue 
      Height          =   285
      Index           =   3
      Left            =   720
      Picture         =   "ctlUcPickBox.ctx":41D6
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imMetallic 
      Height          =   285
      Index           =   2
      Left            =   480
      Picture         =   "ctlUcPickBox.ctx":4597
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imHomeStead 
      Height          =   285
      Index           =   2
      Left            =   480
      Picture         =   "ctlUcPickBox.ctx":4979
      Top             =   1380
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imBlue 
      Height          =   285
      Index           =   2
      Left            =   480
      Picture         =   "ctlUcPickBox.ctx":4D74
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imMetallic 
      Height          =   285
      Index           =   1
      Left            =   240
      Picture         =   "ctlUcPickBox.ctx":51A4
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imHomeStead 
      Height          =   285
      Index           =   1
      Left            =   240
      Picture         =   "ctlUcPickBox.ctx":558D
      Top             =   1380
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imBlue 
      Height          =   285
      Index           =   1
      Left            =   240
      Picture         =   "ctlUcPickBox.ctx":5977
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imMetallic 
      Height          =   285
      Index           =   0
      Left            =   0
      Picture         =   "ctlUcPickBox.ctx":5D9C
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imHomeStead 
      Height          =   285
      Index           =   0
      Left            =   0
      Picture         =   "ctlUcPickBox.ctx":6182
      Top             =   1380
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imBlue 
      Height          =   285
      Index           =   0
      Left            =   0
      Picture         =   "ctlUcPickBox.ctx":6586
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "ctlUcPickBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+  File Description:
'       ucPickBox - Enhanced File Picker Control
'
'   Product Name:
'       ucPickBox.ctl
'
'   Compatability:
'       Windows: 98, ME, NT, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'       (Common Dialog API Calls - Paul Mather)
'           URL: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=3592&lngWId=1
'       (TrimPathLen Function - Wastingtape)
'           URL: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=23456&lngWId=1
'       (FileExists - Eric Russell)
'           URL: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=829&lngWId=1
'       (ComboBox Open/Visible - Francesco Balena)
'           URL: http://www.devx.com/vb2themax/Tip/18336
'       (Max Raskin - Flat Button)
'           http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=6517&lngWId=1
'       (BrowseForFolder - DaVBMan, MrBobo)
'           http://vbcity.com/forums/topic.asp?tid=82667
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=22387&lngWId=1
'       (Randy Birch - IsWinXP)
'           http://vbnet.mvps.org/code/system/getversionex.htm
'       (Dieter Otter - GetCurrentThemeName)
'           http://www.vbarchiv.net/archiv/tipp_805.html
'
'   Legal Copyright & Trademarks:
'       Copyright © 2006, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2006, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this  software. This software is owned by Paul R. Territo, Ph.D and is
'       sold for use as a license in accordance with the terms of the License
'       Agreement in the accompanying the documentation.
'
'       Many thanks to my friend Paul Turcksin for his careful review, suggestions,
'       and support of this UserControl and TestHarness prior to public release. In
'       addtion, I wish to thank the numerous open source authors who provide code
'       and inspiration to make such work possible.
'
'   Contact Information:
'       For Technical Assistance:
'       Email: pwterrito@insightbb.com
'
'-  Modification(s) History:
'       05Nov05 - Initial TestHarness and UserControl finished
'       06Nov05 - Cleaned up bugs in the ShowSave and ShowOpen routines.
'               - Consolidated calls for the Show Open/Save subs to make
'                 param and error handling cleaner.
'               - Added addtional API params to the ShowFont routine.
'               - Updated the ToolBox Image to a more professional image.
'               - Added addtional error handling to the TestHarness...
'       19Nov05 - Added Additional Author Credits to the Header
'               - Added UseDialogColor, UseDialogText, ForeColor, and
'                 BackColor properties to the Control and required code to
'                 allow these routines to work...
'               - Added PrintStatusMsg property to allow the user to specify
'                 what the message should say when the printer returns a value.
'               - Added PrintStaus property to provide the user feedback about
'                 if the Printer dialog "Ok"(1) or "Cancel"(0) button was pressed.
'               - Fixed bug in ShowSave routine which inconssistently computes the
'                 nFileOffset values for a file. We simply set this to "0" and then
'                 extract the values from outside of this of this routine.
'               - Changes Color from Long to OLE_COLOR property to allow for
'                 vb stanard palette.
'               - Added TranslateColor sub to wrap the OleTranslateColor method
'                 for mapping of colors to the current RGB palette.
'       20Nov05 - Added Color RollBack if the value entered is invalid.
'       04Dec05 - Changed the TestHarness layout to make it easier to follow the
'                 flow of the controls and how to use it....
'       06Dec05 - Added MultiFile selection for the ShowOpen routine and fixed several
'                 bugs with the single vs mutiple file selections.
'               - Added a ComboBox to serve at the conatiner and windowing mechanism for
'                 the list and its events....this is a hack, pure and simple. This
'                 approach was selected as it allowes a floating window and list functionality
'                 without the need for building this via API. The combobox is hidden
'                 behind the textbox at runtime and has Visiable = False. Since we
'                 call the droplist window via SendMessage this allows us to have a
'                 floating window like the ComboBox, but none of the overhead to manage ;-D
'               - Add the ability to programmatically Open the MultiFile ComboBox
'                 and check the state of the Droplist.
'               - Added cmdDrop button to simulate the drop button of the ComboBox. The
'                 key feature here being that the button is to the left of the ellipes
'                 button and is resizable with the dialog, unlike the VB ComboBox.
'       13Dec05 - Fixed minor TestHareness bug which displayed the wrong properties when
'                 selecting the lstProperties index.
'       14Dec05 - Fixed single/multiple file open bug in the ShowOpen routine which caused the
'                 the sub to enter into the wrong conditional section when a single file
'                 was selected and the MultiSelect = False.
'               - Fixed PropertyChanged calls for DialogMsg and ToolTipTexts which now supports
'                 individual item settings.
'       15Dec05 - More optimization on the ShowSave and ShowFont routines. These routines now
'                 handle missing extensions and provide a mechanism to enter them. In addtion,
'                 the FontColor property has been added to allow direct color picking of the
'                 font ForeColor, which is not appart of the StdFont structure.
'       16Dec05 - Added Appearance Property and associated API and VB routines to allow for true
'                 3D or Flat appearances of the textbox and buttons.
'       18Dec05 - Fixed Minor bugs in the ShowFont dialog routines which did not preserve the
'                 previous selections by the user. The new addtions resolve all but one known
'                 bug. At the current time, the iPointSize of the FontDialog type structure is
'                 not correctly set via code and the dialog does not respond the changes in this
'                 parameter despite accounting for the size and weight of the font. Verified the
'                 ShowFont code against www.allapi.net example and neither resulted in the pointsize
'                 being selected. For more details see http://mentalis.org/apilist/CHOOSEFONT.shtml
'       25Dec05 - Added Events: DropClick, KeyDown, KeyPress, KeyUp, MouseDown, MouseMove, MouseUp.
'               - Added GetCursorPosition function to allow reporting of the Cursor position via
'                 GetCursorPosition and ScreenToClient API's regardless of which part of the control
'                 the cursor is over. This effectively bypasses the native Event Handlers for each
'                 control, and provides a uniform reporting of the cursor position on the control surface.
'               - Added additional documentation at the Method and Property levels to provide added
'                 clarity of what the functionality is...
'       26Dec05 - Added Filter Property and associated routines to the ShowOpen, ShowSave routines,
'                 see Filter Let property for correct format of the filter string....
'               - Added ProcessFilter to replace string Pipes (|) with vbNullChar and fix the
'                 final size of the passed string to the dialogs.
'               - Added error handling for none initialized Filters to read All Files (*.*)
'       27Dec05 - Added Color, Font, File, and PrinterFlags as Public Enums along with properties
'                 to allow the developer set the styles more easily.
'               - Added SHOWCOLOR_DEFAULT, SHOWFONT_DEFAULT, SHOWOPEN_DEFAULT, SHOWSAVE_DEFAULT,
'                 and SHOWPRINTER_DEFAULT custom Non-Win32 flags to allow for rapid dialog setting
'                 which encompass the most common flags used with this control.
'               - Updated the TestHarness in the UpdatePropertiesDialog to reflect these changes.
'       28Dec05 - Added UseAutoForeColor and associated routines to allow the developer to choose
'                 if the ForeColor is to be selected automatically. The value for the new ForeColor
'                 is based on the XOr of the BackColor and should always produce high contrast text
'                 in the dialog regardless of the color selected.
'       03Jan06 - Added BrowseForFolder functionality and associated routines to round out the collection
'                 based on the request from Richard Mewett.
'       07Mar06 - Added Let Property for Path to pass data to txtResult and m_Path parameter. The displayed
'                 Path is trimmed using the TrimPathLen routine.
'               - Fixed bug which causes the txtResult to display the incorrect message when ucFolder was the
'                 dialog type.
'       16Mar06 - Add Paul Caton's SelfSubclass Thunk code to allow for BrowseForFolder CallBack without the
'                 need for an external bas module. The long point (address) of the z_SubclassProc is held in
'                 in the sc_aSubData(0).nAddrSub provided this is the only item we are subclassing....if we are
'                 subclassing multiple items (i.e. Usercontrol, Parent) then the address for each is stored in
'                 order in the sc_aSubData(n).nAddrSub, where n = 0, 1....n
'       06Jun06 - Added Theme capability and associated routines to allow for XP Themes
'               - Added Theme Properties
'               - Removed Parent subclassing for ThemeChange and SystemColor change messages, because this
'                 caused the IDE to crash on close.
'               - Fixed minor bug in BorderStyle when controls are Flat and change to Theme
'       10Jun06 - Fixed Minor bug in the Refresh routine which did not set the Classic style
'                 correctly if the previous Apearance = Flat
'               - Added LockWindowUpdate to prevent flicker on Picture changes
'       28Jun06 - Fixed TrimPathByLen to be Printer Object independent
'       15Jul06 - Fixed TrackMouse missing Subclaser Code
'       16Jul06 - Fixed Missing IsWinXP routine in GetThemeInfo Method
'       29Jun07 - Fixed bug in ShowSave and ShowOpen dialogs which did not process the default extensions
'               - Added DefaultExt property to allow the developer to set the default extension to
'                 use in the Open/Save Dailogs
'       08Aug07 - Fixed Bug in the BFF section which did not correctly Qualify Paths.
'
'       Recode Control By Romeo91 for Better Subsclassing and Unicode Support for File And Text
'       10Dec13 - Repaint Subsclass Code from SelfSub 2.1 Paul Caton - http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=64867&lngWId=1.
'                 Added Unicode Support for FileOperation Dialog
'                 Added Unicode Support for Text Properties
'       05Mar14   Change subsclass to class cSelfSubHookCallback
'
'   Force Declarations
'   Oroginal Build Date & Time: 8/8/2007 10:22:17 AM
Option Explicit

'   Private API Declarations
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (ByVal pOpenfilename As Long) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameW" (ByVal pOpenfilename As Long) As Long
Private Declare Function SetWindowLongA Lib "user32.dll" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hWndLock As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal Path As String) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hWnd As Long) As Long

'Private Const BFFM_SETSELECTIONA        As Long = (WM_USER + 102)
Private Const FILE_ATTRIBUTE_DIR = &H10

'   Appearance Costants
Public Enum pbAppearanceConstants
    [Flat] = &H0
    [3D] = &H1
End Enum

Public Enum pbThemeEnum
    [pbAuto] = &H0
    [pbClassic] = &H1
    [pbBlue] = &H2
    [pbHomeStead] = &H3
    [pbMetallic] = &H4
End Enum

Private Enum pbStateEnum
    [pbNormal] = &H0
    [pbHover] = &H1
    [pbDown] = &H2
    [pbDisabled] = &H3
End Enum

'   Flat Button API Constants
'   The button style BS_FLAT used to change a button to a Flat one
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE        As Long = -16
Private Const WS_CHILD           As Long = &H40000000

'   GWL_Style is the attribute we will use for changing the style of the button
'   To set the button as a child window and not as a self dependent window
'   Send Message Constants for ComboBoxes
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETDROPPEDSTATE = &H157

'   ShowOpen / ShowSave Flags
Public Enum OpenSaveDialogFlags
    ReadOnly = &H1
    OverwritePrompt = &H2
    HideReadOnly = &H4
    NoChangeDir = &H8
    ShowHelp = &H10
    EnableHook = &H20
    EnableTemplate = &H40
    EnableTemplateHandle = &H80
    NoValidate = &H100
    AllowMultiselect = &H200
    ExtensionDifferent = &H400
    PathMustExist = &H800
    FileMustExist = &H1000
    Createprompt = &H2000
    ShareAware = &H4000
    NoReadOnlyReturn = &H8000
    NoTestFileCreate = &H10000
    NoNetworkButton = &H20000
    NoLongNames = &H40000
    Explorer = &H80000
    LongNames = &H200000
    NoDeReferenceLinks = &H100000
    '   Custom Non-Win32 Flags Which Are A Combinations Of Flags
    ShowOpen_Default = Explorer Or LongNames Or Createprompt Or NoDeReferenceLinks Or HideReadOnly
    ShowSave_Default = Explorer Or LongNames Or OverwritePrompt Or HideReadOnly
End Enum

Public Enum ucDialogConstant
    [ucFolder] = &H0
    [ucOpen] = &H1
    [ucSave] = &H2
End Enum

Private Const str2vbNullChar     As String = vbNullChar & vbNullChar

Private Const MAX_PATH As Long = 4096    '260

Private Type OPENFILENAME
    nStructSize                         As Long
    hWndOwner                           As Long
    hInstance                           As Long
    sFilter                             As String
    sCustomFilter                       As String
    nCustFilterSize                     As Long
    nFilterIndex                        As Long
    sFile                               As String
    nFileSize                           As Long
    sFileTitle                          As String
    nTitleSize                          As Long
    sInitDir                            As String
    sDlgTitle                           As String
    flags                               As Long
    nFileOffset                         As Integer
    nFileExt                            As Integer
    sDefFileExt                         As String
    nCustDataSize                       As Long
    fnHook                              As Long
    sTemplateName                       As String
    pvReserved                          As Long
    dwReserved                          As Long
    FlagsEx                             As Long
End Type

Private Type SelectedFile
    nFilesSelected                      As Integer
    sFiles()                            As String
    sLastDirectory                      As String
    bCanceled                           As Boolean
End Type

Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Type POINTAPI
    X                                   As Long
    Y                                   As Long
End Type

'   Private Dialog Structure Definitions
Private FileDialog         As OPENFILENAME

'   Private UserControl Properties
Private m_Appearance       As pbAppearanceConstants
Private m_UseAutoForeColor As Boolean
Private m_BackColor        As OLE_COLOR
Private m_DefaultExt       As String
Private m_DialogMsg(2)     As String
Private m_DialogType       As ucDialogConstant
Private m_Enabled          As Boolean
Private m_FileCount        As Long
Private m_FileFlags        As OpenSaveDialogFlags
Private m_Filename()       As String
Private m_Filters          As String
Private m_Font             As StdFont
Private m_FontColor        As OLE_COLOR
Private m_Forecolor        As OLE_COLOR
Private m_Hwnd             As Long
Private m_MultiSelect      As Boolean
Private m_Path             As String
Private m_Pnt              As POINTAPI
Private m_PrevBackColor    As OLE_COLOR
Private m_PrevLoc          As POINTAPI
Private m_State            As pbStateEnum
Private m_ToolTipText(2)   As String
Private m_Theme            As pbThemeEnum
Private m_UseDialogColor   As Boolean
Private m_UseDialogText    As Boolean
Private m_Locked           As Boolean
Private m_QualifyPaths     As Boolean
Private m_bIsWinXpOrLater  As Boolean
Private Const Def_DialogMsgFile       As String = "Locate File..."
Private Const Def_DialogMsgFolder     As String = "Locate Folder..."
Private Const Def_ToolTipMsgFile      As String = "Click Here to Locate File"
Private Const Def_ToolTipMsgFiles     As String = "Click Here to Locate Files"
Private Const Def_ToolTipMsgFolder    As String = "Click Here to Locate Folder"

'*************************************************************
'   Public UserControl Events
'*************************************************************
Public Event Click()
Public Event DropClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PathChanged()

'*************************************************************
'   Windows Messages
'*************************************************************
Private Const WM_THEMECHANGED   As Long = &H31A
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_NCACTIVATE     As Long = &H86
Private Const WM_ACTIVATE       As Long = &H6
Private Const WM_SETCURSOR As Long = &H20
Private Const WM_SIZING         As Long = &H214
Private Const WM_NCPAINT        As Long = &H85
Private Const WM_MOVING         As Long = &H216
Private Const WM_EXITSIZEMOVE   As Long = &H232

'*************************************************************
'   TRACK MOUSE
'*************************************************************
Public Event MouseEnter()
Public Event MouseLeave()

Private Const WM_MOUSELEAVE     As Long = &H2A3
Private Const WM_MOUSEMOVE      As Long = &H200

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                              As Long
    dwFlags                             As TRACKMOUSEEVENT_FLAGS
    hWndTrack                           As Long
    dwHoverTime                         As Long
End Type

Private bTrack       As Boolean
Private bTrackUser32 As Boolean
Private bInCtrl      As Boolean

Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "comctl32.dll" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'*************************************************************
'   Subsclass
'*************************************************************
Private m_cSubclass                                    As cSelfSubHookCallback

Private Enum eParamUser
    exParentForm = 1
    exUserControl = 2
End Enum

'*************************************************************
'   Пока не знаю нужно или нет, добавил на всякий случай так как использую класс CommonDialog
'*************************************************************
Implements OLEGuids.IOleInPlaceActiveObjectVB

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Appearance
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Appearance() As pbAppearanceConstants
    Appearance = m_Appearance
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Appearance
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lNewValue (pbAppearanceConstants)
'!--------------------------------------------------------------------------------
Public Property Let Appearance(lNewValue As pbAppearanceConstants)
    '   Store the Value
    m_Appearance = lNewValue
    '   Set the TextBox Style
    txtResult.Appearance = lNewValue
    '   Set the new visual styles to the passed type (3D or Flat)
    Call ButtonAppearance(cmdPick, lNewValue)
    Call ButtonAppearance(cmdDrop, lNewValue)
    '   We need to set the Visible state to False, since the
    '   ButtonAppearance function sets it to True as part of
    '   the window refresh mechanism
    cmdDrop.Visible = False
    '   Now call the resize, as the button position and sizes
    '   are changed when the border style changes...
    Call UserControl_Resize
    PropertyChanged "Appearance"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lNewColor (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let BackColor(ByVal lNewColor As OLE_COLOR)
    m_BackColor = lNewColor
    m_PrevBackColor = lNewColor
    '   Set the BackColor
    UserControl.txtResult.BackColor = lNewColor
    PropertyChanged "BackColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property DefaultExt
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get DefaultExt() As String
    DefaultExt = m_DefaultExt
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property DefaultExt
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (String)
'!--------------------------------------------------------------------------------
Public Property Let DefaultExt(ByVal NewValue As String)

    If LenB(NewValue) Then
        If AscW(NewValue) <> vbDot Then
            NewValue = "." & NewValue
        End If
    End If
    
    m_DefaultExt = NewValue
    PropertyChanged "DefaultExt"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property DialogMsg
'! Description (Описание)  :   [Get the Dialg Textbox Message for the Type selected]
'! Parameters  (Переменные):   lType (ucDialogConstant)
'!--------------------------------------------------------------------------------
Public Property Get DialogMsg(ByVal lType As ucDialogConstant) As String

    DialogMsg = m_DialogMsg(lType)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property DialogMsg
'! Description (Описание)  :   [ Set the Dialog Textbox Message for the Type selected]
'! Parameters  (Переменные):   lType (ucDialogConstant)
'                              sNewValue (String)
'!--------------------------------------------------------------------------------
Public Property Let DialogMsg(ByVal lType As ucDialogConstant, ByVal sNewValue As String)

    If lType < 0 Then lType = 0
    If lType > 2 Then lType = 2
    m_DialogMsg(lType) = sNewValue

    '   Store the changes for later
    Select Case lType

        Case ucFolder
            PropertyChanged "DialogMsg0"

        Case ucOpen
            PropertyChanged "DialogMsg1"

        Case ucSave
            PropertyChanged "DialogMsg2"
    End Select

    Call Refresh(0)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property DialogType
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get DialogType() As ucDialogConstant
    DialogType = m_DialogType
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property DialogType
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lType (ucDialogConstant)
'!--------------------------------------------------------------------------------
Public Property Let DialogType(ByVal lType As ucDialogConstant)

    '   Mkae sure the numbers are in range...
    If lType < 0 Then lType = 0
    If lType > 2 Then lType = 2
    '   Use our new dialog style...
    m_DialogType = lType

    With UserControl
        '   Reset the MutliSelect Drop Button and List
        .cmdDrop.Visible = False
        .pbDrop.Visible = False
        .cmbMultiSel.Clear
    End With

    PropertyChanged "DialogType"
    Call Refresh(0)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Enabled
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Enabled
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   bNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let Enabled(bNewValue As Boolean)
    m_Enabled = bNewValue

    '   As it name implys....
    With UserControl
        .Enabled = bNewValue
        .txtResult.Enabled = bNewValue
        .cmdPick.Enabled = bNewValue
        .cmdDrop.Enabled = bNewValue

        If m_Enabled = True Then
            m_State = pbNormal
        Else
            m_State = pbDisabled
        End If

        Call Refresh(0)
        Call Refresh(1)
    End With

    PropertyChanged "Enabled"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property FileCount
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get FileCount() As Long
    FileCount = m_FileCount
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property FileCount
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lNewCount (Long)
'!--------------------------------------------------------------------------------
Private Property Let FileCount(lNewCount As Long)
    '   The number of files in the MultSelect Mode of the ShowOpen
    m_FileCount = lNewCount
    PropertyChanged "FileCount"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property FileFlags
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get FileFlags() As OpenSaveDialogFlags
    FileFlags = m_FileFlags
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property FileFlags
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sDialogFlags (OpenSaveDialogFlags)
'!--------------------------------------------------------------------------------
Public Property Let FileFlags(sDialogFlags As OpenSaveDialogFlags)
    m_FileFlags = sDialogFlags
    PropertyChanged "FileFlags"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property FileName
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long = 1)
'!--------------------------------------------------------------------------------
Public Property Get FileName(Optional Index As Long = 1) As String
    '   Get the stored data...(File + Path)
    FileName = m_Filename(Index)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Filters
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Filters() As String
    Filters = m_Filters
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Filters
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sFileFilters (String)
'!--------------------------------------------------------------------------------
Public Property Let Filters(sFileFilters As String)
    '   Pass the File Filter string'
    '   i.e. sFileFilters = "Supported Files|*.bmp;*.doc;*.jpg;*.rtf;*.txt;*.tif|Bitmap Files (*.bmp)|*.bmp|Mircosoft Word Files (*.doc)|*.doc|JPEG Files (*.jpg)|*.jpg|Rich Text Format Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt"
    m_Filters = sFileFilters
    PropertyChanged "Filters"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property FontColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get FontColor() As OLE_COLOR
    FontColor = m_FontColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property FontColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lNewColor (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let FontColor(ByVal lNewColor As OLE_COLOR)
    m_FontColor = lNewColor
    PropertyChanged "FontColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Font
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Font() As StdFont
    '   Get the stored data...
    Set Font = m_Font
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_Forecolor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lNewColor (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ForeColor(ByVal lNewColor As OLE_COLOR)
    m_Forecolor = lNewColor
    UserControl.txtResult.ForeColor = lNewColor
    PropertyChanged "ForeColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property hDC
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get hDC()
    hDC = UserControl.hDC
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property hWnd
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get hWnd()
    hWnd = UserControl.hWnd
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Locked
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Locked() As Boolean
    Locked = m_Locked
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Locked
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lNewLocked (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let Locked(ByVal lNewLocked As Boolean)
    m_Locked = lNewLocked
    '   Set the Locked
    UserControl.txtResult.Locked = m_Locked
    PropertyChanged "Locked"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MultiSelect
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get MultiSelect() As Boolean
    '   Get the MutliSelect Status....for ShowOpen
    MultiSelect = m_MultiSelect
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MultiSelect
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   bNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let MultiSelect(bNewValue As Boolean)
    '   Set the MutliSelect State of the Dialog...
    '   NOTE: This is only used for the ShowOpen dialog type.
    m_MultiSelect = bNewValue
    PropertyChanged "MultiSelect"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Path
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Path() As String
    Path = QualifyPath(m_Path)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Path
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sNewPath (String)
'!--------------------------------------------------------------------------------
Public Property Let Path(sNewPath As String)

    If m_QualifyPaths Then
        m_Path = QualifyPath(sNewPath)
        UserControl.txtResult.Text = m_Path
    Else
        m_Path = sNewPath
        UserControl.txtResult.Text = sNewPath
    End If

    PropertyChanged "Path"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property QualifyPaths
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get QualifyPaths() As Boolean
    QualifyPaths = m_QualifyPaths
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property QualifyPaths
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lNewQualifyPaths (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let QualifyPaths(ByVal lNewQualifyPaths As Boolean)
    m_QualifyPaths = lNewQualifyPaths
    PropertyChanged "QualifyPaths"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   :   Property RunMode
'! Description :   [Ambient.UserMode tells us whether the UC's container is in design mode or user mode/run-time.
'                               Unfortunately, this isn't supported in all containers.]
'                               http://www.vbforums.com/showthread.php?805711-VB6-UserControl-Ambient-UserMode-workaround&s=8dd326860cbc22bed07bd13f6959ca70
'! Parameters  :
'!--------------------------------------------------------------------------------
Public Property Get RunMode() As Boolean
    RunMode = True
    On Error Resume Next
    RunMode = Ambient.UserMode
    RunMode = Extender.Parent.RunMode
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Theme
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Theme() As pbThemeEnum
    Theme = m_Theme
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Theme
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_Theme (pbThemeEnum)
'!--------------------------------------------------------------------------------
Public Property Let Theme(ByVal New_Theme As pbThemeEnum)
    m_Theme = New_Theme
    UserControl_Resize
    PropertyChanged "Theme"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ToolTipTexts
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lType (ucDialogConstant)
'!--------------------------------------------------------------------------------
Public Property Get ToolTipTexts(ByVal lType As ucDialogConstant) As String
    '   Get the Dialg ToolTipText Message for the Type selected
    ToolTipTexts = m_ToolTipText(lType)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ToolTipTexts
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lType (ucDialogConstant)
'                              sNewValue (String)
'!--------------------------------------------------------------------------------
Public Property Let ToolTipTexts(ByVal lType As ucDialogConstant, ByVal sNewValue As String)
    '   Set the Dialg ToolTipText Message for the Type selected
    m_ToolTipText(lType) = sNewValue

    Select Case lType

        Case ucFolder
            PropertyChanged "ToolTipText0"

        Case ucOpen
            PropertyChanged "ToolTipText1"

        Case ucSave
            PropertyChanged "ToolTipText2"
    End Select

    Call Refresh(0)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseAutoForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get UseAutoForeColor() As Boolean
    '   Get if we want the forecolor to be set automatically
    '   via XOR in the textbox backcolor
    UseAutoForeColor = m_UseAutoForeColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseAutoForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   bNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let UseAutoForeColor(ByVal bNewValue As Boolean)
    '   Set if we want the forecolor to be set automatically
    '   via XOR in the textbox backcolor
    m_UseAutoForeColor = bNewValue
    PropertyChanged "UseAutoForeColor"
    Call Refresh(0)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseDialogColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get UseDialogColor() As Boolean
    '   Get if we want the color as textbox backcolor
    UseDialogColor = m_UseDialogColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseDialogColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   bNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let UseDialogColor(ByVal bNewValue As Boolean)
    '   Set if we want to use color as the backcolor
    m_UseDialogColor = bNewValue
    PropertyChanged "UseDialogColor"
    Call Refresh(0)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseDialogText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get UseDialogText() As Boolean
    '   Dispaly the dialog text?
    UseDialogText = m_UseDialogText
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseDialogText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   bNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let UseDialogText(ByVal bNewValue As Boolean)
    '   Set if the dialog is to be diaplayed
    '
    '   One might want to turn off the text if using color in the display
    m_UseDialogText = bNewValue
    PropertyChanged "UseDialogText"
    Call Refresh(0)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ButtonAppearance
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   cmdButton (CommandButton)
'                              lButtonStyle (pbAppearanceConstants)
'!--------------------------------------------------------------------------------
Private Function ButtonAppearance(cmdButton As CommandButton, lButtonStyle As pbAppearanceConstants)

    If lButtonStyle = [3D] Then
        '   Here is a small function to change button to 3D (Note the Missing "BS_FLAT" flag)
        SetWindowLongA cmdButton.hWnd, GWL_STYLE, WS_CHILD
        '   Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
        cmdButton.Visible = True
    Else
        '   Here is a small function to change button to flat:-
        SetWindowLongA cmdButton.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
        '   Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
        cmdButton.Visible = True
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ComboBoxListVisible
'! Description (Описание)  :   [Wrapper funtion to allow us to get the drop state of the ComboBox.....]
'! Parameters  (Переменные):   cbo (ComboBox)
'!--------------------------------------------------------------------------------
Private Function ComboBoxListVisible(cbo As ComboBox) As Boolean
    ComboBoxListVisible = SendMessage(cbo.hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0&)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ExtractFilename
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sFileName (Variant)
'!--------------------------------------------------------------------------------
Private Function ExtractFilename(ByVal sFileName) As String

    '   Extract the Path from the full filename...
    Dim lStrCnt As Long

    lStrCnt = InStrRev(sFileName, "\")

    If lStrCnt Then
        ExtractFilename = Mid$(sFileName, lStrCnt + 1)
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ExtractPath
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sFileName (Variant)
'!--------------------------------------------------------------------------------
Private Function ExtractPath(ByVal sFileName) As String

    '   Extract the Path from the full filename...
    Dim lStrCnt As Long

    lStrCnt = InStrRev(sFileName, "\")

    If lStrCnt Then
        ExtractPath = Left$(sFileName, lStrCnt - 1)
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetCursorPosition
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function GetCursorPosition() As POINTAPI

    Dim PT      As POINTAPI
    Dim lWidth  As Long
    Dim lHeight As Long

    '   Get Our Position
    Call GetCursorPos(PT)
    '   Convert coordinates
    Call ScreenToClient(m_Hwnd, PT)

    '   Correct for Offeset of the Borders
    If m_Appearance = [3D] Then
        PT.X = PT.X - 2
        PT.Y = PT.Y - 2
    Else
        PT.X = PT.X - 1
        PT.Y = PT.Y - 1
    End If

    '   Get the size of the TextBox
    lWidth = UserControl.ScaleX(txtResult.Width, vbTwips, vbPixels)
    lHeight = UserControl.ScaleY(txtResult.Height, vbTwips, vbPixels)

    '   Sanity Check...are these real numbers (i.e. outside out control)?
    If PT.X < 0 Then PT.X = 0
    If PT.X > lWidth Then PT.X = lWidth
    If PT.Y < 0 Then PT.Y = 0
    If PT.Y > lHeight Then PT.Y = lHeight
    '   Now convert from Pixels to Twips
    PT.X = UserControl.ScaleX(PT.X, vbPixels, vbTwips)
    PT.Y = UserControl.ScaleY(PT.Y, vbPixels, vbTwips)
    '   Pass back the Corrected Coordinates
    GetCursorPosition = PT
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetThemeInfo
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function GetThemeInfo() As String

    Dim lPtrColorName As Long
    Dim lPtrThemeFile As Long
    Dim hTheme        As Long
    Dim sColorName    As String
    Dim sThemeFile    As String

    If m_bIsWinXpOrLater Then
        hTheme = OpenThemeData(hWnd, StrPtr("Button"))

        If hTheme Then

            ReDim bThemeFile(0 To 260 * 2) As Byte

            lPtrThemeFile = VarPtr(bThemeFile(0))

            ReDim bColorName(0 To 260 * 2) As Byte

            lPtrColorName = VarPtr(bColorName(0))

            If GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0) <> &H0 Then
                GetThemeInfo = "UxTheme_Error"

                Exit Function

            Else
                sThemeFile = TrimNull(bThemeFile)
                sColorName = TrimNull(bColorName)
            End If

            CloseThemeData hTheme
        End If
    End If

    If LenB(Trim$(sColorName)) = 0 Then sColorName = "None"
    GetThemeInfo = sColorName
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function LongToHexColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lNewColor (Long)
'!--------------------------------------------------------------------------------
Public Function LongToHexColor(ByVal lNewColor As Long) As String
    '   Translate the Color to RGB with Current Palette and pass
    '   back the Hex String Equiv...
    LongToHexColor = pHexColorStr(TranslateColor(lNewColor))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub OpenComboBox
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   CBox (ComboBox)
'                              ShowIt (Boolean = True)
'!--------------------------------------------------------------------------------
Private Sub OpenComboBox(CBox As ComboBox, Optional ShowIt As Boolean = True)
    '   A thin wrapper to open the a ComboBox via API
    SendMessage CBox.hWnd, CB_SHOWDROPDOWN, ShowIt, ByVal 0&
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PaintControl
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   AutoTheme (String)
'                              Index (Long)
'!--------------------------------------------------------------------------------
Private Sub PaintControl(ByVal AutoTheme As String, ByVal Index As Long)

    With UserControl
        LockWindowUpdate .hWnd
        ShapeBorder.Visible = True

        Select Case m_Theme

            Case [pbAuto]

                Select Case AutoTheme

                    Case "None"
                        GoTo Classic

                    Case "NormalColor"
                        GoTo NormalColor

                    Case "HomeStead"
                        GoTo HomeStead

                    Case "Metallic"
                        GoTo Metallic

                    Case Else
                        GoTo NormalColor
                End Select

            Case [pbClassic]
Classic:
                BackColor = m_PrevBackColor
                .ShapeBorder.Visible = False
                .pbDrop.Visible = False
                .pbPick.Visible = False
                '   Set the new visual styles to the passed type (3D or Flat)
                Call ButtonAppearance(cmdPick, m_Appearance)
                Call ButtonAppearance(cmdDrop, m_Appearance)
                .txtResult.Appearance = m_Appearance
                .txtResult.BorderStyle = 1
                .cmdPick.Visible = True
                .cmdDrop.Visible = False

            Case [pbBlue]
NormalColor:
                BackColor = &HFFFFFF
                .ShapeBorder.Visible = True
                .pbPick.Visible = True
                .txtResult.Appearance = 0
                .txtResult.BorderStyle = 0
                .cmdPick.Visible = False
                .cmdDrop.Visible = False
                .ShapeBorder.BorderColor = &HB99D7F

                Select Case m_State

                    Case [pbNormal]

                        If Index = 0 Then
                            If .pbPick.Picture <> .imBlue(0).Picture Then
                                Set .pbPick.Picture = .imBlue(0).Picture
                            End If

                        Else

                            If .pbDrop.Picture <> .imBlueDrop(0).Picture Then
                                Set pbDrop.Picture = .imBlueDrop(0).Picture
                            End If
                        End If

                    Case [pbHover]

                        If Index = 0 Then
                            If .pbPick.Picture <> .imBlue(1).Picture Then
                                Set .pbPick.Picture = .imBlue(1).Picture
                            End If

                        Else

                            If pbDrop.Picture <> .imBlueDrop(1).Picture Then
                                Set pbDrop.Picture = .imBlueDrop(1).Picture
                            End If
                        End If

                    Case [pbDown]

                        If Index = 0 Then
                            If .pbPick.Picture <> .imBlue(2).Picture Then
                                Set .pbPick.Picture = .imBlue(2).Picture
                            End If

                        Else

                            If pbDrop.Picture <> .imBlueDrop(2).Picture Then
                                Set pbDrop.Picture = .imBlueDrop(2).Picture
                            End If
                        End If

                    Case [pbDisabled]

                        If Index = 0 Then
                            If .pbPick.Picture <> .imBlue(3).Picture Then
                                Set .pbPick.Picture = .imBlue(3).Picture
                            End If

                        Else

                            If pbDrop.Picture <> .imBlueDrop(3).Picture Then
                                Set pbDrop.Picture = .imBlueDrop(3).Picture
                            End If
                        End If

                        ShapeBorder.BorderColor = &HC0C0C0
                End Select

            Case [pbHomeStead]
HomeStead:
                BackColor = &HFFFFFF
                .ShapeBorder.Visible = True
                .pbPick.Visible = True
                .txtResult.Appearance = 0
                .txtResult.BorderStyle = 0
                .cmdPick.Visible = False
                .cmdDrop.Visible = False
                .ShapeBorder.BorderColor = &H69A18B

                Select Case m_State

                    Case [pbNormal]

                        If Index = 0 Then
                            If pbPick.Picture <> .imHomeStead(0).Picture Then
                                Set pbPick.Picture = .imHomeStead(0).Picture
                            End If

                        Else

                            If .pbDrop.Picture <> .imHomeSteadDrop(0).Picture Then
                                Set .pbDrop.Picture = .imHomeSteadDrop(0).Picture
                            End If
                        End If

                    Case [pbHover]

                        If Index = 0 Then
                            If pbPick.Picture <> .imHomeStead(1).Picture Then
                                Set pbPick.Picture = .imHomeStead(1).Picture
                            End If

                        Else

                            If .pbDrop.Picture <> .imHomeSteadDrop(1).Picture Then
                                Set .pbDrop.Picture = .imHomeSteadDrop(1).Picture
                            End If
                        End If

                    Case [pbDown]

                        If Index = 0 Then
                            If pbPick.Picture <> .imHomeStead(2).Picture Then
                                Set pbPick.Picture = .imHomeStead(2).Picture
                            End If

                        Else

                            If .pbDrop.Picture <> .imHomeSteadDrop(2).Picture Then
                                Set .pbDrop.Picture = .imHomeSteadDrop(2).Picture
                            End If
                        End If

                    Case [pbDisabled]

                        If Index = 0 Then
                            If pbPick.Picture <> .imHomeStead(3).Picture Then
                                Set pbPick.Picture = .imHomeStead(3).Picture
                            End If

                        Else

                            If .pbDrop.Picture <> .imHomeSteadDrop(3).Picture Then
                                Set .pbDrop.Picture = .imHomeSteadDrop(3).Picture
                            End If
                        End If

                        ShapeBorder.BorderColor = &HC0C0C0
                End Select

            Case [pbMetallic]
Metallic:
                BackColor = &HFFFFFF
                .ShapeBorder.Visible = True
                .pbPick.Visible = True
                .txtResult.Appearance = 0
                .txtResult.BorderStyle = 0
                .cmdPick.Visible = False
                .cmdDrop.Visible = False
                .ShapeBorder.BorderColor = &HB99D7F

                Select Case m_State

                    Case [pbNormal]

                        If Index = 0 Then
                            If pbPick.Picture <> .imMetallic(0).Picture Then
                                Set pbPick.Picture = .imMetallic(0).Picture
                            End If

                        Else

                            If .pbDrop.Picture <> .imMetallicDrop(0).Picture Then
                                Set .pbDrop.Picture = .imMetallicDrop(0).Picture
                            End If
                        End If

                    Case [pbHover]

                        If Index = 0 Then
                            If pbPick.Picture <> .imMetallic(1).Picture Then
                                Set pbPick.Picture = .imMetallic(1).Picture
                            End If

                        Else

                            If .pbDrop.Picture <> .imMetallicDrop(1).Picture Then
                                Set .pbDrop.Picture = .imMetallicDrop(1).Picture
                            End If
                        End If

                    Case [pbDown]

                        If Index = 0 Then
                            If pbPick.Picture <> .imMetallic(2).Picture Then
                                Set .pbPick.Picture = .imMetallic(2).Picture
                            End If

                        Else

                            If .pbDrop.Picture <> .imMetallicDrop(2).Picture Then
                                Set .pbDrop.Picture = .imMetallicDrop(2).Picture
                            End If
                        End If

                    Case [pbDisabled]

                        If Index = 0 Then
                            If .pbPick.Picture <> .imMetallic(3).Picture Then
                                Set pbPick.Picture = .imMetallic(3).Picture
                            End If

                        Else

                            If .pbDrop.Picture <> .imMetallicDrop(3).Picture Then
                                Set .pbDrop.Picture = .imMetallicDrop(3).Picture
                            End If
                        End If

                        .ShapeBorder.BorderColor = &HC0C0C0
                End Select
        End Select

        .pbPick.Refresh
        .pbDrop.Refresh
        
        Locked = m_Locked
        LockWindowUpdate 0&
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function pHexColorStr
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lColor (Long)
'!--------------------------------------------------------------------------------
Private Function pHexColorStr(ByVal lColor As Long) As String
    '   Get the Hex version of the color...
    pHexColorStr = UCase$("&H" & Hex$(lColor))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ProcessFilter
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sFilter (String)
'!--------------------------------------------------------------------------------
Private Function ProcessFilter(sFilter As String) As String

    Dim I As Long

    '   This routine replaces the Pipe (|) character for filter
    '   strings and pads the size to the required legnth.
    '
    '   Example:
    '   - Input (String)
    '       "Supported files|*.bmp;*.doc;*.jpg;*.rtf;*.txt;*.tif|Bitmap files (*.bmp)|*.bmp|Word files (*.doc)|*.doc|JPEG files (*.jpg)|*.jpg|RichText files (*.rtf)|*.rtf|Text files (*.txt)|*.txt"
    '   - Output (String)
    '       "Supported files *.bmp;*.doc;*.jpg;*.rtf;*.txt;*.tif Bitmap files (*.bmp) *.bmp Word files (*.doc) *.doc JPEG files (*.jpg) *.jpg RichText files (*.rtf) *.rtf Text files (*.txt) *.txt"
    '
    '   Check to see if the Filter is set....if not then use the "All Files (*.*)"
    If LenB(sFilter) = 0 Then
        sFilter = "Supported Files|*.*|All Files (*.*)"
        '   Make sure to store this in the Control as well...
        m_Filters = sFilter
    End If

    '   Now Replace the Pipes in the Filter String
    For I = 1 To Len(sFilter)

        If (Mid$(sFilter, I, 1) = "|") Then
            Mid$(sFilter, I, 1) = vbNullChar
        End If

    Next

    '   Pad the string to the correct length
    If (Len(sFilter) < MAX_PATH) Then
        sFilter = sFilter & String$(MAX_PATH - Len(sFilter), 0)
    Else
        sFilter = sFilter & str2vbNullChar
    End If

    '   Pass the fixed filter back....
    ProcessFilter = sFilter
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub pSelectText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   TxtBox (TextBox)
'!--------------------------------------------------------------------------------
Private Sub pSelectText(ByVal TxtBox As TextBox)

    With TxtBox
        '   Select the text
        .SelStart = 0
        .SelLength = Len(TxtBox.Text)
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function QualifyPath
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sPath (String)
'!--------------------------------------------------------------------------------
Private Function QualifyPath(ByVal sPath As String) As String

    Dim lStrCnt  As Long
    Dim lStr2Cnt As Long

    If m_QualifyPaths Then
        If Not FileExists(sPath) Then
            '   Look for the PathSep
            lStrCnt = InStrRev(sPath, "\")
            lStr2Cnt = InStrRev(sPath, ":")

            If ((lStrCnt <> Len(sPath)) Or Right$(sPath, 1) <> "\") And lStrCnt > 1 And lStr2Cnt > 2 Then
                '   None, so add it...
                QualifyPath = BackslashAdd2Path(sPath)
            Else
                '   We are good, so return the value unchanged
                QualifyPath = sPath
            End If

        Else
            QualifyPath = sPath
        End If

    Else
        QualifyPath = sPath
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function BackslashAdd2Path
'! Description (Описание)  :   [Добавление слэша на конце]
'! Parameters  (Переменные):   strPath (String)
'!--------------------------------------------------------------------------------
Private Function BackslashAdd2Path(ByVal strPath As String) As String
    strPath = strPath & str2vbNullChar
    PathAddBackslash strPath
    BackslashAdd2Path = TrimNull(strPath)
End Function


'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Refresh
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'!--------------------------------------------------------------------------------
Public Sub Refresh(Optional ByVal Index As Long)

    Dim AutoTheme As String

    With UserControl
        AutoTheme = GetThemeInfo
        Call PaintControl(AutoTheme, Index)

        Select Case m_DialogType

            Case [ucFolder]
                '   Update the Folder PickBox Values
                Locked = m_Locked

                If m_UseDialogText Then
                    If LenB(Path) = 0 Then
                        .txtResult.Text = m_DialogMsg([ucFolder])
                    End If
                Else
                    .txtResult.Text = Path
                End If

                .cmdPick.ToolTipText = m_ToolTipText([ucFolder])

            Case [ucOpen]

                Locked = m_Locked
                '   Update the Open PickBox Values
                If m_UseDialogText Then
                    If (LenB(m_Path) = 0) Or (Left$(m_Path, 3) <> Left$(.txtResult.Text, 3)) Then
                        If LenB(Path) = 0 Then
                            .txtResult.Text = m_DialogMsg([ucOpen])
                        End If
                    End If

                Else
                    .txtResult.Text = Path
                End If

                .cmdPick.ToolTipText = m_ToolTipText([ucOpen])

            Case [ucSave]

                Locked = m_Locked
                '   Update the Save PickBox Values
                If m_UseDialogText Then
                    If LenB(Path) = 0 Then
                        .txtResult.Text = m_DialogMsg([ucSave])
                    End If
                Else
                    .txtResult.Text = Path
                End If

                .cmdPick.ToolTipText = m_ToolTipText([ucSave])
        End Select

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Reset
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Reset()

    '   Reset everthing to defaults....
    On Error Resume Next

    Appearance = 1
    '[3D]
    BackColor = &HFFFFFF
    m_DialogMsg([ucFolder]) = Def_DialogMsgFolder
    m_DialogMsg([ucOpen]) = Def_DialogMsgFile
    m_DialogMsg([ucSave]) = Def_DialogMsgFile
    m_Filters = "Supported files|*.*|All Files (*.*)"
    m_FileFlags = IIf(m_DialogType = ucOpen, ShowOpen_Default, ShowSave_Default)

    If Not m_Font Is Nothing Then
        m_Font = Nothing
    End If

    ForeColor = &H0

    ReDim m_Filename(1 To 1)

    m_Filename(1) = vbNullString
    m_Path = vbNullString
    m_ToolTipText([ucFolder]) = Def_ToolTipMsgFolder

    If m_MultiSelect Then
        m_ToolTipText([ucOpen]) = Def_ToolTipMsgFile
    Else
        m_ToolTipText([ucOpen]) = Def_ToolTipMsgFile
    End If

    m_ToolTipText([ucSave]) = Def_ToolTipMsgFile
    'm_UseDialogColor = False
    'm_UseDialogText = True
    'm_Locked = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ShowOpen
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sFilter (String)
'                              sInitPath (String)
'!--------------------------------------------------------------------------------
Private Function ShowOpen(sFilter As String, sInitPath As String) As SelectedFile

    Dim lRet                As Long
    Dim count               As Integer
    Dim LastCharacter       As Integer
    Dim NewCharacter        As Integer
    Dim tempFiles()         As String
    
    ReDim tempFiles(1 To 200)

    '   Open Common Dialog Controls
    '   Note: This has been modified to allow the user to select either
    '         a Single or Mutliple Files...In either case the data is sent
    '         back to the caller as part of the SelectedFile data structure
    '         which has been modified to allow for Array of strings in the
    '         sFiles section.
    With FileDialog
        .nStructSize = Len(FileDialog)
        .hWndOwner = UserControl.Parent.hWnd
        .sFileTitle = String$(2048, vbNullChar)
        .nTitleSize = Len(FileDialog.sFileTitle)
        .sFile = FileDialog.sFile & String$(2048, vbNullChar)
        .nFileSize = Len(FileDialog.sFile)

        If LenB(sInitPath) Then
            .sInitDir = sInitPath
        Else
            .sInitDir = GetAppPath
        End If

        If m_FileFlags <> 0 Then
            .flags = m_FileFlags
        Else
            .flags = ShowOpen_Default
        End If

        If m_MultiSelect Then
            .flags = .flags Or AllowMultiselect
        End If
        
        .sDlgTitle = m_DialogMsg(ucOpen) & String$(2048, vbNullChar)

        '   Init the File Names
        .sFile = vbNullString & String$(2048, vbNullChar)
        '   Process the Filter string to replace the
        '   pipes and fix the len to correct dims
        sFilter = ProcessFilter(sFilter)
        '   Set the Filter for Use...
        .sFilter = sFilter
        '   Set the Default Extension
        .sDefFileExt = m_DefaultExt
    End With

    '   Open the Common Dialog via API Calls
    lRet = GetOpenFileName(VarPtr(FileDialog))

    If lRet Then
        '   Retry Flag
GoAgain:

        If (FileDialog.nFileOffset = 0) Then

            '   This is a first time through, so the Offset will be zero. This is the
            '   case when MultiSelect = False and this is our first file selected.
            '   For cases where this is not our first time, then see "Else" notes below.
            '
            '   Extract the single Filename and pass it back....
            ReDim ShowOpen.sFiles(1 To 1)

            ShowOpen.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
            ShowOpen.nFilesSelected = 1
            ShowOpen.sFiles(1) = Mid$(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(FileDialog.sFile, vbNullChar) - FileDialog.nFileOffset - 1)
        ElseIf (InStr(FileDialog.nFileOffset, FileDialog.sFile, vbNullChar) = FileDialog.nFileOffset) Then
            '   See if we have an offset by the dialog and see if this matches the position of
            '   the (vbNullChar) character. If this is the case, then we have Mulplitple files selected
            '   in the FileDialog.sFile array. The GetOpenFileName passes back (vbNullChar) delimited filenames
            '   when we are in Multipile File selection mode, and the stripping of the names needs to be handled
            '   differently than when there is simply one....
            '
            '   Extract all of the files selected and pass them back in an array.
            LastCharacter = 0
            count = 0

            While ShowOpen.nFilesSelected = 0

                NewCharacter = InStr(LastCharacter + 1, FileDialog.sFile, vbNullChar)

                If count Then
                    tempFiles(count) = Mid$(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                Else
                    ShowOpen.sLastDirectory = Mid$(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                End If

                count = count + 1

                If InStr(NewCharacter + 1, FileDialog.sFile, vbNullChar) = InStr(NewCharacter + 1, FileDialog.sFile, str2vbNullChar) Then
                    tempFiles(count) = Mid$(FileDialog.sFile, NewCharacter + 1, InStr(NewCharacter + 1, FileDialog.sFile, str2vbNullChar) - NewCharacter - 1)
                    ShowOpen.nFilesSelected = count
                End If

                LastCharacter = NewCharacter

            Wend

            ReDim ShowOpen.sFiles(1 To ShowOpen.nFilesSelected)

            For count = 1 To ShowOpen.nFilesSelected

                If (Right$(tempFiles(count), 4) <> m_DefaultExt) Then
                    If (Len(m_DefaultExt) > 1) Then
                        tempFiles(count) = tempFiles(count) & m_DefaultExt
                    End If
                End If

                ShowOpen.sFiles(count) = tempFiles(count)
            Next

        Else
            '   This is the case where we have MutliSelect = False, but this is our
            '   Second through "n" times through...To fix this case we simlply set the
            '   FileOffset like it is our first time and then re-run the routine....
            '   The net effect is that the sub acts as if this were the first time and
            '   yeilds the name and path correctly.
            FileDialog.nFileOffset = 0
            GoTo GoAgain
        End If

        ShowOpen.bCanceled = False

        Exit Function

    Else
        '   The Cancel Button was pressed
        ShowOpen.sLastDirectory = vbNullString
        ShowOpen.nFilesSelected = 0
        ShowOpen.bCanceled = True
        Erase ShowOpen.sFiles

        Exit Function

    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ShowSave
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sFilter (String)
'!--------------------------------------------------------------------------------
Private Function ShowSave(ByVal sFilter As String) As SelectedFile

    Dim lRet      As Long
    Dim sFileName As String

    '   Save Common Dialog Controls
    With FileDialog
        .nStructSize = Len(FileDialog)
        .hWndOwner = UserControl.Parent.hWnd
        .sFileTitle = String$(2048, vbNullChar)
        .nTitleSize = Len(FileDialog.sFileTitle)
        .sFile = String$(2048, vbNullChar)
        .nFileSize = Len(FileDialog.sFile)

        If m_FileFlags <> 0 Then
            .flags = m_FileFlags
        Else
            .flags = ShowSave_Default
        End If

        '   Process the Filter string to replace the
        '   pipes and fix the len to correct dims
        sFilter = ProcessFilter(sFilter)
        '   Set the Filter for Use...
        .sFilter = sFilter
        '   Set the Default Extension
        .sDefFileExt = Mid$(m_DefaultExt, 2)
    End With

    lRet = GetSaveFileName(VarPtr(FileDialog))

    ReDim ShowSave.sFiles(1)

    If lRet Then
        '   This is a work around to a bug in the FileDialog.nFileOffset routine
        '   We will trim the path and filenames outside of this routine
        '   to yeild a more consistent result....
        FileDialog.nFileOffset = 0
        ShowSave.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
        ShowSave.nFilesSelected = 1
        sFileName = Mid$(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(FileDialog.sFile, vbNullChar) - FileDialog.nFileOffset - 1)

        If Right$(sFileName, 4) <> m_DefaultExt Then
            sFileName = sFileName & m_DefaultExt
        End If

        ShowSave.sFiles(1) = sFileName
        ShowSave.bCanceled = False

        Exit Function

    Else
        ShowSave.sLastDirectory = vbNullString
        ShowSave.nFilesSelected = 0
        ShowSave.bCanceled = True
        Erase ShowSave.sFiles

        Exit Function

    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub TrackMouseLeave
'! Description (Описание)  :   [Track the mouse leaving the indicated window]
'! Parameters  (Переменные):   lng_hWnd (Long)
'!--------------------------------------------------------------------------------
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

    Dim TME As TRACKMOUSEEVENT_STRUCT

    If bTrack Then

        With TME
            .cbSize = LenB(TME)
            .dwFlags = TME_LEAVE
            .hWndTrack = lng_hWnd
            .dwHoverTime = 1
        End With

        If bTrackUser32 Then
            TrackMouseEvent TME
        Else
            TrackMouseEventComCtl TME
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function TranslateColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lColor (Long)
'!--------------------------------------------------------------------------------
Public Function TranslateColor(ByVal lColor As Long) As Long

    On Error GoTo Func_ErrHandler

    '   System Color code to long RGB
    If OleTranslateColor(lColor, 0, TranslateColor) Then
        TranslateColor = -1
    End If

    Exit Function

Func_ErrHandler:
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function TrimPathByLen
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sInput (String)
'                              iTextWidth (Integer)
'                              sReplaceString (String = "...")
'                              sFont (String = "Tahoma")
'                              iFontSize (Integer = 8)
'!--------------------------------------------------------------------------------
Public Function TrimPathByLen(ByVal sInput As String, ByVal iTextWidth As Integer, Optional ByVal sReplaceString As String = "...", Optional ByVal sFont As String = "Tahoma", Optional ByVal iFontSize As Integer = 8) As String

    '**************************************************************************
    'Function TrimPathByLen
    '
    'Inputs:
    'sInput As String :         the path to alter
    'iTextWidth as Integer :    the desired length of the inputted path in twips
    'sReplaceString as String : the string which is interted for missing text.  Default "..."
    'sFont as String :          the font being used for display.  Default "Tahoma"
    'iFontSize as Integer :     the font size being used for display.  Default "8"
    'Output:
    'TrimPathByLen intellengently cuts the input (sInput) to a string that fits
    'within the desired Width.
    '
    '**************************************************************************
    Dim iInputLen As Integer, sBeginning As String, sEnd As String
    Dim aBuffer() As String, bAddedTrailSlash As Boolean
    Dim iIndex    As Integer, iArrayCount      As Integer
    Dim OldFont   As String, OldFontSize As Integer, OldScaleMode As ScaleModeConstants

    OldFont = UserControl.Font
    OldFontSize = UserControl.FontSize
    OldScaleMode = UserControl.ScaleMode
    'setup font attributes
    UserControl.Font = sFont
    UserControl.FontSize = iFontSize
    UserControl.ScaleMode = vbTwips
    'get length of input string in twips
    iInputLen% = UserControl.TextWidth(sInput$)

    'let's be reasonable here on the TextWidth
    If iTextWidth% < 200 Then

        Exit Function

    End If

    iTextWidth% = iTextWidth% - 400

    'make sure the desired text Width is smaller than
    'the length of the current string
    If iTextWidth < iInputLen% Then

        'now that we know how much to trim, we need to
        'determine the path type: local, network, or URL
        If InStr(sInput$, "\") Then

            'LOCAL
            'add trailing slash if there is none
            If Right$(sInput$, 1) <> "\" Then
                bAddedTrailSlash = True
                sInput$ = sInput$ & "\"
            End If

            'throw path into an array
            aBuffer() = Split(sInput$, "\")

            If UBound(aBuffer()) > LBound(aBuffer()) Then
                iArrayCount% = UBound(aBuffer()) - 1
                'the last element is blank
                sBeginning$ = aBuffer(0) & "\" & aBuffer(1) & "\"
                sEnd$ = "\" & aBuffer(iArrayCount%)

                If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                    'if the total outputed string is too big then stop
                    sBeginning$ = aBuffer(0) & "\"

                    If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                        TrimPathByLen$ = sReplaceString$ & sEnd$
                    Else
                        TrimPathByLen$ = sBeginning$ & sReplaceString$ & sEnd$
                    End If

                Else

                    For iIndex% = iArrayCount% - 1 To 1 Step -1
                        'go throug the remaing elements to get the best fit
                        sEnd$ = "\" & aBuffer(iIndex%) & sEnd$

                        If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                            'if the total outputed string is too big then stop
                            TrimPathByLen$ = sBeginning$ & sReplaceString$ & Mid$(sEnd$, Len(aBuffer(iIndex%)) + 2)

                            Exit For

                        End If

                        DoEvents
                    Next

                End If

            Else
                'there is only one array element: bad.
                TrimPathByLen$ = sInput$
            End If

            Exit Function

        ElseIf InStr(sInput$, "\") Then

            If InStr(sInput$, ":") Then

                'URL
                'start by triming off the extra params
                If InStr(sInput$, "?") Then sInput$ = Left$(sInput$, InStr(sInput$, "?") - 1)

                'add trailing slash if there is none
                If Right$(sInput$, 1) <> "\" Then
                    bAddedTrailSlash = True
                    sInput$ = sInput$ & "\"
                End If

                'throw path into an array
                aBuffer() = Split(sInput$, "\")

                If UBound(aBuffer()) > LBound(aBuffer()) Then
                    iArrayCount% = UBound(aBuffer()) - 1
                    'the last element is blank
                    sBeginning$ = aBuffer(0) & "\" & aBuffer(1) & "\"
                    sEnd$ = "\" & aBuffer(iArrayCount%)

                    If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                        'if the total outputed string is too big then stop
                        sBeginning$ = aBuffer(0) & "\"

                        If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                            TrimPathByLen$ = sReplaceString$ & sEnd$
                        Else
                            TrimPathByLen$ = sBeginning$ & sReplaceString$ & sEnd$
                        End If

                    Else

                        For iIndex% = iArrayCount% - 1 To 1 Step -1
                            'go throug the remaing elements to get the best fit
                            sEnd$ = "\" & aBuffer(iIndex%) & sEnd$

                            If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                                'if the total outputed string is too big then stop
                                TrimPathByLen$ = sBeginning$ & sReplaceString$ & Mid$(sEnd$, Len(aBuffer(iIndex%)) + 2)

                                Exit For

                            End If

                            DoEvents
                        Next

                    End If

                Else
                    'there is only one array element: bad.
                    TrimPathByLen$ = sInput$
                End If

            Else

                ' NETWORK
                'add trailing slash if there is none
                If Right$(sInput$, 1) <> "\" Then
                    bAddedTrailSlash = True
                    sInput$ = sInput$ & "\"
                End If

                'throw path into an array
                aBuffer() = Split(sInput$, "\")

                If UBound(aBuffer()) > LBound(aBuffer()) Then
                    iArrayCount% = UBound(aBuffer()) - 1
                    'the last element is blank
                    sBeginning$ = aBuffer(0) & "\" & aBuffer(1) & "\"
                    sEnd$ = "\" & aBuffer(iArrayCount%)

                    If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                        'if the total outputed string is too big then stop
                        sBeginning$ = aBuffer(0) & "\"

                        If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                            TrimPathByLen$ = sReplaceString$ & sEnd$
                        Else
                            TrimPathByLen$ = sBeginning$ & sReplaceString$ & sEnd$
                        End If

                    Else

                        For iIndex% = iArrayCount% - 1 To 1 Step -1
                            'go throug the remaing elements to get the best fit
                            sEnd$ = "\" & aBuffer(iIndex%) & sEnd$

                            If (UserControl.TextWidth(sBeginning$) + UserControl.TextWidth(sReplaceString$) + UserControl.TextWidth(sEnd$)) > iTextWidth% Then
                                'if the total outputed string is too big then stop
                                TrimPathByLen$ = sBeginning$ & sReplaceString$ & Mid$(sEnd$, Len(aBuffer(iIndex%)) + 2)

                                Exit For

                            End If

                            DoEvents
                        Next

                    End If

                Else
                    'there is only one array element: bad.
                    TrimPathByLen$ = sInput$
                End If
            End If
        End If

    Else
        'we can return the value since it's already small enough
        TrimPathByLen$ = sInput$
    End If

    '   set them back
    UserControl.Font = OldFont
    UserControl.FontSize = OldFontSize
    UserControl.ScaleMode = OldScaleMode
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbMultiSel_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbMultiSel_Click()

    With UserControl
        '   Display the selected results from the ComboBox List
        .txtResult.Text = .cmbMultiSel.List(.cmbMultiSel.ListIndex)
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbMultiSel_KeyDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub cmbMultiSel_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUp Then

        '   See if we are at the top, if so then change
        '   the focus back to the textbox....as if it were
        '   part of the control
        If UserControl.cmbMultiSel.ListIndex = 0 Then
            UserControl.txtResult.SetFocus
        End If
    End If


End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDrop_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdDrop_Click()

    With UserControl

        If Not ComboBoxListVisible(.cmbMultiSel) Then
            '   It is closed, so open it via code....
            Call OpenComboBox(.cmbMultiSel, True)
        Else
            '   Set the focus to our TextBox
            .txtResult.SetFocus
        End If

        '   Drop List Clicked...
        RaiseEvent DropClick
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDrop_MouseDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub cmdDrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Get the Cursor Position
    m_Pnt = GetCursorPosition
    RaiseEvent MouseDown(Button, Shift, CSng(m_Pnt.X), CSng(m_Pnt.Y))
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDrop_MouseMove
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub cmdDrop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Get the Cursor Position
    m_Pnt = GetCursorPosition

    If (m_PrevLoc.X <> m_Pnt.X) Then
        If (m_PrevLoc.Y <> m_Pnt.Y) Then
            RaiseEvent MouseMove(Button, Shift, CSng(m_Pnt.X), CSng(m_Pnt.Y))
            m_PrevLoc = m_Pnt
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDrop_MouseUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub cmdDrop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '   Make sure the focus is on the TextBox and not the drop button
    UserControl.txtResult.SetFocus

    '   Get the Cursor Position
    m_Pnt = GetCursorPosition
    RaiseEvent MouseUp(Button, Shift, CSng(m_Pnt.X), CSng(m_Pnt.Y))
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdPick_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdPick_Click()

    Dim psFile    As SelectedFile
    Dim I         As Long
    Dim sExt      As String
    Dim sFolder   As String
    Dim AutoTheme As String

    On Error Resume Next

    With UserControl
        AutoTheme = GetThemeInfo
        '   Make sure the Combobox is hidden
        .cmdDrop.Visible = False
        .pbDrop.Visible = False

        '   Which dialog is active?
        Select Case m_DialogType

            Case [ucFolder]

                'ShowFolder_Default
                With New CommonDialog
                    .InitDir = PathCollect(txtResult.Text)
                    .flags = CdlBIFNewDialogStyle
                    .DialogTitle = m_DialogMsg(ucFolder)
                                        
                    If .ShowFolderBrowser = True Then
                        sFolder = .FileName
                    End If
                    

                End With

                If LenB(sFolder) Then
                    m_Path = QualifyPath(sFolder)
                    PropertyChanged "Path"

                    If m_UseDialogText Then

                        '   Trim the display name
                        If m_QualifyPaths Then
                            .txtResult.Text = TrimPathByLen(m_Path, .txtResult.Width - .cmdPick.Width - 40)
                        Else
                            .txtResult.Text = m_Path
                        End If
                    End If
                End If

            Case [ucOpen], [ucSave]

                '   Same basic routine, with different calls to start
                If m_DialogType = [ucOpen] Then
                    'psFile = ShowOpen(m_Filters, txtResult.Text)
                    psFile = ShowOpen(m_Filters, PathCollect(txtResult.Text))
                Else
                    psFile = ShowSave(m_Filters)
                End If

                If (psFile.bCanceled = False) Then
                    If (psFile.nFilesSelected) Then
                        If m_DialogType = [ucOpen] Then

                            '   Set the Command Button visable
                            If (m_Theme = pbClassic) Or (AutoTheme = "None") Then
                                .cmdDrop.Visible = m_MultiSelect
                            Else
                                .pbDrop.Visible = m_MultiSelect
                            End If

                            '   Concatinate the filename and path
                            If m_MultiSelect Then
                                '   Store the qaulified path
                                m_Path = QualifyPath(psFile.sLastDirectory)
                                PropertyChanged "Path"
                                '   Count the Files
                                FileCount = UBound(psFile.sFiles) - LBound(psFile.sFiles) + 1

                                If m_FileCount = 1 Then
                                    '   Erase the array...this is over kill
                                    '   but better to be safe than sorry ;-)
                                    Erase m_Filename

                                    '   Redim to a vector...
                                    ReDim m_Filename(1 To 1)

                                    '   Clear the ComboBox
                                    .cmbMultiSel.Clear
                                    '   Store the Filename
                                    m_Filename(1) = psFile.sFiles(1)
                                    PropertyChanged "Filename"
                                    '   Add the Trimmed Filename and Path
                                    .cmbMultiSel.AddItem TrimPathByLen(m_Path & psFile.sFiles(1), .txtResult.Width - 40)
                                Else
                                    '   Erase the array...this is over kill
                                    '   but better to be safe than sorry ;-)
                                    Erase m_Filename

                                    '   Redim to a vector...
                                    ReDim m_Filename(1 To m_FileCount)

                                    '   Clear the ComboBox
                                    .cmbMultiSel.Clear

                                    '   Store the Filenames
                                    For I = 1 To m_FileCount
                                        .cmbMultiSel.AddItem TrimPathByLen(QualifyPath(m_Path) & psFile.sFiles(I), .txtResult.Width - 40)
                                        m_Filename(I) = m_Path & psFile.sFiles(I)
                                    Next

                                End If

                            Else

                                ReDim m_Filename(1 To 1)

                                '   Store the qaulified path
                                m_Path = QualifyPath(ExtractPath(psFile.sFiles(1)))
                                PropertyChanged "Path"
                                m_Filename(1) = psFile.sFiles(1)
                                m_FileCount = 1
                            End If

                            PropertyChanged "Filename"

                            If m_UseDialogText Then

                                '   Trim the display name
                                If m_MultiSelect Then
                                    '   Adjust the name len to account for our new button
                                    .txtResult.Text = TrimPathByLen(m_Filename(1), .txtResult.Width - .cmdPick.Width - .cmdDrop.Width - 40)
                                Else

                                    If m_QualifyPaths Then
                                        .txtResult.Text = TrimPathByLen(m_Filename(1), .txtResult.Width - .cmdPick.Width - 40)
                                    Else
                                        .txtResult.Text = m_Filename(1)
                                    End If
                                End If
                            End If

                            '   Focus on the final name
                            .txtResult.SetFocus
                        Else

                            '   Concatinate the filename and path
                            ReDim m_Filename(1 To 1)

                            If Not (Right$(psFile.sFiles(1), 4) Like ".*") Then
Retry:
                                '   This section handles files which are returned without extnsions
                                sExt = InputBox("The File Extension is Missing!" & vbCrLf & "Please Enter a Valid Extension Below...", "ucPickBox", , (.Parent.ScaleWidth \ 2) + .Parent.Left - 2700, (.Parent.ScaleHeight \ 2) + .Parent.Top - 800)

                                If LenB(sExt) = 0 Then
                                    If MsgBox("     The File Extension is Invalid!" & vbCrLf & vbCrLf & "File will be saved with .txt extension.", vbExclamation + vbOKCancel, "ucPickBox") = vbOK Then
                                        '   Just use the default text file type
                                        sExt = ".txt"
                                    Else
                                        '   Give them another try to get this right...
                                        GoTo Retry
                                    End If
                                End If

                                '   Fix missing "." in the extension
                                If (InStr(sExt, ".") = 0) Or (Len(sExt) = 3) Then
                                    psFile.sFiles(1) = psFile.sFiles(1) & "." & sExt
                                Else
                                    psFile.sFiles(1) = psFile.sFiles(1) & sExt
                                End If
                            End If

                            '   Store the Filename
                            m_Filename(1) = psFile.sFiles(1)
                            PropertyChanged "Filename"
                            '   Store the qualified path
                            m_Path = QualifyPath(ExtractPath(m_Filename(1)))
                            PropertyChanged "Path"

                            If m_UseDialogText Then
                                '   Trim the display name
                                .txtResult.Text = TrimPathByLen(psFile.sFiles(1), .txtResult.Width - .cmdPick.Width - 40)
                            End If

                            FileCount = 1
                        End If

                        '   Focus on the final name
                        .txtResult.SetFocus
                    End If
                End If

                RaiseEvent PathChanged
        End Select

        RaiseEvent Click
        m_Pnt = GetCursorPosition()
        RaiseEvent MouseDown(vbLeftButton, 0, CSng(m_Pnt.X), CSng(m_Pnt.Y))
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdPick_MouseDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub cmdPick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Get the Cursor Position
    m_Pnt = GetCursorPosition
    RaiseEvent MouseDown(Button, Shift, CSng(m_Pnt.X), CSng(m_Pnt.Y))
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdPick_MouseMove
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub cmdPick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Get the Cursor Position
    m_Pnt = GetCursorPosition

    If (m_PrevLoc.X <> m_Pnt.X) Then
        If (m_PrevLoc.Y <> m_Pnt.Y) Then
            RaiseEvent MouseMove(Button, Shift, CSng(m_Pnt.X), CSng(m_Pnt.Y))
            m_PrevLoc = m_Pnt
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdPick_MouseUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub cmdPick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Get the Cursor Position
    m_Pnt = GetCursorPosition
    RaiseEvent MouseUp(Button, Shift, CSng(m_Pnt.X), CSng(m_Pnt.Y))
End Sub

Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
    On Error Resume Next
    Dim This As OLEGuids.IOleInPlaceActiveObjectVB
    
    Set This = UserControl.ActiveControl.Object
    This.TranslateAccelerator Handled, RetVal, wMsg, wParam, lParam, Shift
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub pbDrop_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub pbDrop_Click()
    Call cmdDrop_Click
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub pbDrop_MouseDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub pbDrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        m_State = pbDown
        Call Refresh(1)
    End If

    Call cmdDrop_MouseDown(Button, Shift, X, Y)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub pbDrop_MouseMove
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub pbDrop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call cmdDrop_MouseMove(Button, Shift, X, Y)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub pbDrop_MouseUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub pbDrop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call cmdDrop_MouseUp(Button, Shift, X, Y)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub pbPick_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub pbPick_Click()
    Call cmdPick_Click
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub pbPick_MouseDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub pbPick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        m_State = pbDown
        Call Refresh(0)
    End If

    Call cmdPick_MouseDown(Button, Shift, X, Y)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub pbPick_MouseMove
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub pbPick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call cmdPick_MouseMove(Button, Shift, X, Y)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub pbPick_MouseUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub pbPick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call cmdPick_MouseUp(Button, Shift, X, Y)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Show_FolderBrowse
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Show_FolderBrowse()
    DialogType = ucFolder
    cmdPick_Click
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Show_Open
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Show_Open()
    DialogType = ucOpen
    cmdPick_Click
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Show_Save
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Show_Save()
    DialogType = ucSave
    cmdPick_Click
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtResult_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtResult_GotFocus()

    With UserControl
        '   Select the text for changing...
        Call pSelectText(.txtResult)
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtResult_KeyDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub txtResult_KeyDown(KeyCode As Integer, Shift As Integer)

    With UserControl

        Select Case KeyCode

            Case vbKeyReturn
                '   Call the LostFocus Event Handler
                Call txtResult_LostFocus

            Case vbKeyDown

                '   This routine allow the user to arrow down to the combobox
                '   droplist. The uparrow function is in the combobox keydown
                '   event handler...
                If m_DialogType = ucOpen Then
                    If m_MultiSelect Then
                        If .cmbMultiSel.ListCount Then
                            '   Set the ListIndex to 0
                            .cmbMultiSel.ListIndex = 0
                            '   Now drop the box
                            Call OpenComboBox(.cmbMultiSel, True)
                            '   Now set the focus there
                            .cmbMultiSel.SetFocus
                        End If
                    End If
                End If

        End Select

    End With

    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtResult_KeyPress
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyAscii (Integer)
'!--------------------------------------------------------------------------------
Private Sub txtResult_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtResult_KeyUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub txtResult_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtResult_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtResult_LostFocus()

    Dim TmpName As String
    Dim I       As Long

    On Error Resume Next

    With UserControl

        Select Case m_DialogType

            Case [ucFolder]

                '   Nothing...this is locked
            Case [ucOpen], [ucSave]

                If (LenB(.txtResult.Text) = 0) Then

                    Exit Sub

                End If

                '   See if we have a compacted path...
                '   Note: This happens when we pick a file and
                '         compact the Path Name using the cmdPick Button.
                '         The TextBox gets focus on completion of the
                '         file selection, then when the TextBox looses focus
                '         for the next selection the path does not make sense
                '         due to ellipses (...), and therefore should be ignored.
                If InStr(.txtResult.Text, "...") Then
                    TmpName = m_Filename(1)
                Else
                    TmpName = .txtResult.Text
                End If

                '   Handle cases where the file name is not set (i.e. Cancel)
                If LenB(.txtResult.Text) = 0 Then

                    Exit Sub

                End If

                If .txtResult.Text = m_DialogMsg(m_DialogType) Then

                    Exit Sub

                End If

                '   We have a valid name, so process it...
                If FileExists(TmpName) Then
                    '   Store this for later..
                    m_Filename(1) = TmpName
                Else

                    If .txtResult.Text <> m_DialogMsg(m_DialogType) Then
                        '   Pass the value to the textbox
                        'MsgBox "The Name Entered is Invalid!", vbExclamation + vbOKOnly, "ucPickBox"
                    End If
                End If

        End Select

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtResult_MouseDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub txtResult_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Get the Cursor Position
    m_Pnt = GetCursorPosition
    RaiseEvent MouseDown(Button, Shift, CSng(m_Pnt.X), CSng(m_Pnt.Y))
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtResult_MouseMove
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub txtResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Get the Cursor Position
    m_Pnt = GetCursorPosition

    If (m_PrevLoc.X <> m_Pnt.X) Then
        If (m_PrevLoc.Y <> m_Pnt.Y) Then
            RaiseEvent MouseMove(Button, Shift, CSng(m_Pnt.X), CSng(m_Pnt.Y))
            m_PrevLoc = m_Pnt
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtResult_MouseUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub txtResult_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Get the Cursor Position
    m_Pnt = GetCursorPosition
    RaiseEvent MouseUp(Button, Shift, CSng(m_Pnt.X), CSng(m_Pnt.Y))
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_GotFocus()

    With UserControl.txtResult
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Initialize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Initialize()
    m_bIsWinXpOrLater = IsWinXPOrLater
    '   Get Our Handle
    m_Hwnd = UserControl.hWnd
    '   Rest the Control to its defaults...
    Call Reset
    
    Set m_cSubclass = New cSelfSubHookCallback
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_InitProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()
    m_Appearance = [3D]
    m_BackColor = IIf(m_BackColor = &H0, &HFFFFFF, m_BackColor)
    m_Filters = "Supported files|*.*|All Files (*.*)"
    m_FileFlags = IIf(m_DialogType = ucOpen, ShowOpen_Default, ShowSave_Default)
    m_Forecolor = &H0
    m_Theme = pbAuto
    m_UseAutoForeColor = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_ReadProperties
'! Description (Описание)  :   [Read the properties from the property bag - also, a good place to start the subclassing (if we're running)]
'! Parameters  (Переменные):   PropBag (PropertyBag)
'!--------------------------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        m_Appearance = .ReadProperty("Appearance", [3D])
        m_UseAutoForeColor = .ReadProperty("UseAutoForeColor", True)
        m_BackColor = .ReadProperty("BackColor", &HFFFFFF)
        m_DefaultExt = .ReadProperty("DefaultExt", ".txt")
        m_DialogMsg([ucFolder]) = .ReadProperty("DialogMsg0", Def_DialogMsgFolder)
        m_DialogMsg([ucOpen]) = .ReadProperty("DialogMsg1", Def_DialogMsgFile)
        m_DialogMsg([ucSave]) = .ReadProperty("DialogMsg2", Def_DialogMsgFile)
        m_DialogType = .ReadProperty("DialogType", [ucFolder])
        m_Enabled = .ReadProperty("Enabled", True)
        m_FileFlags = .ReadProperty("FileFlags", IIf(m_DialogType = ucOpen, ShowOpen_Default, ShowSave_Default))
        m_Filters = .ReadProperty("Filters", vbNullString)
        Set m_Font = .ReadProperty("Font", Nothing)
        m_Forecolor = .ReadProperty("ForeColor", &H0)
        m_MultiSelect = .ReadProperty("MultiSelect", False)
        m_Path = .ReadProperty("Path", vbNullString)
        m_Theme = .ReadProperty("Theme", [pbAuto])
        m_ToolTipText([ucFolder]) = .ReadProperty("ToolTipText0", Def_ToolTipMsgFolder)
        m_ToolTipText([ucOpen]) = .ReadProperty("ToolTipText1", Def_ToolTipMsgFile)
        m_ToolTipText([ucSave]) = .ReadProperty("ToolTipText2", Def_ToolTipMsgFile)
        m_UseDialogColor = .ReadProperty("UseDialogColor", False)
        m_UseDialogText = .ReadProperty("UseDialogText", True)
        m_Locked = .ReadProperty("Locked", False)
    End With

    UserControl_Resize
    '   Set the focus on the caller
    Call SetFocusAPI(UserControl.Parent.hWnd)
    
    'If we're not in design mode
    On Error GoTo H

    'If we're not in design mode
    If RunMode Then
        
        bTrack = True
        bTrackUser32 = APIFunctionPresent("TrackMouseEvent", "user32.dll")

        If Not bTrackUser32 Then
            If Not APIFunctionPresent("_TrackMouseEvent", "comctl32") Then
                bTrack = False
            End If
        End If

        If bTrack Then
                
            'Add the messages that we're interested in
            With m_cSubclass
                '   Start Subclassing using our Handle
                If .ssc_Subclass(UserControl.hWnd, ByVal exUserControl, 1, Me) Then
                    .ssc_AddMsg UserControl.hWnd, MSG_AFTER, WM_MOUSEMOVE, WM_MOUSELEAVE, WM_THEMECHANGED, WM_SYSCOLORCHANGE
                End If
                
                '   Subclass the Ellipse (Pick) Picturebox
                If .ssc_Subclass(UserControl.pbPick.hWnd, ByVal exUserControl, 1, Me) Then
                    .ssc_AddMsg UserControl.pbPick.hWnd, MSG_AFTER, WM_MOUSEMOVE, WM_MOUSELEAVE
                End If
                
                '   Subclass the Dropdown (Drop) Picturebox
                If .ssc_Subclass(UserControl.pbDrop.hWnd, ByVal exUserControl, 1, Me) Then
                    .ssc_AddMsg UserControl.pbDrop.hWnd, MSG_AFTER, WM_MOUSEMOVE, WM_MOUSELEAVE
                End If
                
                '   Subclass the Dropdown (Drop) Picturebox
                If .ssc_Subclass(UserControl.txtResult.hWnd, ByVal exUserControl, 1, Me) Then
                    .ssc_AddMsg UserControl.txtResult.hWnd, MSG_AFTER, WM_MOUSEMOVE, WM_MOUSELEAVE
                End If
    
            End With
        End If

    End If

H:

    On Error GoTo 0

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Resize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Resize()

    Dim AutoTheme   As String
    Dim lTextHeight As Long

    On Error Resume Next

    With UserControl
        '   Get the TextHeight for the textbox
        lTextHeight = .TextHeight("gЋ")
        '   Lock the window
        LockWindowUpdate .hWnd

        If .Width <= 1455 Then .Width = 1455
        AutoTheme = GetThemeInfo

        With .txtResult

            If (m_Theme = pbClassic) Or (AutoTheme = "None") Then

                '   Set the Min Height in Twips
                If Height <= 315 Then Height = 315
                .Top = 0
                .Left = 0
                .Width = ScaleWidth
                .Height = ScaleHeight
                UserControl.BackColor = vbButtonFace
            Else

                '   Set the Min Height in Twips
                If Height <> imBlue(0).Height + 30 Then Height = imBlue(0).Height + 30
                .Height = lTextHeight
                .Top = Height \ 2 - .Height \ 2
                .Left = ShapeBorder.BorderWidth * 2 * Screen.TwipsPerPixelX
                .Width = Width - (ShapeBorder.BorderWidth * 3 * Screen.TwipsPerPixelX)
                UserControl.BackColor = vbWhite
            End If

        End With

        With .ShapeBorder
            .Left = 0
            .Top = 0
            .Width = ScaleWidth
            .Height = ScaleHeight
        End With

        With .cmdPick

            '   Adjust the position if this is 3D or Flat
            If m_Appearance = [3D] Then
                .Left = UserControl.Width - .Width - 30
                .Top = txtResult.Top + 30
                .Height = Height - 30
            Else
                .Left = Width - .Width
                .Top = txtResult.Top
                .Height = txtResult.Height
            End If

        End With

        With .pbPick
            .Left = Width - .Width - 15
            .Top = Height \ 2 - imBlue(0).Height \ 2
            .Height = imBlue(0).Height
        End With

        With .cmbMultiSel

            '   Adjust the position if this is 3D or Flat
            If m_Appearance = [3D] Then
                .Left = 0
                .Top = 0
                .Width = txtResult.Width
            Else
                .Left = 0
                .Top = 10
                .Width = txtResult.Width
            End If

        End With

        With .cmdDrop

            '   Adjust the position if this is 3D or Flat
            If m_Appearance = [3D] Then
                .Left = cmdPick.Left - .Width + 10
            Else
                .Left = cmdPick.Left - .Width + 20
            End If

            .Top = cmdPick.Top
            .Width = cmdPick.Width
            .Height = cmdPick.Height
        End With

        '   Adjust the Dropbutton Image
        With .pbDrop
            .Left = pbPick.Left - .Width + 20
            .Top = pbPick.Top
            .Width = pbPick.Width
            .Height = pbPick.Height
        End With
    End With

    Call Refresh(0)
    Call Refresh(1)
    LockWindowUpdate 0&
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Show
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Show()
    Call Refresh(0)
    Call Refresh(1)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Terminate
'! Description (Описание)  :   [The control is terminating - a good place to stop the subclasser]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Terminate()
    'Terminate all subclassing
    m_cSubclass.ssc_Terminate
    Set m_cSubclass = Nothing
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_WriteProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropBag (PropertyBag)
'!--------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("Appearance", m_Appearance, [3D])
        Call .WriteProperty("UseAutoForeColor", m_UseAutoForeColor, True)
        Call .WriteProperty("BackColor", m_BackColor, &HFFFFFF)
        Call .WriteProperty("DefaultExt", m_DefaultExt, ".txt")
        Call .WriteProperty("DialogMsg0", m_DialogMsg([ucFolder]), Def_DialogMsgFolder)
        Call .WriteProperty("DialogMsg1", m_DialogMsg([ucOpen]), Def_DialogMsgFile)
        Call .WriteProperty("DialogMsg2", m_DialogMsg([ucSave]), Def_DialogMsgFile)
        Call .WriteProperty("DialogType", m_DialogType, [ucFolder])
        Call .WriteProperty("Enabled", m_Enabled, True)
        Call .WriteProperty("FileFlags", m_FileFlags, IIf(m_DialogType = ucOpen, ShowOpen_Default, ShowSave_Default))
        Call .WriteProperty("Filters", m_Filters, vbNullString)
        Call .WriteProperty("Font", m_Font, Nothing)
        Call .WriteProperty("ForeColor", m_Forecolor, &H0)
        Call .WriteProperty("MultiSelect", m_MultiSelect, False)
        Call .WriteProperty("Path", m_Path, vbNullString)
        Call .WriteProperty("Theme", m_Theme, [pbAuto])
        Call .WriteProperty("ToolTipText0", m_ToolTipText([ucFolder]), Def_ToolTipMsgFolder)
        Call .WriteProperty("ToolTipText1", m_ToolTipText([ucOpen]), Def_ToolTipMsgFile)
        Call .WriteProperty("ToolTipText2", m_ToolTipText([ucSave]), Def_ToolTipMsgFile)
        Call .WriteProperty("UseDialogColor", m_UseDialogColor, False)
        Call .WriteProperty("UseDialogText", m_UseDialogText, True)
        Call .WriteProperty("Locked", m_Locked, False)
        Call .WriteProperty("QualifyPaths", m_QualifyPaths, False)
    End With

End Sub

'======================================================================================================
'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub zWndProc1
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   bBefore (Boolean)
'                              bHandled (Boolean)
'                              lReturn (Long)
'                              lng_hWnd (Long)
'                              uMsg (Long)
'                              wParam (Long)
'                              lParam (Long)
'                              lParamUser (Long)
'!--------------------------------------------------------------------------------
Private Sub z_WndProc1(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)

    '*************************************************************************************************
    '* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
    '*              you will know unless the callback for the uMsg value is specified as
    '*              MSG_BEFORE_AFTER (both before and after the original WndProc).
    '* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
    '*              message being passed to the original WndProc and (if set to do so) the after
    '*              original WndProc callback.
    '* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
    '*              and/or, in an after the original WndProc callback, act on the return value as set
    '*              by the original WndProc.
    '* lng_hWnd   - Window handle.
    '* uMsg       - Message value.
    '* wParam     - Message related data.
    '* lParam     - Message related data.
    '* lParamUser - User-defined callback parameter
    '*************************************************************************************************
    'If you really know what you're doing, it's possible to change the values of the
    'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
    'values get passed to the default handler.. and optionaly, the 'after' callback
    
    Select Case uMsg

       Case WM_MOUSEMOVE

            If (lng_hWnd = pbPick.hWnd) Then
                If m_State <> pbHover Then
                    m_State = pbHover
                    Call Refresh(0)
                End If

                If Not bInCtrl Then
                    bInCtrl = True
                    TrackMouseLeave lng_hWnd
                    RaiseEvent MouseEnter
                End If

            ElseIf (lng_hWnd = pbDrop.hWnd) Then

                If m_State <> pbHover Then
                    m_State = pbHover
                    Call Refresh(1)
                End If

                If Not bInCtrl Then
                    bInCtrl = True
                    TrackMouseLeave lng_hWnd
                    RaiseEvent MouseEnter
                End If

            Else

                If m_State <> pbNormal Then
                    m_State = pbNormal
                    Call Refresh(0)
                    Call Refresh(1)
                End If

                bInCtrl = False
            End If

        Case WM_MOUSELEAVE

            If (lng_hWnd = pbPick.hWnd) Then
                m_State = pbNormal
                Call Refresh(0)
                bInCtrl = False
                RaiseEvent MouseLeave
            ElseIf (lng_hWnd = pbDrop.hWnd) Then
                m_State = pbNormal
                Call Refresh(1)
                bInCtrl = False
                RaiseEvent MouseLeave
            Else

                If m_State <> pbNormal Then
                    m_State = pbNormal
                    Call Refresh(0)
                    Call Refresh(1)
                End If

                bInCtrl = False
                RaiseEvent MouseLeave
            End If

        Case WM_SYSCOLORCHANGE, WM_THEMECHANGED
            m_State = pbNormal
            Call Refresh(0)
            Call Refresh(1)

    End Select

End Sub

