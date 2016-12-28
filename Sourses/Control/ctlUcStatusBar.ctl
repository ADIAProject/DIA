VERSION 5.00
Begin VB.UserControl ctlUcStatusBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DrawWidth       =   56
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ctlUcStatusBar.ctx":0000
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   600
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "ctlUcStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+  File Description:
'       ucStatusBar - A Selfsubclassed Theme Aware ucStatusBar Control which Provides Dynamic Properties
'
'   Product Name:
'       ucStatusBar.ctl
'
'   Compatability:
'       Windows: 9x, ME, NT, 2K, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'       (Paul Caton - Self-Subclassser)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'       (Randy Birch - IsWinXP)
'           http://vbnet.mvps.org/code/system/getversionex.htm
'       (Dieter Otter - GetCurrentThemeName)
'           http://www.vbarchiv.net/archiv/tipp_805.html
'
'   Legal Copyright & Trademarks:
'       Copyright © 2007, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2007, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Advance Research Systems shall not be liable for
'       any incidental or consequential damages suffered by any use of this software.
'       This software is owned by Paul R. Territo, Ph.D and is free for use
'       in accordance with the terms of the License Agreement in the accompanying
'       documentation.
'
'   Contact Information:
'       For Technical Assistance:
'       pwterrito@insightbb.com
'
'-  Modification(s) History:
'
'       13Jul07 - Initial Usercontrol Build
'       14Jul07 - Fixed Aligment bug in the PanelAlign method which passed the wrong constant values to the
'                 drawing routines.
'               - Added Private StatusBar constants for clarity of the text alignments
'               - Added Theme Support (non-subclassed).
'               - Added usbClassic Theme Style for Win9x drawing support.
'               - Added Version property
'               - Added HitTest for Events to allow for determining which panel we are over
'               - Optimized the Drawing routines to prevent flicker on resize
'               - Added All Normal UserControl Events
'               - Added Panel Specific Events
'       15Jul07 - Added BoundControl Method for Binding External Objects into Panels
'               - Added Boundry checking for the Index property variables to ensure we are in bounds
'               - Optimized PaintPanels method to group activities by Icon or BoundObject states.
'               - Optimized BoundObject handling for resizing and auto hide if the control has
'                 a minimum width property...like ComboBoxes etc...
'               - Added Subclass support for SysColor, Theme, NonClient Paint uMsgs
'               - Added MouseEnter & MouseExit events with subclasser uMsgs
'               - Added Editable Property and updated AddPanel to reflect this
'               - Added txtEdit to allow for direct Panel modifications in DblClick.
'       16Jul07 - Added addtional drawing optimizations for painting in the IDE
'               - Added Theme Color Specific AlphaBlends for the top line of the gradient under XP LnF.
'               - Added alignmnet adjustments for Edit TextBox in usbClassic theme
'               - Added Auto selection of text on focus for Edit TextBox
'               - Fixed BoundObject Width in usbClassic theme
'               - Removed AutoHide of BoundObject when usbNoSize
'               - Fixed Grip highlight Painting for usbClassic theme
'       17Jul07 - Added painting refinements to the top gradient within PaintGradient
'       03Aug07 - Added Sizable property to allow for removal of this functionality
'       08Aug07 - Fixed Minor Redraw bug in the Refresh method which did not allow all panels
'                 to repaint correctly when updated.
'
'       Recode Control By Romeo91 for Better Subsclassing and Unicode Support for File And Text
'       10Dec13 - Repaint Subsclass Code from SelfSub 2.1 Paul Caton - http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=64867&lngWId=1.
'                 Added Unicode Support for FileOperation Dialog
'                 Added Unicode Support for Text Properties
'                 Added Unicode Support for Font Properties - Thanks Krool (http://www.vbforums.com/showthread.php?698563-CommonControls-%28Replacement-of-the-MS-common-controls%29)
'
'   Oroginal Build Date & Time: 8/3/2007 11:43:17 AM
'   Force Declarations
Option Explicit

'*************************************************************
'   API DECLARATION
'*************************************************************
Private Type RGBQUAD
    Blue                                As Byte
    Green                               As Byte
    Red                                 As Byte
    Alpha                               As Byte
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

'  for gradient painting and bitmap tiling
Private Type BITMAPINFOHEADER
    biSize                              As Long
    biWidth                             As Long
    biHeight                            As Long
    biPlanes                            As Integer
    biBitCount                          As Integer
    biCompression                       As Long
    biSizeImage                         As Long
    biXPelsPerMeter                     As Long
    biYPelsPerMeter                     As Long
    biClrUsed                           As Long
    biClrImportant                      As Long
End Type

Private Type BITMAPINFO8
    bmiHeader                           As BITMAPINFOHEADER
    bmiColors(255)                      As RGBQUAD
End Type

'   DrawEdge Message Constants
Private Const BDR_SUNKENOUTER   As Long = &H2
Private Const BDR_SUNKENINNER   As Long = &H8
Private Const BF_LEFT           As Long = &H1
Private Const BF_TOP            As Long = &H2
Private Const BF_RIGHT          As Long = &H4
Private Const BF_BOTTOM         As Long = &H8
Private Const BF_RECT           As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Declare Sub ReleaseCapture Lib "user32.dll" ()
Private Declare Sub CopyMemoryLong Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal XLeft As Long, ByVal YTop As Long, ByVal hIcon As Long, ByVal CXWidth As Long, ByVal CYWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, ByRef qRC As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InflateRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function CreateDIBSection8 Lib "gdi32.dll" Alias "CreateDIBSection" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO8, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long

'*************************************************************
'   DRAW TEXT
'*************************************************************
' --Formatting Text Consts
Private Const DT_LEFT          As Long = &H0
Private Const DT_CENTER        As Long = &H1
Private Const DT_RIGHT         As Long = &H2
Private Const DT_WORDBREAK     As Long = &H10
Private Const DT_RTLREADING    As Long = &H20000
Private Const DT_VCENTER       As Long = &H4
Private Const DT_WORD_ELLIPSIS As Long = &H40000

'   Private Local StatusBar Text Alignment Constants
Private Const DT_SB_LEFT = (DT_VCENTER Or DT_LEFT Or DT_WORD_ELLIPSIS Or DT_WORDBREAK)
Private Const DT_SB_CENTER = (DT_VCENTER Or DT_CENTER Or DT_WORD_ELLIPSIS Or DT_WORDBREAK)
Private Const DT_SB_RIGHT = (DT_VCENTER Or DT_RIGHT Or DT_WORD_ELLIPSIS Or DT_WORDBREAK)

Private Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

'*************************************************************
'   FONT PROPERTIES
'*************************************************************
Private Const LF_FACESIZE     As Long = 32

Private Type LOGFONT
    LFHeight As Long
    LFWidth As Long
    LFEscapement As Long
    LFOrientation As Long
    LFWeight As Long
    LFItalic As Byte
    LFUnderline As Byte
    LFStrikeOut As Byte
    LFCharset As Byte
    LFOutPrecision As Byte
    LFClipPrecision As Byte
    LFQuality As Byte
    LFPitchAndFamily As Byte
    LFFaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
End Type

'   MouseDown Message Constants for Corner Drag
Private Const HTBOTTOMRIGHT = 17
Private Const WM_NCLBUTTONDOWN = &HA1

'   Constants used by new transparent support in NT.
Private Const CAPS1 = 94                 '  other caps
Private Const C1_TRANSPARENT = &H1       '  new raster cap
Private Const NEWTRANSPARENT = 3         '  use with SetBkMode()

'   Ternary raster operations
Private Const SRCCOPY = &HCC0020         ' (DWORD) dest = source

'*************************************************************
'   CONTROL PROPERTIES
'*************************************************************
Public Enum usbAlignEnum
    usbLeft = &H0
    usbCenter = &H1
    usbRight = &H2
End Enum

#If False Then

    Const usbLeft = &H0
    Const usbCenter = &H1
    Const usbRight = &H2

#End If

Public Enum usbGripEnum
    usbNone = &H0
    usbSquare = &H1
    usbBars = &H2
End Enum

#If False Then

    Const usbNone = &H0
    Const usbSquare = &H1
    Const usbBars = &H2

#End If

Public Enum usbSizeEnum
    usbNoSize = &H0
    usbAutoSize = &H1
End Enum

#If False Then

    Const usbNoSize = &H0
    Const usbAutoSize = &H1

#End If

Public Enum usbStateEnum
    usbEnabled = &H0
    usbDisabled = &H1
End Enum

#If False Then

    Const usbEnabled = &H0
    Const usbDisabled = &H1

#End If

Public Enum usbThemeEnum
    usbAuto = &H0
    usbClassic = &H1
    usbBlue = &H2
    usbHomeStead = &H3
    usbMetallic = &H4
End Enum

#If False Then

    Const usbAuto = &H0
    Const usbClassic = &H1
    Const usbBlue = &H2
    Const usbHomeStead = &H3
    Const usbMetallic = &H4

#End If

'   Private StatusBar Item Type
Private Type PanelItem
    Alignment                   As Long
    AutoSize                    As Boolean
    BoundObject                 As Object
    BoundParent                 As Long
    BoundSize                   As usbSizeEnum
    Editable                    As Boolean
    ForeColor                   As OLE_COLOR
    Font                        As StdFont
    Icon                        As StdPicture
    IconState                   As usbStateEnum
    ItemRect                    As RECT
    MaskColor                   As OLE_COLOR
    Text                        As String
    ToolTipText                 As String
    UseMaskColor                As Boolean
    Width                       As Long
End Type

Private m_ActivePanel     As Long             'Current Active Panel
Private m_BackColor       As OLE_COLOR        'UserControl BackColor
Private m_Forecolor       As OLE_COLOR        'UserControl ForeColor
Private m_Font            As StdFont          'UserControl Font
Private m_GripRect        As RECT             'Grip Retangle
Private m_GripShape       As usbGripEnum      'Grip Shape...Auto Set when Theme is Set
Private m_Sizable         As Boolean          'Resizable
Private m_PanelCount      As Long             'Panel Count
Private m_PanelItems()    As PanelItem        'Panel Items
Private m_Theme           As usbThemeEnum     'Theme Set by the User
Private m_iTheme          As usbThemeEnum     'Theme Stored internally for determination of named themes + auto equivelant
Private m_bIsWinXpOrLater As Boolean

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PanelClick(Index As Long)
Public Event PanelDblClick(Index As Long)
Public Event PanelMouseDown(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PanelMouseMove(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PanelMouseUp(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    BackColor = m_BackColor
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.BackColor", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'Description: Use this color for drawing
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_BackColor = NewValue
    UserControl.BackColor = m_BackColor
    Refresh
    PropertyChanged "BackColor"
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.BackColor", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Font
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Font() As StdFont

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set Font = m_Font
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.Font", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Font
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewFont (StdFont)
'!--------------------------------------------------------------------------------
Public Property Set Font(ByVal NewFont As StdFont)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set m_Font = NewFont
    Refresh
    PropertyChanged "Font"
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.Font", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ForeColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ForeColor = m_Forecolor
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.ForeColor", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'Description: Use this color for drawing
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_BackColor = NewValue
    UserControl.ForeColor = m_Forecolor
    Refresh
    PropertyChanged "ForeColor"
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.ForeColor", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property GripShape
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get GripShape() As usbGripEnum

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    GripShape = m_GripShape
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.GripShape", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property GripShape
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lShape (usbGripEnum)
'!--------------------------------------------------------------------------------
Public Property Let GripShape(lShape As usbGripEnum)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    '   Check to see if this changed, otherwise we get an
    '   "Out of Stack Space" error with recursive changes...
    If lShape <> m_GripShape Then
        m_GripShape = lShape
        Refresh
        PropertyChanged "GripShape"
    End If

Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.GripShape", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelAlignment
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'!--------------------------------------------------------------------------------
Public Property Get PanelAlignment(ByVal Index As Long) As usbAlignEnum

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount

    Select Case m_PanelItems(Index).Alignment

        Case DT_SB_LEFT
            PanelAlignment = usbLeft

        Case DT_SB_CENTER
            PanelAlignment = usbCenter

        Case DT_SB_RIGHT
            PanelAlignment = usbRight
    End Select

Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelAlignment", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelAlignment
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'                              NewValue (usbAlignEnum)
'!--------------------------------------------------------------------------------
Public Property Let PanelAlignment(ByVal Index As Long, ByVal NewValue As usbAlignEnum)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then

        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount

    Select Case NewValue

        Case usbLeft
            m_PanelItems(Index).Alignment = DT_SB_LEFT

        Case usbCenter
            m_PanelItems(Index).Alignment = DT_SB_CENTER

        Case usbRight
            m_PanelItems(Index).Alignment = DT_SB_RIGHT
    End Select

    Refresh
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelAlignment", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelAutoSize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'!--------------------------------------------------------------------------------
Public Property Get PanelAutoSize(ByVal Index As Long) As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelAutoSize = m_PanelItems(Index).AutoSize
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelAutoSize", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelAutoSize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'                              NewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let PanelAutoSize(ByVal Index As Long, ByVal NewValue As Boolean)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then

        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).AutoSize = NewValue
    Refresh
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelAutoSize", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelCount
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get PanelCount() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_PanelCount = UBoundEx(m_PanelItems)
    PanelCount = m_PanelCount
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelCount", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelEditable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'!--------------------------------------------------------------------------------
Public Property Get PanelEditable(ByVal Index As Long) As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelEditable = m_PanelItems(Index).Editable
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelEditable", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelEditable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'                              NewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let PanelEditable(ByVal Index As Long, ByVal NewValue As Boolean)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then

        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).Editable = NewValue
    Refresh
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelEditable", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelFont
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'!--------------------------------------------------------------------------------
Public Property Get PanelFont(ByVal Index As Long) As StdFont

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    Set PanelFont = m_PanelItems(Index).Font
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelFont", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelFont
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'                              NewItem (StdFont)
'!--------------------------------------------------------------------------------
Public Property Let PanelFont(ByVal Index As Long, ByVal NewItem As StdFont)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then

        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    Set m_PanelItems(Index).Font = NewItem
    Refresh
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelFont", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'!--------------------------------------------------------------------------------
Public Property Get PanelForeColor(ByVal Index As Long) As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelForeColor = m_PanelItems(Index).ForeColor
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelForeColor", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'                              NewItem (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let PanelForeColor(ByVal Index As Long, ByVal NewItem As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then

        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).ForeColor = NewItem
    Refresh
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelForeColor", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelIcon
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'!--------------------------------------------------------------------------------
Public Property Get PanelIcon(ByVal Index As Long) As StdPicture

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    Set PanelIcon = m_PanelItems(Index).Icon
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelIcon", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelIcon
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'                              NewItem (StdPicture)
'!--------------------------------------------------------------------------------
Public Property Set PanelIcon(ByVal Index As Long, ByVal NewItem As StdPicture)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then

        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    Set m_PanelItems(Index).Icon = NewItem
    Refresh
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelIcon", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'!--------------------------------------------------------------------------------
Public Property Get PanelText(ByVal Index As Long) As String

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelText = m_PanelItems(Index).Text
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelText", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'                              NewItem (String)
'!--------------------------------------------------------------------------------
Public Property Let PanelText(ByVal Index As Long, ByVal NewItem As String)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then

        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).Text = NewItem
    Refresh
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelText", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelToolTipText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'!--------------------------------------------------------------------------------
Public Property Get PanelToolTipText(ByVal Index As Long) As String

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelToolTipText = m_PanelItems(Index).ToolTipText
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelToolTipText", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelToolTipText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'                              NewValue (String)
'!--------------------------------------------------------------------------------
Public Property Let PanelToolTipText(ByVal Index As Long, NewValue As String)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then

        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).ToolTipText = NewValue
    Refresh
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelToolTipText", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelWidth
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'!--------------------------------------------------------------------------------
Public Property Get PanelWidth(ByVal Index As Long) As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelWidth = m_PanelItems(Index).Width
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelWidth", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PanelWidth
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'                              NewItem (Long)
'!--------------------------------------------------------------------------------
Public Property Let PanelWidth(ByVal Index As Long, ByVal NewItem As Long)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then

        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).Width = NewItem
    Refresh
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PanelWidth", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

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
'! Procedure   (Функция)   :   Property Sizable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Sizable() As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Sizable = m_Sizable
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.Sizable", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Sizable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let Sizable(ByVal NewValue As Boolean)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_Sizable = NewValue

    If m_Sizable Then
        If IsWinXPOrLater Then
            m_GripShape = usbSquare
        Else
            m_GripShape = usbBars
        End If

    Else
        m_GripShape = usbNone
    End If

    Refresh
    PropertyChanged "Sizable"
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.Sizable", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Theme
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Theme() As usbThemeEnum

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Theme = m_Theme
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.Theme", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Theme
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (usbThemeEnum)
'!--------------------------------------------------------------------------------
Public Property Let Theme(ByVal NewValue As usbThemeEnum)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_Theme = NewValue
    Refresh
    PropertyChanged "Theme"
Prop_ErrHandlerExit:

    Exit Property

Prop_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.Theme", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function AddPanel
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sText (String)
'                              uTextAlign (usbAlignEnum = usbLeft)
'                              bAutoSize (Boolean = True)
'                              bEditable (Boolean)
'                              oIcon (StdPicture)
'                              bIconState (usbStateEnum = usbEnabled)
'                              bUseMaskColor (Boolean)
'                              lMaskColor (OLE_COLOR = vbMagenta)
'                              lForeColor (OLE_COLOR = vbButtonText)
'                              oFont (StdFont)
'                              sToolTipText (String)
'                              lWidth (Long = 40)
'!--------------------------------------------------------------------------------
Public Function AddPanel(Optional ByVal sText As String, Optional ByVal uTextAlign As usbAlignEnum = usbLeft, Optional ByVal bAutoSize As Boolean = True, Optional ByVal bEditable As Boolean, Optional ByVal oIcon As StdPicture, Optional ByVal _
                            bIconState As usbStateEnum = usbEnabled, Optional ByVal bUseMaskColor As Boolean, Optional ByVal lMaskColor As OLE_COLOR = vbMagenta, Optional ByVal lForeColor As OLE_COLOR = vbButtonText, Optional ByVal oFont As _
                            StdFont, Optional ByVal sToolTipText As String, Optional ByVal lWidth As Long = 40) As Boolean

    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    m_PanelCount = m_PanelCount + 1

    ReDim Preserve m_PanelItems(1 To m_PanelCount)

    With m_PanelItems(m_PanelCount)

        Select Case uTextAlign

            Case usbLeft
                .Alignment = DT_SB_LEFT

            Case usbCenter
                .Alignment = DT_SB_CENTER

            Case usbRight
                .Alignment = DT_SB_RIGHT
        End Select

        .AutoSize = bAutoSize
        .Editable = bEditable

        If Not oFont Is Nothing Then
            Set .Font = oFont
        Else

            If Not m_Font Is Nothing Then
                Set .Font = m_Font
            Else
                Set .Font = Ambient.Font
            End If
        End If

        .ForeColor = lForeColor

        If Not oIcon Is Nothing Then
            Set .Icon = oIcon
        End If

        .IconState = bIconState
        .MaskColor = lMaskColor
        .Text = sText
        .ToolTipText = sToolTipText
        .UseMaskColor = bUseMaskColor

        If lWidth Then
            .Width = lWidth
        Else
            .Width = 40
        End If

    End With

    Refresh
Func_ErrHandlerExit:

    Exit Function

Func_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.AddPanel", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Func_ErrHandlerExit:

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function AlphaBlend
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   FirstColor (Long)
'                              SecondColor (Long)
'                              AlphaValue (Long)
'!--------------------------------------------------------------------------------
Private Function AlphaBlend(ByVal FirstColor As Long, ByVal SecondColor As Long, ByVal AlphaValue As Long) As Long

    Dim iForeColor As RGBQUAD
    Dim iBackColor As RGBQUAD

    OleTranslateColor FirstColor, 0, ByVal VarPtr(iForeColor)
    OleTranslateColor SecondColor, 0, ByVal VarPtr(iBackColor)

    With iForeColor
        .Red = (.Red * AlphaValue + iBackColor.Red * (255 - AlphaValue)) / 255
        .Green = (.Green * AlphaValue + iBackColor.Green * (255 - AlphaValue)) / 255
        .Blue = (.Blue * AlphaValue + iBackColor.Blue * (255 - AlphaValue)) / 255
    End With

    CopyMemoryLong VarPtr(AlphaBlend), VarPtr(iForeColor), 4
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub APILine
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   X1 (Long)
'                              Y1 (Long)
'                              X2 (Long)
'                              Y2 (Long)
'                              lColor (Long)
'!--------------------------------------------------------------------------------
Private Sub APILine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lColor As Long)

    'Use the API LineTo for Fast Drawing
    Dim PT   As POINTAPI
    Dim hPen As Long, hPenOld               As Long

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    lColor = TranslateColor(lColor)
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(UserControl.hDC, hPen)
    MoveToEx UserControl.hDC, X1, Y1, PT
    LineTo UserControl.hDC, X2, Y2
    SelectObject UserControl.hDC, hPenOld
    DeleteObject hPen
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.APILine", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub BoundControl
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Long)
'                              Control (Object)
'                              SizeMethod (usbSizeEnum)
'!--------------------------------------------------------------------------------
Public Sub BoundControl(ByVal Index As Long, Control As Object, ByVal SizeMethod As usbSizeEnum)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If m_PanelCount < 1 Then

        Exit Sub

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    If Not Control Is Nothing Then
        m_PanelItems(Index).BoundParent = GetParent(Control.hWnd)
        Set m_PanelItems(Index).BoundObject = Control
        SetParent m_PanelItems(Index).BoundObject.hWnd, UserControl.hWnd
    Else

        '   See if the control exists, if so, then we should set the parent back
        '   and destroy the reference to it...
        If Not m_PanelItems(Index).BoundObject Is Nothing Then
            SetParent m_PanelItems(Index).BoundObject.hWnd, m_PanelItems(Index).BoundParent
            Set m_PanelItems(Index).BoundObject = Nothing
        End If
    End If

    m_PanelItems(Index).BoundSize = SizeMethod
    Refresh
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.BoundControl", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Clear
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Clear()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Dim lpRect As RECT
    Dim hBrush As Long
    Dim lColor As Long

    With lpRect
        .Left = 0
        .Top = 0
        .Right = ScaleWidth
        .Bottom = ScaleHeight
    End With

    lColor = TranslateColor(m_BackColor)
    hBrush = CreateSolidBrush(lColor)
    Call FillRect(UserControl.hDC, lpRect, hBrush)
    Call DeleteObject(hBrush)
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.Clear", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetPanelIndex
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function GetPanelIndex() As Long

    Dim I      As Long
    Dim tPt    As POINTAPI
    Dim lpRect As RECT

    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    If m_PanelCount Then
        '   Get our position
        Call GetCursorPos(tPt)
        '   Convert coordinates
        Call ScreenToClient(UserControl.hWnd, tPt)

        '   Loop Over the RECTs a see if it is in
        For I = 1 To m_PanelCount
            lpRect = m_PanelItems(I).ItemRect

            If Not m_PanelItems(I).Icon Is Nothing Then
                If m_PanelItems(I).Alignment = DT_SB_LEFT Then
                    OffsetRect lpRect, -16, 0
                ElseIf m_PanelItems(I).Alignment = DT_SB_CENTER Then
                    OffsetRect lpRect, -8, 0
                ElseIf m_PanelItems(I).Alignment = DT_SB_RIGHT Then
                    InflateRect lpRect, 2, 0
                End If
            End If

            If I > 1 Then
                If (m_PanelItems(I - 1).ItemRect.Right + 10) < lpRect.Left Then
                    OffsetRect lpRect, -8, 0
                    InflateRect lpRect, 6, 0
                End If
            End If

            If PtInRect(lpRect, tPt.X, tPt.Y) Then
                GetPanelIndex = I

                Exit For

            End If

        Next

    End If

Func_ErrHandlerExit:

    Exit Function

Func_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.GetPanelIndex", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Func_ErrHandlerExit:

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
'! Procedure   (Функция)   :   Sub GrayBlt
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   hDstDC (Long)
'                              hSrcDC (Long)
'                              nWidth (Long)
'                              nHeight (Long)
'!--------------------------------------------------------------------------------
Private Sub GrayBlt(ByVal hDstDC As Long, ByVal hSrcDC As Long, ByVal nWidth As Long, ByVal nHeight As Long)

    Dim MakePal As Long
    Dim DIBInf  As BITMAPINFO8
    Dim gsDIB   As Long
    Dim hTmpDC  As Long
    Dim OldDIB  As Long

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    hTmpDC = CreateCompatibleDC(hSrcDC)

    With DIBInf
        With .bmiHeader
            ' Same size as picture
            .biWidth = nWidth
            .biHeight = nHeight
            .biBitCount = 8
            .biPlanes = 1
            .biClrUsed = 256
            .biClrImportant = 256
            .biSize = Len(DIBInf.bmiHeader)
        End With

        ' Palette is Greyscale
        For MakePal = 0 To 255

            With .bmiColors(MakePal)
                .Red = MakePal
                .Green = MakePal
                .Blue = MakePal
            End With

        Next

    End With

    gsDIB = CreateDIBSection8(hTmpDC, DIBInf, 0, ByVal 0&, 0, 0)

    If (hTmpDC) Then
        ' Validate and select DIB
        OldDIB = SelectObject(hTmpDC, gsDIB)
        ' Draw original picture to the greyscale DIB
        BitBlt hTmpDC, 0, 0, DIBInf.bmiHeader.biWidth, DIBInf.bmiHeader.biHeight, hSrcDC, 0, 0, vbSrcCopy
        ' Draw the greyscale image back to the hDC
        BitBlt hDstDC, 0, 0, DIBInf.bmiHeader.biWidth, DIBInf.bmiHeader.biHeight, hTmpDC, 0, 0, vbSrcCopy
        ' Clean up DIB
        SelectObject hTmpDC, OldDIB
        DeleteObject gsDIB
        DeleteObject hTmpDC
    End If

Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.GrayBlt", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PaintGradients
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub PaintGradients()

    Dim I       As Long
    Dim Y1      As Long
    Dim BtnFace As Long
    Dim lColor  As Long

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With UserControl

        If (m_iTheme = usbClassic) Then
            '   Clear the control to start using the
            '   optimized repaint method instead Cls to avoid flicker
            Clear
        Else
            '   Get the BackColor and Offset it by 2 Units
            BtnFace = ShiftColor(.BackColor, -&H1)
            '   Clear the control to start using the
            '   optimized repaint method instead Cls to avoid flicker
            Clear

            '   Draw the Smooth Gradient across the whole control
            For I = 0 To ScaleHeight
                Y1 = I
                APILine 0, Y1, .ScaleWidth, Y1, AlphaBlend(&HFFFFFF, BtnFace, (I / ScaleHeight) * 48)
            Next

            '   Draw The Top Lines
            Select Case m_iTheme

                Case usbBlue
                    lColor = AlphaBlend(ShiftColor(BtnFace, -&H40), &HB99D7F, 128)

                Case usbHomeStead
                    lColor = AlphaBlend(ShiftColor(BtnFace, -&H40), &H69A18B, 128)

                Case usbMetallic
                    lColor = AlphaBlend(ShiftColor(BtnFace, -&H40), &H947C7C, 128)

                Case Else
                    lColor = ShiftColor(BtnFace, -&H50)
            End Select

            APILine 0, 0, .ScaleWidth, 0, &HFFFFFF
            'AlphaBlend(ShiftColor(BtnFace, -&H8), &HFFFFFF, 128)
            APILine 0, 1, .ScaleWidth, 1, lColor
            '   Draw the Top Gradient
            APILine 0, 2, .ScaleWidth, 2, ShiftColor(BtnFace, -&H25)
            APILine 0, 3, .ScaleWidth, 3, ShiftColor(BtnFace, -&H9)

            '   Draw the Bottom Gradient
            For I = 0 To 5
                Y1 = .ScaleHeight - 5 + I
                APILine 0, Y1, .ScaleWidth, Y1, ShiftColor(BtnFace, -&H1 * ((((I / 3) * 100) * .ScaleHeight) / 100))
            Next

        End If

    End With

Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PaintGradients", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PaintGrip
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub PaintGrip()

    Dim AdjWidth  As Long
    Dim AdjHeight As Long

    '   Custom reoutine, to paint/repaint the shapes on the
    '   screen to represent the Grip Style selected...
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With UserControl
        AdjWidth = (.ScaleWidth - 15)
        AdjHeight = (.ScaleHeight - 16)

        '   See if this is XP, if so then paint the correct Resize Button
        If (m_GripShape = usbSquare) And (m_iTheme <> usbClassic) Then
            '   Paint the Shadows first....
            .ForeColor = vbWhite
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 4, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 4, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 5, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 5, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 9, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 9, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 13, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 13, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 9, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 9, .ForeColor
            SetPixelV .hDC, AdjWidth + 9, AdjHeight + 9, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 9, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 13, .ForeColor
            SetPixelV .hDC, AdjWidth + 9, AdjHeight + 13, .ForeColor
            SetPixelV .hDC, AdjWidth + 4, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 5, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 4, AdjHeight + 13, .ForeColor
            SetPixelV .hDC, AdjWidth + 5, AdjHeight + 13, .ForeColor
            '   Shift the Color to be a Blend of the BackColor and Medium Grey
            .ForeColor = AlphaBlend(&H909090, .BackColor, 128)
            '   Paint the Grips Next....
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 3, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 3, .ForeColor
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 4, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 4, .ForeColor
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 7, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 7, .ForeColor
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 7, AdjHeight + 7, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 7, .ForeColor
            SetPixelV .hDC, AdjWidth + 7, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 7, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 7, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 3, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 4, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 3, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 4, AdjHeight + 12, .ForeColor
        ElseIf (m_GripShape = usbBars) And (m_iTheme = usbClassic) Then
            '   Draw the White Highlight Lines First in groups of two
            .ForeColor = vbWhite
            APILine AdjWidth + 12, AdjHeight + 13, AdjWidth + 14, AdjHeight + 11, .ForeColor
            APILine AdjWidth + 9, AdjHeight + 13, AdjWidth + 14, AdjHeight + 8, .ForeColor
            APILine AdjWidth + 6, AdjHeight + 13, AdjWidth + 14, AdjHeight + 5, .ForeColor
            APILine AdjWidth + 3, AdjHeight + 13, AdjWidth + 14, AdjHeight + 2, .ForeColor
            '   Now Draw the Lowlight Lines in groups of two
            .ForeColor = AlphaBlend(vbWhite, ShiftColor(.BackColor, -&H70), 128)
            APILine AdjWidth + 13, AdjHeight + 14, AdjWidth + 14, AdjHeight + 13, .ForeColor
            APILine AdjWidth + 12, AdjHeight + 14, AdjWidth + 14, AdjHeight + 12, .ForeColor
            APILine AdjWidth + 10, AdjHeight + 14, AdjWidth + 14, AdjHeight + 10, .ForeColor
            APILine AdjWidth + 9, AdjHeight + 14, AdjWidth + 14, AdjHeight + 9, .ForeColor
            APILine AdjWidth + 7, AdjHeight + 14, AdjWidth + 14, AdjHeight + 7, .ForeColor
            APILine AdjWidth + 6, AdjHeight + 14, AdjWidth + 14, AdjHeight + 6, .ForeColor
            APILine AdjWidth + 4, AdjHeight + 14, AdjWidth + 14, AdjHeight + 4, .ForeColor
            APILine AdjWidth + 3, AdjHeight + 14, AdjWidth + 14, AdjHeight + 3, .ForeColor
        End If

    End With

Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PaintGrip", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PaintPanels
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub PaintPanels()

    Dim I           As Long
    Dim lX          As Long
    Dim lForeColor  As Long
    Dim lIconOffset As Long
    Dim lGripSize   As Long
    Dim bMinWidth   As Boolean

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    lForeColor = UserControl.ForeColor
    lIconOffset = 18

    If (m_iTheme = usbClassic) Then
        lGripSize = 16
    Else
        lGripSize = 18
    End If

    For I = 1 To PanelCount

        With m_PanelItems(I)
            '   Set the Individual ForeColor & Font
            UserControl.ForeColor = .ForeColor
            Set UserControl.Font = .Font

            '   Autosize the Text + Icon?
            If .AutoSize Then
                '   Set the Left & Top
                .ItemRect.Left = lX
                .ItemRect.Top = 5

                '   Do we have a valid Icon?
                If .Icon Is Nothing Then
                    '   Compute the Distance we need to Extend the Rect
                    .ItemRect.Right = lX + TextWidth(.Text) + 8
                Else
                    '   Compute the Distance we need to Extend the Rect + Icon Distance
                    .ItemRect.Right = lX + TextWidth(.Text) + 8 + lIconOffset
                End If

                '   Set the Bottom of the Rect
                .ItemRect.Bottom = ScaleHeight - 5

                '   Use a default for blank text
                If LenB(.Text) Then
                    lX = .ItemRect.Right
                Else
                    lX = lX + 20
                End If

                '   Check to see if the control is smaller then the
                '   right most separator, if so correct it
                If lX >= (ScaleWidth - lGripSize) Then
                    '   Yep, so make the Rect scaller to match
                    .ItemRect.Right = (ScaleWidth - lGripSize)
                    lX = .ItemRect.Right
                End If

            Else
                '   Set the Left & Top
                .ItemRect.Left = lX
                .ItemRect.Top = 5

                '   Do we have a valid Icon?
                If .Icon Is Nothing Then
                    '   Compute the Distance we need to Extend the Rect
                    .ItemRect.Right = lX + .Width
                Else
                    '   Compute the Distance we need to Extend the Rect + Icon Distance
                    .ItemRect.Right = lX + .Width + lIconOffset
                End If

                '   Set the Bottom of the Rect
                .ItemRect.Bottom = ScaleHeight - 5
                lX = .ItemRect.Right

                '   Check to see if the control is smaller then the
                '   right most separator, if so correct it
                If lX >= (ScaleWidth - lGripSize) Then
                    '   Yep, so make the Rect scaller to match
                    .ItemRect.Right = (ScaleWidth - lGripSize)
                    lX = .ItemRect.Right
                End If
            End If

            '   Now draw the Theme Based Borders....
            If (m_iTheme = usbClassic) Then
                '   Draw the Panels as Sunken Boxes as per 9x LnF
                InflateRect .ItemRect, 0, 3
                DrawEdge UserControl.hDC, .ItemRect, EDGE_SUNKEN, BF_RECT
                InflateRect .ItemRect, -5, -3
            Else
                '   Draw the Lines for the Dividors as per XP LnF
                APILine lX, .ItemRect.Top, lX, .ItemRect.Bottom, AlphaBlend(&H909090, m_BackColor, 128)
                APILine lX + 1, .ItemRect.Top, lX + 1, .ItemRect.Bottom, vbWhite
                '   Decrease the RECT by 4
                InflateRect .ItemRect, -4, 0
            End If

            '   Does this have a bound object?
            If .BoundObject Is Nothing Then

                '   Do we have a Picture?
                If Not .Icon Is Nothing Then

                    '   Adjust the Initial Items RECT to line up correctly
                    If I = 1 Then
                        OffsetRect .ItemRect, -2, 0
                    End If

                    '   See if the size of the StatusBar is too small for an Icon + Padding
                    If (.ItemRect.Left + lIconOffset) <= (ScaleWidth - lGripSize) Then
                        '   Yep, so paint it centered vertically
                        TransBltEx UserControl.hDC, .ItemRect.Left, ScaleHeight \ 2 - 8, 16, 16, .Icon, 0, 0, BackColor, IIf(.IconState = usbEnabled, False, True)
                        '   Now offset th RECT so the text starts in the corect position
                        OffsetRect .ItemRect, lIconOffset \ 2, 0
                        InflateRect .ItemRect, -lIconOffset \ 2, 0

                        '   Perform adjustments as needed depending on Aligment
                        If .Alignment = DT_SB_LEFT Then

                            '   Adjust the Right most extent if the item is smaller
                            '   than the RECT....
                            If lX >= (ScaleWidth - lGripSize) Then
                                '   Yep, so make the Rect scaller to match
                                .ItemRect.Right = (ScaleWidth - lGripSize)
                            End If

                        ElseIf .Alignment = DT_SB_CENTER Then

                            '   Adjust the Right most extent if the item is smaller
                            '   than the RECT....
                            If lX >= (ScaleWidth - lGripSize) Then
                                '   Yep, so make the Rect scaller to match
                                OffsetRect .ItemRect, 0, 0
                                .ItemRect.Right = (ScaleWidth - lGripSize) - 2
                            End If

                        ElseIf .Alignment = DT_SB_RIGHT Then

                            '   Adjust the Right most extent if the item is smaller
                            '   than the RECT....
                            If lX >= (ScaleWidth - lGripSize) Then
                                '   Yep, so make the Rect scaller to match
                                OffsetRect .ItemRect, lIconOffset, 0
                                InflateRect .ItemRect, lIconOffset, 0
                                .ItemRect.Right = (ScaleWidth - lGripSize) - 2
                            End If
                        End If
                    End If

                    '   See if the size of the StatusBar is too small for an Icon + Padding
                    '   if so then we don't want to paint the text where the icon was located
                    If (.Alignment = DT_SB_LEFT) Or (.Alignment = DT_SB_RIGHT) Then

                        '   If there is enough room, print the text
                        If ((.ItemRect.Left + lIconOffset) <= (ScaleWidth - lGripSize)) Or ((.ItemRect.Right - .ItemRect.Left) > 16) Then
                            'DrawText UserControl.hDC, .Text, -1, .ItemRect, .Alignment
                            DrawTextW UserControl.hDC, StrPtr(.Text & vbNullChar), -1, .ItemRect, .Alignment
                        End If

                    Else

                        '   If there is enough room, print the text
                        If (.ItemRect.Left + lIconOffset \ 2) <= (ScaleWidth - lGripSize) Or ((.ItemRect.Right - .ItemRect.Left) > 16) Then
                            'DrawText UserControl.hDC, .Text, -1, .ItemRect, .Alignment
                            DrawTextW UserControl.hDC, StrPtr(.Text & vbNullChar), -1, .ItemRect, .Alignment
                        End If
                    End If

                Else

                    '   If there is enough room, print the text
                    If (.ItemRect.Left + 2) <= (ScaleWidth - lGripSize) Then
                        'DrawText UserControl.hDC, .Text, -1, .ItemRect, .Alignment
                        DrawTextW UserControl.hDC, StrPtr(.Text & vbNullChar), -1, .ItemRect, .Alignment
                    End If
                End If

            Else

                '   Set the Bound Object onto the Control
                '   Handle errors quietly in this section as we are late bound
                '   so it is hard to predict if all controls will support certain
                '   object interfaces....
                On Error Resume Next

                '   Only deal with real controls
                If Not .BoundObject Is Nothing Then

                    '   Is this going to be resized or not....
                    If .BoundSize = usbNoSize Then

                        '   Keep the Width, but set the Left, Top and Height
                        With .BoundObject
                            .Left = m_PanelItems(I).ItemRect.Left * Screen.TwipsPerPixelX
                            .Top = m_PanelItems(I).ItemRect.Top * Screen.TwipsPerPixelY
                            .Height = 16 * Screen.TwipsPerPixelY
                            '   Under development....;-)
                            '   Should be hidden if too small to fit the control..
                            'If (.Width <= ((m_PanelItems(i).ItemRect.Right - m_PanelItems(i).ItemRect.Left)) * Screen.TwipsPerPixelX) Then
                            '    .Visible = False
                            'Else
                            '    .Visible = True
                            'End If
                            .ZOrder 0
                        End With

                    Else

                        With .BoundObject

                            '   Resize all properties to make it fit
                            If m_iTheme <> usbClassic Then
                                .Left = (m_PanelItems(I).ItemRect.Left) * Screen.TwipsPerPixelX
                                .Width = ((m_PanelItems(I).ItemRect.Right - m_PanelItems(I).ItemRect.Left)) * Screen.TwipsPerPixelX

                                '   See if we were avel to resize the controls width, if not
                                '   then the control might have a minimum width (i.e. ComboBox)
                                '   so we can simply use this as an indicator to hide the control...
                                If (.Width <> (((m_PanelItems(I).ItemRect.Right - m_PanelItems(I).ItemRect.Left)) * Screen.TwipsPerPixelX)) Then
                                    bMinWidth = True
                                Else
                                    bMinWidth = False
                                End If

                            Else
                                .Left = (m_PanelItems(I).ItemRect.Left - 4) * Screen.TwipsPerPixelX
                                .Width = ((m_PanelItems(I).ItemRect.Right - m_PanelItems(I).ItemRect.Left) + 9) * Screen.TwipsPerPixelX

                                '   See if we were avel to resize the controls width, if not
                                '   then the control might have a minimum width (i.e. ComboBox)
                                '   so we can simply use this as an indicator to hide the control...
                                If (.Width <> (((m_PanelItems(I).ItemRect.Right - m_PanelItems(I).ItemRect.Left) + 9) * Screen.TwipsPerPixelX)) Then
                                    bMinWidth = True
                                Else
                                    bMinWidth = False
                                End If
                            End If

                            .Height = 16 * Screen.TwipsPerPixelY
                            .Top = Height \ 2 - .Height \ 2

                            If (.Width <= 30) Or (bMinWidth = True) Then
                                .Visible = False
                            Else
                                .Visible = True
                            End If

                            .ZOrder 0
                        End With

                    End If
                End If

                '   Turn the normal Error handing back on....
                On Error GoTo 0

            End If

        End With

    Next

    '   Set the ForeColor back...
    UserControl.ForeColor = lForeColor
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PaintPanels", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function PtInRect
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lpRect (RECT)
'                              X (Long)
'                              Y (Long)
'!--------------------------------------------------------------------------------
Private Function PtInRect(ByRef lpRect As RECT, X As Long, Y As Long) As Boolean

    '   This is a replacemnt for the PtInRect API call which seems to always
    '   return 0 depite the X & Y Points being in the RECT...
    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    If (X >= lpRect.Left) Then
        If (X <= lpRect.Right) Then
            If (Y >= lpRect.Top) Then
                If (Y <= lpRect.Bottom) Then
                    PtInRect = True
                End If
            End If
        End If
    End If

Func_ErrHandlerExit:

    Exit Function

Func_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.PtInRect", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Func_ErrHandlerExit:

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Refresh
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Refresh()

    Dim AutoTheme As String

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Select Case m_Theme

        Case [usbAuto]
            AutoTheme = GetThemeInfo

            Select Case AutoTheme

                Case "None", "UxTheme_Error"
                    m_iTheme = usbClassic

                    If m_GripShape <> usbNone Then
                        m_GripShape = usbBars
                    End If

                Case "NormalColor"
                    m_iTheme = usbBlue

                    If m_GripShape <> usbNone Then
                        m_GripShape = usbSquare
                    End If

                Case "HomeStead"
                    m_iTheme = usbHomeStead

                    If m_GripShape <> usbNone Then
                        m_GripShape = usbSquare
                    End If

                Case "Metallic"
                    m_iTheme = usbMetallic

                    If m_GripShape <> usbNone Then
                        m_GripShape = usbSquare
                    End If

                Case Else
                    m_iTheme = usbBlue

                    If m_GripShape <> usbNone Then
                        m_GripShape = usbSquare
                    End If

            End Select

        Case [usbClassic]
            m_iTheme = usbClassic

            If m_GripShape <> usbNone Then
                m_GripShape = usbBars
            End If

        Case [usbBlue]
            m_iTheme = usbBlue

            If m_GripShape <> usbNone Then
                m_GripShape = usbSquare
            End If

        Case [usbHomeStead]
            m_iTheme = usbHomeStead

            If m_GripShape <> usbNone Then
                m_GripShape = usbSquare
            End If

        Case [usbMetallic]
            m_iTheme = usbMetallic

            If m_GripShape <> usbNone Then
                m_GripShape = usbSquare
            End If

        Case Else
            m_iTheme = usbBlue

            If m_GripShape <> usbNone Then
                m_GripShape = usbSquare
            End If

    End Select

    '   Paint the Gradient for the whole control
    PaintGradients
    '   Now Paint the Grip according to style
    PaintGrip
    '   Paint the Divisions which represent the panels
    PaintPanels

    '   Only refresh if in the IDE (Otherwise it will Flicker!!)
    If Not RunMode Then
        AutoRedraw = False
    Else
        AutoRedraw = True
        '   Refresh the Window
        UserControl.Refresh
    End If

Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.Refresh", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ShiftColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Color (Long)
'                              Value (Long)
'!--------------------------------------------------------------------------------
Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long) As Long

    Dim lR As Long
    Dim lg As Long
    Dim LB As Long

    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    Color = TranslateColor(Color)
    lR = (Color And &HFF) + Value
    lg = ((Color \ &H100) Mod &H100) + Value
    LB = ((Color \ &H10000) Mod &H100)
    LB = LB + ((LB * Value) \ &HC0)

    If Value Then
        If lR > 255 Then lR = 255
        If lg > 255 Then lg = 255
        If LB > 255 Then LB = 255
    ElseIf Value < 0 Then

        If lR < 0 Then lR = 0
        If lg < 0 Then lg = 0
        If LB < 0 Then LB = 0
    End If

    ShiftColor = lR + 256& * lg + 65536 * LB
Func_ErrHandlerExit:

    Exit Function

Func_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.ShiftColor", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Func_ErrHandlerExit:

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
'! Procedure   (Функция)   :   Sub TransBltEx
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   hDestDC (Long)
'                              X (Long)
'                              Y (Long)
'                              nWidth (Long)
'                              nHeight (Long)
'                              hSrcImg (StdPicture)
'                              XSrc (Long)
'                              YSrc (Long)
'                              TransColor (Long)
'                              Disabled (Boolean)
'!--------------------------------------------------------------------------------
Public Sub TransBltEx(ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcImg As StdPicture, ByVal XSrc As Long, ByVal YSrc As Long, ByVal transColor As Long, ByVal Disabled As Boolean)

    '
    '   32-Bit Transparent BitBlt Function
    '   Written by Karl E. Peterson, 9/20/96.
    '   Portions borrowed and modified from KB.
    '   Other portions modified following input from users. <g>
    '
    '   Modified by Paul R. Territo, Ph.D 02Apr07 to allow
    '   passing of a StdPicture object and populating a private
    '   hSrcDC instead of the original method which passed the hScrDC
    '
    '   Modified by Paul R. Territo, Ph.D 11Apr07 to allow for GrayScaling of
    '   the passed image via the GrayBlt method implemented in the UserControl.
    '
    'Parameters ************************************************************
    '   hDestDC:     Destination device context
    '   x, y:        Upper-left destination coordinates (pixels)
    '   nWidth:      Width of destination
    '   nHeight:     Height of destination
    '   hSrcImg:     Source StdPicture Object
    '   xSrc, ySrc:  Upper-left source coordinates (pixels)
    '   TransColor:  RGB value for transparent pixels, typically &HC0C0C0.
    '***********************************************************************
    ' Holds original background color
    Dim OrigColor As Long

    ' Holds original background drawing mode
    Dim OrigMode  As Long
    Dim hSrcDC    As Long
    Dim tObj      As Long

    'Handle to the Brush we are using for MaskColor
    Dim hBrush    As Long
    Dim hTmp      As Long

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Create a DC which is compatible with the destination DC
    hSrcDC = CreateCompatibleDC(hDestDC)

    '   Check if it is an Icon or a Bitmap
    If hSrcImg.Type = vbPicTypeBitmap Then
        '   Bitmap, so simply Select it into the DC
        tObj = SelectObject(hSrcDC, hSrcImg.Handle)
        DeleteObject tObj
    Else
        '   This is an Icon, so we need to Draw this into the DC
        '   at the new size....we are using the TransColor here as the
        '   MaskColor so pass the handled to the brush
        hTmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
        tObj = SelectObject(hSrcDC, hTmp)
        hBrush = CreateSolidBrush(transColor)
        'MaskColor)
        DrawIconEx hSrcDC, 0, 0, hSrcImg.Handle, nWidth, nHeight, 0, hBrush, &H1 Or &H2
        '   Clean up the brush
        DeleteObject hBrush
        DeleteObject hTmp
        DeleteObject tObj
    End If

    If (GetDeviceCaps(hDestDC, CAPS1) And C1_TRANSPARENT) Then
        ' Some NT machines support this *super* simple method!
        ' Save original settings, Blt, restore settings.
        OrigMode = SetBkMode(hDestDC, NEWTRANSPARENT)
        OrigColor = SetBkColor(hDestDC, transColor)

        '   Check to see if this is a GreyScale Image, if so then GrayBlt it
        '   to the DC it is located on...
        If Disabled Then
            GrayBlt hSrcDC, hSrcDC, nWidth, nHeight
        End If

        Call BitBlt(hDestDC, X, Y, nWidth, nHeight, hSrcDC, XSrc, YSrc, SRCCOPY)
        Call SetBkColor(hDestDC, OrigColor)
        Call SetBkMode(hDestDC, OrigMode)
    Else

        ' Backup copy of source bitmap
        Dim saveDC       As Long

        ' Mask bitmap (monochrome)
        Dim maskDC       As Long

        ' Inverse of mask bitmap (monochrome)
        Dim invDC        As Long

        ' Combination of source bitmap & background
        Dim resultDC     As Long

        ' Bitmap stores backup copy of source bitmap
        Dim hSaveBmp     As Long

        ' Bitmap stores mask (monochrome)
        Dim hMaskBmp     As Long

        ' Bitmap holds inverse of mask (monochrome)
        Dim hInvBmp      As Long

        ' Bitmap combination of source & background
        Dim hResultBmp   As Long

        ' Holds previous bitmap in saved DC
        Dim hSavePrevBmp As Long

        ' Holds previous bitmap in the mask DC
        Dim hMaskPrevBmp As Long

        ' Holds previous bitmap in inverted mask DC
        Dim hInvPrevBmp  As Long

        ' Holds previous bitmap in destination DC
        Dim hDestPrevBmp As Long

        ' Create DCs to hold various stages of transformation.
        saveDC = CreateCompatibleDC(hDestDC)
        maskDC = CreateCompatibleDC(hDestDC)
        invDC = CreateCompatibleDC(hDestDC)
        resultDC = CreateCompatibleDC(hDestDC)
        ' Create monochrome bitmaps for the mask-related bitmaps.
        hMaskBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
        hInvBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
        ' Create color bitmaps for final result & stored copy of source.
        hResultBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
        hSaveBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
        ' Select bitmaps into DCs.
        hSavePrevBmp = SelectObject(saveDC, hSaveBmp)
        hMaskPrevBmp = SelectObject(maskDC, hMaskBmp)
        hInvPrevBmp = SelectObject(invDC, hInvBmp)
        hDestPrevBmp = SelectObject(resultDC, hResultBmp)
        ' Create mask: set background color of source to transparent color.
        OrigColor = SetBkColor(hSrcDC, transColor)
        Call BitBlt(maskDC, 0, 0, nWidth, nHeight, hSrcDC, XSrc, YSrc, vbSrcCopy)
        transColor = SetBkColor(hSrcDC, OrigColor)
        ' Create inverse of mask to AND w/ source & combine w/ background.
        Call BitBlt(invDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbNotSrcCopy)
        ' Copy background bitmap to result.
        Call BitBlt(resultDC, 0, 0, nWidth, nHeight, hDestDC, X, Y, vbSrcCopy)
        ' AND mask bitmap w/ result DC to punch hole in the background by
        ' painting black area for non-transparent portion of source bitmap.
        Call BitBlt(resultDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbSrcAnd)

        '   Check to see if this is a GreyScale Image, if so then GrayBlt it
        '   to the DC it is located on...
        If Disabled Then
            GrayBlt hSrcDC, hSrcDC, nWidth, nHeight
        End If

        ' get overlapper
        Call BitBlt(saveDC, 0, 0, nWidth, nHeight, hSrcDC, XSrc, YSrc, vbSrcCopy)
        ' AND with inverse monochrome mask
        Call BitBlt(saveDC, 0, 0, nWidth, nHeight, invDC, 0, 0, vbSrcAnd)
        ' XOR these two
        Call BitBlt(resultDC, 0, 0, nWidth, nHeight, saveDC, 0, 0, vbSrcInvert)
        ' Display transparent bitmap on background.
        Call BitBlt(hDestDC, X, Y, nWidth, nHeight, resultDC, 0, 0, vbSrcCopy)
        ' Select original objects back.
        Call SelectObject(saveDC, hSavePrevBmp)
        Call SelectObject(resultDC, hDestPrevBmp)
        Call SelectObject(maskDC, hMaskPrevBmp)
        Call SelectObject(invDC, hInvPrevBmp)
        ' Deallocate system resources.
        Call DeleteObject(hSaveBmp)
        Call DeleteObject(hMaskBmp)
        Call DeleteObject(hInvBmp)
        Call DeleteObject(hResultBmp)
        Call DeleteDC(saveDC)
        Call DeleteDC(invDC)
        Call DeleteDC(maskDC)
        Call DeleteDC(resultDC)
    End If

    Call DeleteDC(hSrcDC)
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.TransBltEx", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function TranslateColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lColor (Long)
'!--------------------------------------------------------------------------------
Private Function TranslateColor(ByVal lColor As Long) As Long

    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    If OleTranslateColor(lColor, 0, TranslateColor) Then
        TranslateColor = -1
    End If

Func_ErrHandlerExit:

    Exit Function

Func_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.TranslateColor", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Func_ErrHandlerExit:

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function UBoundEx
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   uArr() (PanelItem)
'!--------------------------------------------------------------------------------
Private Function UBoundEx(uArr() As PanelItem) As Long

    On Error Resume Next

    UBoundEx = UBound(uArr, 1)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtEdit_KeyUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Select Case KeyCode

        Case vbKeyEscape

            If txtEdit.Visible = True Then
                txtEdit.Visible = False
            End If

        Case vbKeyReturn

            If txtEdit.Visible = True Then
                m_PanelItems(m_ActivePanel).Text = txtEdit.Text
                txtEdit.Visible = False
                Refresh
            End If

    End Select

Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.txtEdit_KeyUp", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtEdit_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtEdit_LostFocus()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If txtEdit.Visible = True Then
        m_PanelItems(m_ActivePanel).Text = txtEdit.Text
        txtEdit.Visible = False
    End If

Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.txtEdit_LostFocus", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Click()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If txtEdit.Visible = True Then
        txtEdit.Visible = False
    End If

    RaiseEvent Click
    RaiseEvent PanelClick(GetPanelIndex())
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_Click", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_DblClick
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_DblClick()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    m_ActivePanel = GetPanelIndex()

    If m_ActivePanel Then

        With m_PanelItems(m_ActivePanel)

            If .Editable Then
                If m_iTheme <> usbClassic Then
                    txtEdit.BackColor = m_BackColor
                    txtEdit.Left = .ItemRect.Left
                    txtEdit.Height = 16
                    txtEdit.Top = ScaleHeight \ 2 - txtEdit.Height \ 2
                    txtEdit.Width = ((.ItemRect.Right - .ItemRect.Left))
                Else
                    txtEdit.BackColor = m_BackColor

                    If Not .Icon Is Nothing Then
                        txtEdit.Left = .ItemRect.Left - 1
                    Else
                        txtEdit.Left = .ItemRect.Left - 4
                    End If

                    txtEdit.Height = 16 - 12
                    txtEdit.Top = (ScaleHeight \ 2 - txtEdit.Height \ 2) - 1
                    txtEdit.Width = ((.ItemRect.Right - .ItemRect.Left)) + 8
                End If

                txtEdit.Text = .Text
                txtEdit.SelStart = 0
                txtEdit.SelLength = Len(.Text)
                txtEdit.Visible = True
                txtEdit.ZOrder 0
                txtEdit.SetFocus
            End If

        End With

    End If

    RaiseEvent DblClick
    RaiseEvent PanelDblClick(m_ActivePanel)
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_DblClick", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Initialize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Initialize()
    m_bIsWinXpOrLater = IsWinXPOrLater
    
    Set m_cSubclass = New cSelfSubHookCallback
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_InitProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    m_BackColor = vbButtonFace
    m_Forecolor = vbButtonText
    Set m_Font = UserControl.Font
    m_GripShape = usbSquare
    m_Sizable = True
    m_Theme = usbAuto
    m_iTheme = m_Theme
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_InitProperties", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_KeyDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent KeyDown(KeyCode, Shift)
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_KeyDown", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_KeyPress
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyAscii (Integer)
'!--------------------------------------------------------------------------------
Private Sub UserControl_KeyPress(KeyAscii As Integer)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent KeyPress(KeyAscii)
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_KeyPress", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_KeyUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent KeyUp(KeyCode, Shift)
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_KeyUp", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_LostFocus()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If txtEdit.Visible = True Then
        m_PanelItems(m_ActivePanel).Text = txtEdit.Text
        txtEdit.Visible = False
    End If

Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_LostFocus", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_MouseDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If m_Sizable Then
        If PtInRect(m_GripRect, CLng(X), CLng(Y)) Then
            '   Relase any events captured previously
            ReleaseCapture
            '   Send a message that we are resizing the form
            SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
        End If
    End If

    RaiseEvent MouseDown(Button, Shift, X, Y)
    RaiseEvent PanelMouseDown(GetPanelIndex(), Button, Shift, X, Y)
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_MouseDown", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_MouseMove
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If PtInRect(m_GripRect, CLng(X), CLng(Y)) Then
        UserControl.MousePointer = vbSizeNWSE
    Else
        UserControl.MousePointer = vbDefault
    End If

    RaiseEvent MouseMove(Button, Shift, X, Y)
    RaiseEvent PanelMouseMove(GetPanelIndex(), Button, Shift, X, Y)
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_MouseMove", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_MouseUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent MouseUp(Button, Shift, X, Y)
    RaiseEvent PanelMouseUp(GetPanelIndex(), Button, Shift, X, Y)
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_MouseUp", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Paint
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Paint()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Refresh
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_Paint", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_ReadProperties
'! Description (Описание)  :   [Read the properties from the property bag - also, a good place to start the subclassing (if we're running)]
'! Parameters  (Переменные):   PropBag (PropertyBag)
'!--------------------------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With PropBag
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        Set UserControl.Font = m_Font
        m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        m_Forecolor = .ReadProperty("ForeColor", Ambient.ForeColor)
        m_GripShape = .ReadProperty("GripShape", usbSquare)
        m_PanelCount = .ReadProperty("PanelCount", 0)
        m_Sizable = .ReadProperty("Sizable", True)
        Theme = .ReadProperty("Theme", usbAuto)
    End With
        
    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_Forecolor
    UserControl.Extender.Align = vbAlignBottom
    m_iTheme = m_Theme

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
                    .ssc_AddMsg UserControl.hWnd, MSG_AFTER, WM_MOUSEMOVE, WM_MOUSELEAVE, WM_NCPAINT, WM_SIZING, WM_THEMECHANGED, WM_SYSCOLORCHANGE
                End If
                
            End With

        End If
    End If

Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_ReadProperties", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Resize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Resize()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With UserControl
        '.Height = 360
        .Height = 700
    End With

    With m_GripRect
        .Left = ScaleWidth - 15
        .Top = ScaleHeight - 15
        .Right = .Left + 15
        .Bottom = .Top + 15
    End With

    UserControl.Refresh
    Refresh
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_Resize", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Show
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Show()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Refresh
Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_Show", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Terminate
'! Description (Описание)  :   [The control is terminating - a good place to stop the subclasser]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Terminate()

    Dim I As Long

    On Error Resume Next

    '   Set the Parents of the Object Back....
    For I = 1 To m_PanelCount

        With m_PanelItems(I)

            If Not .BoundObject Is Nothing Then
                SetParent .BoundObject.hWnd, .BoundParent
            End If

        End With

    Next

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

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With PropBag
        Call .WriteProperty("Font", m_Font, Ambient.Font)
        Call .WriteProperty("BackColor", m_BackColor, Ambient.BackColor)
        Call .WriteProperty("ForeColor", m_Forecolor, Ambient.ForeColor)
        Call .WriteProperty("GripShape", m_GripShape, usbSquare)
        Call .WriteProperty("PanelCount", m_PanelCount, 0)
        Call .WriteProperty("Sizable", m_Sizable, True)
        Call .WriteProperty("Theme", m_Theme, usbAuto)
    End With

Sub_ErrHandlerExit:

    Exit Sub

Sub_ErrHandler:
    Err.Raise Err.Number, "ucStatusBar.UserControl_WriteProperties", Err.Description, Err.HelpFile, Err.HelpContext

    Resume Sub_ErrHandlerExit:

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

            If Not bInCtrl Then
                bInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
                RaiseEvent MouseEnter
            End If

        Case WM_MOUSELEAVE
            bInCtrl = False
            RaiseEvent MouseLeave

        Case WM_NCPAINT, WM_SIZING, WM_SYSCOLORCHANGE, WM_THEMECHANGED
            Refresh

    End Select

End Sub

