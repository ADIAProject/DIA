VERSION 5.00
Begin VB.UserControl ctlFontCombo 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ForeColor       =   &H80000008&
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   ToolboxBitmap   =   "ctlFontCombo.ctx":0000
   Begin VB.Timer TmrAutoText 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   90
   End
   Begin VB.Timer TmrFocus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3330
      Top             =   90
   End
   Begin VB.PictureBox PicPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   3330
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox PicList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   3285
      Begin VB.VScrollBar VScroll1 
         CausesValidation=   0   'False
         Height          =   2595
         Left            =   2760
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   240
      End
      Begin VB.Shape SelBox 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         DrawMode        =   14  'Copy Pen
         FillColor       =   &H8000000D&
         Height          =   285
         Left            =   0
         Top             =   600
         Width           =   2565
      End
   End
   Begin VB.Timer TmrOver 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2820
      Top             =   90
   End
   Begin VB.Shape FocusBox 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H8000000D&
      Height          =   285
      Left            =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   2565
   End
End
Attribute VB_Name = "ctlFontCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within Drivers Installer Assistant and Drivers BackUp Solution.
'This code was originally
'You may download the original version of this code from the following link (good as of June '12):

Option Explicit

Dim mEnabled              As Boolean
Dim mBorderStyle          As CfBdrStyle
Dim mSorted               As Boolean
Dim inRct                 As Boolean
Dim tPos                  As Integer
Dim mButtonBackColor      As Long
Dim mButtonForeColor      As Long
Dim mButtonOverColor      As Long
Dim mButtonBorderStyle    As CfBdrStyle
Dim mShowFocus            As Boolean

Private mListFont()       As String
Private mListCount        As Integer
Private mListPos          As Integer
Private mUsedList()       As String
Private mUsedCount        As Integer
Private mUsedBackColor    As Long
Private mUsedForeColor    As Long
Private mRecent()         As tpRecents
Private mRecentCount      As Integer
Private mRecentMax        As Integer
Private mRecentBackColor  As Long
Private mRecentForeColor  As Long
Private mPreviewText      As String
Private mShowPreview      As Boolean
Private mShowFontName     As Boolean
Private mPreviewSize      As Integer
Private mShowFontInCombo  As Boolean
Private mComboFontCount   As Integer
Private mComboFontSize    As Integer
Private mComboFontBold    As Boolean
Private mComboFontItalic  As Boolean
Private mComboWidth       As Single
Private mForeColor        As Long
Private mBackColor        As Long
Private mComboForeColor   As Long
Private mComboBackColor   As Long
Private mComboSelectColor As Long
Private mUseMouseWheel    As Boolean
Private mAutoText         As String
Private CloseMe           As Boolean
Private doNothing         As Boolean
Private fList()           As tpRecents
Private fPos              As Integer
Private bCancel           As Boolean
Private Resultat          As Long
Private Ident             As Long
Private Donnee            As String
Private TailleBuffer      As Long
Private mXPStyle          As Boolean

Private Type POINTAPI
    X                     As Long
    Y                     As Long
End Type

Private Type RECT
    Left                  As Long
    Top                   As Long
    Right                 As Long
    Bottom                As Long
End Type

Private Btn               As RECT
Private uRct              As RECT
Private MouseCoords       As POINTAPI

Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (lpMsg As TMSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32.dll" (lpMsg As TMSG) As Long
Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (lpMsg As TMSG) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, qRC As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRECT As RECT, pClipRect As RECT) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetFocus Lib "user32.dll" () As Long

Private Const HWND_TOP           As Long = 0
Private Const WM_MOUSEWHEEL      As Long = &H20A
Private Const GWL_EXSTYLE        As Long = (-20)
Private Const WS_EX_TOOLWINDOW   As Long = &H80

' --Formatting Text Consts
Private Const DT_LEFT           As Long = &H0
Private Const DT_CENTER         As Long = &H1
Private Const DT_RIGHT          As Long = &H2
Private Const DT_NOCLIP         As Long = &H100
Private Const DT_WORDBREAK      As Long = &H10
Private Const DT_CALCRECT       As Long = &H400
Private Const DT_RTLREADING     As Long = &H20000
Private Const DT_DRAWFLAG       As Long = DT_CENTER Or DT_WORDBREAK
Private Const DT_TOP            As Long = &H0
Private Const DT_BOTTOM         As Long = &H8
Private Const DT_VCENTER        As Long = &H4
Private Const DT_SINGLELINE     As Long = &H20
Private Const DT_WORD_ELLIPSIS  As Long = &H40000

Private Const SWP_REFRESH        As Long = (&H1 Or &H2 Or &H4 Or &H20)
Private Const SWP_NOACTIVATE     As Long = &H10
Private Const SWP_NOMOVE         As Long = &H2
Private Const SWP_NOSIZE         As Long = &H1
Private Const SWP_SHOWWINDOW     As Long = &H40
Private Const SWP_NOOWNERZORDER  As Long = &H200
Private Const SWP_NOZORDER       As Long = &H4
Private Const SWP_FRAMECHANGED   As Long = &H20

Public Enum CfBdrStyle
    sNone = 0
    sRaised = &H1 Or &H4
    sSunken = &H2 Or &H8
    sBump = &H1 Or &H8
    sEtched = &H2 Or &H4
    sSmoothRaised = &H4
    sSmoothSunken = &H2
End Enum

Public Enum CfEdgeStyle
    edgeAll = &HF
    edgeLeft = &H2
    edgeTop = &H4
    edgeRight = &H1
    edgeBottom = &H8
End Enum

Public Enum HkeyLoc2
    'HKEY_CLASSES_ROOT = &H80000000
    'HKEY_CURRENT_USER = &H80000001
    'HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_DYN_DATA = &H80000004
End Enum

Private Enum eBtnState
    bUp = 0
    bOver = 1
    bDown = 2
End Enum

Private Enum sTxtPosition
    TopLeft = 0
    TopCenter = 1
    TopRight = 2
    MiddleLeft = 3
    MiddleCenter = 4
    MiddleRight = 5
    BottomLeft = 6
    BottomCenter = 7
    BottomRight = 8
End Enum

Private Enum HkeyLoc
    'HKEY_CLASSES_ROOT = &H80000000
    'HKEY_CURRENT_USER = &H80000001
    'HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_DYN_DATA = &H80000004
End Enum

Private Type tpRecents
    fName                               As String
    fIndex                              As String
    fRecent                             As Boolean
End Type

Private Type TMSG
    hWnd                                As Long
    nMsg                                As Long
    wParam                              As Long
    lParam                              As Long
    Time                                As Long
    PT                                  As POINTAPI
End Type

Private Msg As TMSG

Public Event SelectedFontChanged(NewFontName As String)
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event FontNotFound(FontName As String)

Private OldDr   As eBtnState
Private OldFont As String

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Initialize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Initialize()
    CloseMe = False
    SetWindowLong PicList.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
    SetParent PicList.hWnd, 0
    SetWindowLong PicPreview.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
    SetParent PicPreview.hWnd, 0
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    mBackColor = vNewValue
    UserControl.BackColor = mBackColor
    DrawControl , True
    PropertyChanged "BackColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BorderStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get BorderStyle() As CfBdrStyle
    BorderStyle = mBorderStyle
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BorderStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (CfBdrStyle)
'!--------------------------------------------------------------------------------
Public Property Let BorderStyle(ByVal vNewValue As CfBdrStyle)
    mBorderStyle = vNewValue
    UserControl_Resize
    PropertyChanged "BorderStyle"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ButtonBackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ButtonBackColor() As OLE_COLOR
    ButtonBackColor = mButtonBackColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ButtonBackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ButtonBackColor(ByVal vNewValue As OLE_COLOR)
    mButtonBackColor = vNewValue
    DrawControl , True
    PropertyChanged "ButtonBackColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ButtonBorderStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ButtonBorderStyle() As CfBdrStyle
    ButtonBorderStyle = mButtonBorderStyle
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ButtonBorderStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (CfBdrStyle)
'!--------------------------------------------------------------------------------
Public Property Let ButtonBorderStyle(ByVal vNewValue As CfBdrStyle)
    mButtonBorderStyle = vNewValue
    DrawControl , True
    PropertyChanged "ButtonBorderStyle"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ButtonForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ButtonForeColor() As OLE_COLOR
    ButtonForeColor = mButtonForeColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ButtonForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ButtonForeColor(ByVal vNewValue As OLE_COLOR)
    mButtonForeColor = vNewValue
    DrawControl , True
    PropertyChanged "ButtonForeColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ButtonOverColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ButtonOverColor() As OLE_COLOR
    ButtonOverColor = mButtonOverColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ButtonOverColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ButtonOverColor(ByVal vNewValue As OLE_COLOR)
    mButtonOverColor = vNewValue
    PropertyChanged "ButtonOverColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboBackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ComboBackColor() As OLE_COLOR
    ComboBackColor = mComboBackColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboBackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ComboBackColor(ByVal vNewValue As OLE_COLOR)
    mComboBackColor = vNewValue
    PropertyChanged "ComboBackColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboFontBold
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ComboFontBold() As Boolean
Attribute ComboFontBold.VB_MemberFlags = "400"
    ComboFontBold = mComboFontBold
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboFontBold
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let ComboFontBold(ByVal vNewValue As Boolean)
    mComboFontBold = vNewValue
    PropertyChanged "ComboFontBold"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboFontCount
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ComboFontCount() As Integer
    ComboFontCount = mComboFontCount
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboFontCount
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Integer)
'!--------------------------------------------------------------------------------
Public Property Let ComboFontCount(ByVal vNewValue As Integer)

    If vNewValue > 50 Or vNewValue < 5 Then vNewValue = 20
    mComboFontCount = vNewValue
    PropertyChanged "ComboFontCount"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboFontItalic
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ComboFontItalic() As Boolean
Attribute ComboFontItalic.VB_MemberFlags = "400"
    ComboFontItalic = mComboFontItalic
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboFontItalic
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let ComboFontItalic(ByVal vNewValue As Boolean)
    mComboFontItalic = vNewValue
    PropertyChanged "ComboFontItalic"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboFontSize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ComboFontSize() As Integer
    ComboFontSize = mComboFontSize
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboFontSize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Integer)
'!--------------------------------------------------------------------------------
Public Property Let ComboFontSize(ByVal vNewValue As Integer)

    If vNewValue > 50 Or vNewValue < 6 Then vNewValue = 8
    mComboFontSize = vNewValue
    PropertyChanged "ComboFontSize"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ComboForeColor() As OLE_COLOR
    ComboForeColor = mComboForeColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ComboForeColor(ByVal vNewValue As OLE_COLOR)
    mComboForeColor = vNewValue
    PropertyChanged "ComboForeColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboSelectColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ComboSelectColor() As OLE_COLOR
    ComboSelectColor = mComboSelectColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboSelectColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ComboSelectColor(ByVal vNewValue As OLE_COLOR)
    mComboSelectColor = vNewValue
    PropertyChanged "ComboSelectColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboWidth
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ComboWidth() As Single
    ComboWidth = mComboWidth
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ComboWidth
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Single)
'!--------------------------------------------------------------------------------
Public Property Let ComboWidth(ByVal vNewValue As Single)
    mComboWidth = vNewValue
    PropertyChanged "ComboWidth"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Enabled
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Enabled
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let Enabled(ByVal vNewValue As Boolean)
    mEnabled = vNewValue
    DrawControl , True
    PropertyChanged "Enabled"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Font
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Font
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (StdFont)
'!--------------------------------------------------------------------------------
Public Property Set Font(ByVal vNewValue As StdFont)
    Set UserControl.Font = vNewValue
    UserControl_Resize
    PropertyChanged "Font"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    mForeColor = vNewValue
    UserControl.ForeColor = mForeColor
    DrawControl , True
    PropertyChanged "ForeColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property hWnd
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ListCount
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ListCount() As Integer
Attribute ListCount.VB_MemberFlags = "400"
    ListCount = mListCount
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ListFont
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Public Property Get ListFont(Index As Integer) As String
    ListFont = mListFont(Index)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ListIndex
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = mListPos
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ListIndex
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Integer)
'!--------------------------------------------------------------------------------
Public Property Let ListIndex(ByVal vNewValue As Integer)
    mListPos = vNewValue
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PreviewSize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get PreviewSize() As Integer
    PreviewSize = mPreviewSize
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PreviewSize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Integer)
'!--------------------------------------------------------------------------------
Public Property Let PreviewSize(ByVal vNewValue As Integer)

    If vNewValue > 10 Then
        If vNewValue < 200 Then
            mPreviewSize = vNewValue
            PropertyChanged "PreviewSize"
        End If
    End If

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PreviewText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get PreviewText() As String
    PreviewText = mPreviewText
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PreviewText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (String)
'!--------------------------------------------------------------------------------
Public Property Let PreviewText(ByVal vNewValue As String)
    mPreviewText = vNewValue
    PropertyChanged "PreviewText"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property RecentBackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get RecentBackColor() As OLE_COLOR
    RecentBackColor = mRecentBackColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property RecentBackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let RecentBackColor(ByVal vNewValue As OLE_COLOR)
    mRecentBackColor = vNewValue
    PropertyChanged "RecentBackColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property RecentForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get RecentForeColor() As OLE_COLOR
    RecentForeColor = mRecentForeColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property RecentForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let RecentForeColor(ByVal vNewValue As OLE_COLOR)
    mRecentForeColor = vNewValue
    PropertyChanged "RecentForeColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property RecentMax
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get RecentMax() As Integer
Attribute RecentMax.VB_Description = "If you don't want to use Recents feature enter 0"
    RecentMax = mRecentMax
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property RecentMax
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Integer)
'!--------------------------------------------------------------------------------
Public Property Let RecentMax(ByVal vNewValue As Integer)
    mRecentMax = vNewValue
    PropertyChanged "RecentMax"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property RunMode
'! Description (Описание)  :   [Ambient.UserMode tells us whether the UC's container is in design mode or user mode/run-time.
'                               Unfortunately, this isn't supported in all containers.]
'                               http://www.vbforums.com/showthread.php?805711-VB6-UserControl-Ambient-UserMode-workaround&s=8dd326860cbc22bed07bd13f6959ca70
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get RunMode() As Boolean
    RunMode = True
    On Error Resume Next
    RunMode = Ambient.UserMode
    RunMode = Extender.Parent.RunMode
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property SelectedFont
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get SelectedFont() As String
Attribute SelectedFont.VB_MemberFlags = "400"
    SelectedFont = mListFont(mListPos)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property SelectedFont
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (String)
'!--------------------------------------------------------------------------------
Public Property Let SelectedFont(ByVal vNewValue As String)

    Dim ii As Integer

    ii = FontExist(vNewValue)

    If ii > -1 Then
        mListPos = ii
        RaiseEvent SelectedFontChanged(mListFont(mListPos))
        DrawControl , True
    Else
        RaiseEvent FontNotFound(vNewValue)
    End If

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ShowFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ShowFocus() As Boolean
    ShowFocus = mShowFocus
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ShowFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let ShowFocus(ByVal vNewValue As Boolean)
    mShowFocus = vNewValue
    PropertyChanged "ShowFocus"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ShowFontInCombo
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ShowFontInCombo() As Boolean
    ShowFontInCombo = mShowFontInCombo
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ShowFontInCombo
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let ShowFontInCombo(ByVal vNewValue As Boolean)
    mShowFontInCombo = vNewValue
    PropertyChanged "ShowFontInCombo"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ShowFontName
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ShowFontName() As Boolean
    ShowFontName = mShowFontName
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ShowFontName
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let ShowFontName(ByVal vNewValue As Boolean)
    mShowFontName = vNewValue
    PropertyChanged "ShowFontName"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ShowPreview
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ShowPreview() As Boolean
    ShowPreview = mShowPreview
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ShowPreview
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let ShowPreview(ByVal vNewValue As Boolean)
    mShowPreview = vNewValue
    PropertyChanged "ShowPreview"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Sorted
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Sorted() As Boolean
    Sorted = mSorted
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Sorted
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let Sorted(ByVal vNewValue As Boolean)

    Dim ii As Integer
    Dim fI As Integer

    mSorted = vNewValue

    If RunMode Then
        FillList

        If mSorted = True Then SortList

        For ii = 0 To mRecentCount - 1
            fI = FontExist(mRecent(ii).fName)
            mRecent(ii).fIndex = fI
        Next

    End If

    DrawControl , True
    PropertyChanged "Sorted"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UsedBackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get UsedBackColor() As OLE_COLOR
    UsedBackColor = mUsedBackColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UsedBackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let UsedBackColor(ByVal vNewValue As OLE_COLOR)
    mUsedBackColor = vNewValue
    PropertyChanged "UsedBackColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UsedCount
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get UsedCount() As Integer
    UsedCount = mUsedCount
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UsedForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get UsedForeColor() As OLE_COLOR
    UsedForeColor = mUsedForeColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UsedForeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let UsedForeColor(ByVal vNewValue As OLE_COLOR)
    mUsedForeColor = vNewValue
    PropertyChanged "UsedForeColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseMouseWheel
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get UseMouseWheel() As Boolean
    UseMouseWheel = mUseMouseWheel
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseMouseWheel
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let UseMouseWheel(ByVal vNewValue As Boolean)
    mUseMouseWheel = vNewValue
    PropertyChanged "UseMouseWheel"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property XpStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get XpStyle() As Boolean
    XpStyle = mXPStyle
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property XpStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vNewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let XpStyle(ByVal vNewValue As Boolean)
    mXPStyle = vNewValue
    UserControl_Resize
    PropertyChanged "XPStyle"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function AddToUsedList
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   FontName (String)
'!--------------------------------------------------------------------------------
Public Function AddToUsedList(FontName As String) As Integer

    Dim ii As Integer
    Dim F  As Boolean

    For ii = 0 To mUsedCount - 1

        If LCase$(mUsedList(ii)) = LCase$(FontName) Then
            F = True

            Exit For

        End If

    Next

    If F = False Then
        mUsedCount = mUsedCount + 1

        ReDim Preserve mUsedList(mUsedCount)

        mUsedList(mUsedCount - 1) = FontName
        AddToUsedList = mUsedCount - 1
    Else
        AddToUsedList = -1
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ClearRecent
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub ClearRecent()
    mRecentCount = 0

    ReDim mRecent(0)

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ClearUsedList
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub ClearUsedList()
    mUsedCount = 0

    ReDim mUsedList(0)

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawArw
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   ArrowColor (Long = -1)
'!--------------------------------------------------------------------------------
Private Sub DrawArw(Optional ArrowColor As Long = -1)

    Dim ColUp As Long
    Dim tCol  As Long

    If ArrowColor = -1 Then
        tCol = mButtonForeColor
    Else
        tCol = ArrowColor
    End If

    If mEnabled = False Then
        OleTranslateColor vbGrayText, 0, ColUp
    Else
        OleTranslateColor tCol, 0, ColUp
    End If

    SetPixel UserControl.hDC, Btn.Left - 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left - 2 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 2 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left - 3 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 3 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left - 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left - 2 + (Btn.Right - Btn.Left) \ 2, Btn.Top + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 2 + (Btn.Right - Btn.Left) \ 2, Btn.Top + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + (Btn.Right - Btn.Left) \ 2, Btn.Top + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left - 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top + 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top + 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + (Btn.Right - Btn.Left) \ 2, Btn.Top + 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + (Btn.Right - Btn.Left) \ 2, Btn.Top + 2 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    Refresh
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawControl
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   eDraw (eBtnState = bUp)
'                              DrawAll (Boolean = False)
'!--------------------------------------------------------------------------------
Private Sub DrawControl(Optional eDraw As eBtnState = bUp, Optional DrawAll As Boolean = False)

    Dim Br       As Long
    Dim tC       As Long
    Dim tCol     As Long

    UserControl.Enabled = mEnabled
    mXPStyle = mXPStyle And DrawTheme("Button", 1, 1, Btn)

    If mXPStyle = False Then
        OleTranslateColor mButtonBackColor, 0, tC
        Br = CreateSolidBrush(tC)

        If mEnabled = False Then
            Cls
            FillRect UserControl.hDC, Btn, Br
            DrawEdge UserControl.hDC, uRct, mBorderStyle, edgeAll
            DrawEdge UserControl.hDC, Btn, mButtonBorderStyle, edgeAll
            tCol = UserControl.ForeColor
            UserControl.ForeColor = &H80000011
            UserControl.CurrentY = ((ScaleHeight - TextHeight("X")) / 2) - 1
            UserControl.CurrentX = 4

            If RunMode = True And mListCount Then
                UserControl.Print mListFont(mListPos)
            Else
                UserControl.Print Ambient.DisplayName
            End If

            UserControl.ForeColor = tCol
            DrawArw
            DeleteObject Br

            Exit Sub

        End If

        If OldDr = eDraw Then
            If DrawAll = False Then

                Exit Sub

            End If
        End If

        UserControl.Cls
        UserControl.CurrentY = ((ScaleHeight - TextHeight("X")) / 2) - 1
        UserControl.CurrentX = 4

        If RunMode = True And mListCount Then
            UserControl.Print mListFont(mListPos)
        Else
            UserControl.Print Ambient.DisplayName
        End If

        Select Case eDraw

            Case bUp
                DrawEdge UserControl.hDC, uRct, mBorderStyle, edgeAll
                FillRect UserControl.hDC, Btn, Br
                DrawEdge UserControl.hDC, Btn, ButtonBorderStyle, edgeAll

            Case bOver
                DrawEdge UserControl.hDC, uRct, mBorderStyle, edgeAll
                FillRect UserControl.hDC, Btn, Br
                DrawEdge UserControl.hDC, Btn, ButtonBorderStyle, edgeAll

            Case bDown
                DrawEdge UserControl.hDC, uRct, mBorderStyle, edgeAll
                FillRect UserControl.hDC, Btn, Br
                DrawEdge UserControl.hDC, Btn, InvBdr(ButtonBorderStyle), edgeAll
        End Select

        DeleteObject Br

        If eDraw = bOver Then
            DrawArw mButtonOverColor
        Else
            DrawArw
        End If

    Else
        UserControl.Cls

        If mEnabled = True Then

            Select Case eDraw

                Case bUp
                    DrawTheme "ComboBox", 2, 1, uRct
                    DrawTheme "ComboBox", 1, 1, Btn

                Case bOver
                    DrawTheme "ComboBox", 2, 2, uRct
                    DrawTheme "ComboBox", 1, 2, Btn

                Case bDown
                    DrawTheme "ComboBox", 2, 3, uRct
                    DrawTheme "ComboBox", 1, 3, Btn
            End Select

        Else
            DrawTheme "ComboBox", 2, 4, uRct
            DrawTheme "ComboBox", 1, 4, Btn
            tCol = UserControl.ForeColor
            UserControl.ForeColor = &H80000011
            UserControl.CurrentY = ((ScaleHeight - TextHeight("X")) / 2) - 1
            UserControl.CurrentX = 4
            UserControl.Print mListFont(mListPos)
            UserControl.ForeColor = tCol

            Exit Sub

        End If

        UserControl.CurrentY = ((ScaleHeight - TextHeight("X")) / 2) - 1
        UserControl.CurrentX = 4

        If RunMode = True And mListCount Then
            UserControl.Print mListFont(mListPos)
        Else
            UserControl.Print Ambient.DisplayName
        End If
    End If

    OldDr = eDraw
    Refresh
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawList
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub DrawList()
    On Local Error Resume Next

    Dim ii  As Integer
    Dim Br  As Long
    Dim tC  As Long
    Dim rct As RECT

    OleTranslateColor mRecentBackColor, 0, tC
    Br = CreateSolidBrush(tC)
    PicList.Cls
    doNothing = True
    VScroll1.Max = mListCount - mComboFontCount + mRecentCount
    VScroll1.LargeChange = ((mListCount + mRecentCount) \ mComboFontCount) + 1
    SetList
    SetRect rct, 0, 0, PicList.ScaleWidth, mRecentCount * (mComboFontSize * 2)
    FillRect PicList.hDC, rct, Br
    DeleteObject Br
    OleTranslateColor mUsedBackColor, 0, tC
    Br = CreateSolidBrush(tC)
    PicList.Line (0, mRecentCount * (mComboFontSize * 2))-(PicList.ScaleWidth, mRecentCount * (mComboFontSize * 2))

    For ii = 0 To mRecentCount - 1
        PicList.CurrentX = 2
        PicList.CurrentY = (ii * (mComboFontSize * 2)) + 2

        If mShowFontInCombo = True Then PicList.FontName = mRecent(ii).fName
        PicList.FontSize = mComboFontSize
        PicList.FontItalic = mComboFontItalic
        PicList.FontBold = mComboFontBold

        If IsUsed(mRecent(ii).fName) = False Then
            PicList.ForeColor = mRecentForeColor
        Else
            SetRect rct, 0, ii * (mComboFontSize * 2), PicList.ScaleWidth, (ii + 1) * (mComboFontSize * 2)
            FillRect PicList.hDC, rct, Br
            PicList.ForeColor = mUsedForeColor
        End If

        PicList.Print mRecent(ii).fName
    Next

    For ii = 0 To mComboFontCount - 1

        If IsUsed(fList(ii).fName) = False Then
            PicList.ForeColor = mComboForeColor
        Else
            SetRect rct, 0, (ii * (mComboFontSize * 2)) + ((mComboFontSize * 2) * mRecentCount) + 2, PicList.ScaleWidth, ((ii + 1) * (mComboFontSize * 2)) + ((mComboFontSize * 2) * mRecentCount)
            FillRect PicList.hDC, rct, Br
            PicList.ForeColor = mUsedForeColor
        End If

        PicList.CurrentX = 2
        PicList.CurrentY = (ii * (mComboFontSize * 2)) + 2 + ((mComboFontSize * 2) * mRecentCount)

        If mShowFontInCombo = True Then PicList.FontName = fList(ii).fName
        PicList.FontSize = mComboFontSize
        PicList.FontItalic = mComboFontItalic
        PicList.FontBold = mComboFontBold
        PicList.Print fList(ii).fName
    Next

    DeleteObject Br
    SelBox.Move 0, (fPos - VScroll1.Value + mRecentCount) * (mComboFontSize * 2), PicList.ScaleWidth + 2, (mComboFontSize * 2) + 2
    doNothing = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function DrawTheme
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sClass (String)
'                              iPart (Long)
'                              iState (Long)
'                              rtRect (RECT)
'!--------------------------------------------------------------------------------
Private Function DrawTheme(sClass As String, ByVal iPart As Long, ByVal iState As Long, rtRect As RECT) As Boolean

    Dim hTheme  As Long
    Dim lResult As Long

    On Error GoTo NoXP

    hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))

    If (hTheme) Then
        lResult = DrawThemeBackground(hTheme, UserControl.hDC, iPart, iState, rtRect, rtRect)
        DrawTheme = IIf(lResult, False, True)
    Else
        DrawTheme = False
    End If

    Call CloseThemeData(hTheme)

    Exit Function

NoXP:
    DrawTheme = False
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawTxt
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   ObjhDC (Long)
'                              oText (String)
'                              TxtRect (RECT)
'                              mPosition (sTxtPosition)
'                              MultiLine (Boolean = False)
'                              WordWrap (Boolean = False)
'                              WordEllipsis (Boolean = False)
'!--------------------------------------------------------------------------------
Private Sub DrawTxt(ObjhDC As Long, oText As String, TxtRect As RECT, mPosition As sTxtPosition, Optional MultiLine As Boolean = False, Optional WordWrap As Boolean = False, Optional WordEllipsis As Boolean = False)

    Dim tFormat As Long

    Select Case mPosition

        Case TopLeft
            tFormat = DT_TOP + DT_LEFT

        Case TopCenter
            tFormat = DT_TOP + DT_CENTER

        Case TopRight
            tFormat = DT_TOP + DT_RIGHT

        Case MiddleLeft
            tFormat = DT_VCENTER + DT_LEFT

        Case MiddleCenter
            tFormat = DT_VCENTER + DT_CENTER

        Case MiddleRight
            tFormat = DT_VCENTER + DT_RIGHT

        Case BottomLeft
            tFormat = DT_BOTTOM + DT_LEFT

        Case BottomCenter
            tFormat = DT_BOTTOM + DT_CENTER

        Case BottomRight
            tFormat = DT_BOTTOM + DT_RIGHT
    End Select

    If MultiLine = False Then
        tFormat = tFormat + DT_SINGLELINE
    End If

    If WordWrap = True And MultiLine = True Then
        tFormat = tFormat + DT_WORDBREAK
    End If

    If WordEllipsis = True Then
        tFormat = tFormat + DT_WORD_ELLIPSIS
    End If

    tFormat = tFormat + DT_NOCLIP
    DrawTextW ObjhDC, StrPtr(oText & vbNullChar), -1, TxtRect, tFormat
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FillList
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub FillList()

    Dim ii As Integer

    mListCount = Screen.FontCount - 1

    ReDim mListFont(mListCount)

    For ii = 0 To Screen.FontCount - 1
        mListFont(ii) = Screen.Fonts(ii)
    Next

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FontExist
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Font2Find (String)
'                              StartPos (Integer = 0)
'!--------------------------------------------------------------------------------
Public Function FontExist(Font2Find As String, Optional StartPos As Integer = 0) As Integer

    Dim ii As Integer

    FontExist = -1

    For ii = StartPos To mListCount

        If LCase$(mListFont(ii)) Like LCase$(Font2Find) Then
            FontExist = ii

            Exit For

        End If

    Next

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function InvBdr
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Bdr (CfBdrStyle)
'!--------------------------------------------------------------------------------
Private Function InvBdr(Bdr As CfBdrStyle) As CfBdrStyle

    Select Case Bdr

        Case sNone
            InvBdr = sNone

        Case sRaised
            InvBdr = sSunken

        Case sSunken
            InvBdr = sRaised

        Case sBump
            InvBdr = sEtched

        Case sEtched
            InvBdr = sBump

        Case sSmoothRaised
            InvBdr = sSmoothSunken

        Case sSmoothSunken
            InvBdr = sSmoothRaised
    End Select

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsUsed
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   FontName (String)
'!--------------------------------------------------------------------------------
Private Function IsUsed(FontName As String) As Boolean

    Dim ii As Integer
    Dim F  As Boolean

    For ii = 0 To mUsedCount - 1

        If LCase$(mUsedList(ii)) = LCase$(FontName) Then
            F = True

            Exit For

        End If

    Next

    IsUsed = F
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadRecentFonts
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   MyHkey (HkeyLoc2)
'                              MyGroup (String)
'                              MySection (String)
'                              myKey (String)
'!--------------------------------------------------------------------------------
Public Sub LoadRecentFonts(MyHkey As HkeyLoc2, MyGroup As String, MySection As String, myKey As String)

    Dim ii As Integer
    Dim fN As String
    Dim fI As Integer

    ReDim mRecent(mRecentMax)

    For ii = 0 To mRecentMax - 1
        fN = ReadValue(MyHkey, MyGroup & "\" & MySection & "\" & myKey, "RecentFontName" & ii + 1, "")
        fI = FontExist(fN)

        If fI > -1 Then
            mRecent(ii).fName = fN
            mRecent(ii).fIndex = fI
        End If

    Next

    SetRecents
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mgSort
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   pStart (Long)
'                              pEnd (Long)
'!--------------------------------------------------------------------------------
Private Sub mgSort(ByVal pStart As Long, ByVal pEnd As Long)

    Dim m     As Long
    Dim n     As Long
    Dim tStr1 As String

    m = pStart
    n = pEnd
    tStr1 = LCase$(mListFont((pStart + pEnd) \ 2))

    Do
        Do While LCase$(mListFont(m)) < tStr1
            m = m + 1
        Loop

        Do While LCase$(mListFont(n)) > tStr1
            n = n - 1
        Loop

        If m <= n Then
            SwapStrings mListFont(m), mListFont(n)
            m = m + 1
            n = n - 1
        End If

    Loop Until m > n

    If pStart < n Then Call mgSort(pStart, n)
    If m < pEnd Then Call mgSort(m, pEnd)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PicList_KeyDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub PicList_KeyDown(KeyCode As Integer, Shift As Integer)
    UserControl_KeyDown KeyCode, Shift
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PicList_KeyPress
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyAscii (Integer)
'!--------------------------------------------------------------------------------
Private Sub PicList_KeyPress(KeyAscii As Integer)
    UserControl_KeyPress KeyAscii
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PicList_KeyUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub PicList_KeyUp(KeyCode As Integer, Shift As Integer)
    UserControl_KeyUp KeyCode, Shift
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PicList_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub PicList_LostFocus()
    PicList.Visible = False
    PicPreview.Visible = False
    TmrFocus.Enabled = False
    CloseMe = True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PicList_MouseDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub PicList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim TI As Integer

    TI = Int(Y \ (mComboFontSize * 2))

    If TI < mRecentCount Then
        mListPos = mRecent(TI).fIndex
    Else
        mListPos = fList(TI - mRecentCount).fIndex
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PicList_MouseMove
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub PicList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next

    Dim tFont As String
    Dim TI    As Integer

    TI = Int(Y \ (mComboFontSize * 2))
    fPos = TI

    If TmrAutoText.Enabled = False Then
        SelBox.Move 0, CLng(Y \ (mComboFontSize * 2)) * (mComboFontSize * 2), PicList.ScaleWidth + 2, (mComboFontSize * 2) + 2

        If TI < mRecentCount Then
            tFont = mRecent(TI).fName
        Else
            tFont = fList(TI - mRecentCount).fName
        End If

        ShowFont tFont
        DoEvents
    End If

    If TmrAutoText.Enabled = True Then

        Exit Sub

    End If

    Do
        GetCursorPos MouseCoords

        If WindowFromPoint(MouseCoords.X, MouseCoords.Y) = PicList.hWnd Then
            If mUseMouseWheel = True Then
                GetMessage Msg, Parent.hWnd, 0, 0
                DispatchMessage Msg
                TranslateMessage Msg
                DoEvents

                With Msg

                    If .nMsg = WM_MOUSEWHEEL Then
                        If VScroll1.Value < VScroll1.Max And Sgn(.wParam) < 0 Then
                            If VScroll1.Value + 3 > VScroll1.Max Then
                                VScroll1.Value = VScroll1.Max
                            Else
                                VScroll1.Value = VScroll1.Value + 3
                            End If

                        Else

                            If VScroll1.Value - 3 < 0 Then
                                VScroll1.Value = 0
                            Else
                                VScroll1.Value = VScroll1.Value - 3
                            End If
                        End If
                    End If

                End With

            End If

        ElseIf CloseMe = False Then

            If WindowFromPoint(MouseCoords.X, MouseCoords.Y) = UserControl.hWnd Then Exit Do
            GetMessage Msg, Parent.hWnd, 0, 0
            DispatchMessage Msg
            TranslateMessage Msg
            DoEvents

            If Msg.nMsg = 513 Then
                CloseMe = True

                Exit Do

            End If

        Else

            Exit Do

        End If

        DoEvents
    Loop

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PicList_MouseUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub PicList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetRecents mListFont(mListPos), mListPos
    PicList.Visible = False
    PicPreview.Visible = False
    TmrFocus.Enabled = False
    DrawControl , True
    CloseMe = True
    RaiseEvent SelectedFontChanged(mListFont(mListPos))
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PicPreview_KeyDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub PicPreview_KeyDown(KeyCode As Integer, Shift As Integer)
    UserControl_KeyDown KeyCode, Shift
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PicPreview_KeyPress
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyAscii (Integer)
'!--------------------------------------------------------------------------------
Private Sub PicPreview_KeyPress(KeyAscii As Integer)
    UserControl_KeyPress KeyAscii
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PicPreview_KeyUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub PicPreview_KeyUp(KeyCode As Integer, Shift As Integer)
    UserControl_KeyUp KeyCode, Shift
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ReadValue
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   MyHkey (HkeyLoc)
'                              myKey (String)
'                              MyValue (String)
'                              MyDefaultData (String = vbnullstring)
'!--------------------------------------------------------------------------------
Private Function ReadValue(MyHkey As HkeyLoc, myKey As String, MyValue As String, Optional ByVal MyDefaultData As String = vbNullString) As String

    On Error GoTo ReadValue_Error

    Resultat = 0
    Ident = 0
    TailleBuffer = 0
    Resultat = RegCreateKey(MyHkey, myKey, Ident)

    If Resultat <> 0 Then

        Exit Function

    End If

    Resultat = RegQueryValueEx(Ident, MyValue, 0&, 1, 0&, TailleBuffer)

    If TailleBuffer < 2 Then
        ReadValue = MyDefaultData

        Exit Function

    End If

    Donnee = Space$(TailleBuffer + 1)
    Resultat = RegQueryValueEx(Ident, MyValue, 0&, 1, ByVal Donnee, TailleBuffer)
    Donnee = Left$(Donnee, TailleBuffer - 1)
    ReadValue = Donnee

    On Error GoTo 0

ReadValue_Error:

    Exit Function

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub RemoveFromUsedList
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   FontName (String)
'!--------------------------------------------------------------------------------
Public Sub RemoveFromUsedList(FontName As String)

    Dim ii    As Integer
    Dim tUL() As String
    Dim fQ    As Integer

    ReDim tUL(mUsedCount)

    fQ = 1

    For ii = 0 To mUsedCount - 1

        If LCase$(mUsedList(ii)) <> LCase$(FontName) Then
            tUL(fQ - 1) = mUsedList(ii)
            fQ = fQ + 1
        End If

    Next

    mUsedList = tUL
    mUsedCount = fQ

    ReDim Preserve mUsedList(mUsedCount)

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveRecentFonts
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   MyHkey (HkeyLoc2)
'                              MyGroup (String)
'                              MySection (String)
'                              myKey (String)
'!--------------------------------------------------------------------------------
Public Sub SaveRecentFonts(MyHkey As HkeyLoc2, MyGroup As String, MySection As String, myKey As String)

    Dim ii As Integer

    For ii = 0 To mRecentCount - 1
        SetValue MyHkey, MyGroup & "\" & MySection & "\" & myKey, "RecentFontName" & ii + 1, mRecent(ii).fName
    Next

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetList
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SetList()

    Dim ii    As Integer
    Dim RecQ  As Integer
    Dim Start As Integer

    ReDim fList(mComboFontCount)

    Start = fPos

    If Start + mComboFontCount > mListCount Then
        Start = mListCount - mComboFontCount
    End If

    VScroll1.Value = Start

    For ii = Start To Start + mComboFontCount - RecQ
        fList(RecQ).fName = mListFont(ii)
        fList(RecQ).fIndex = ii
        fList(RecQ).fRecent = False
        RecQ = RecQ + 1
    Next

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetRecents
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   CurRecent (String)
'                              CurIndex (Integer)
'!--------------------------------------------------------------------------------
Private Sub SetRecents(Optional CurRecent As String, Optional CurIndex As Integer)

    Dim m         As Integer
    Dim n         As Integer
    Dim TmpLast() As tpRecents
    Dim a%, b%
    Dim myLast    As tpRecents

    For n = 0 To mRecentMax - 1

        If mRecent(0).fName = CurRecent Then
            If n <> 0 Then
                myLast = mRecent(0)
                mRecent(0) = mRecent(n)
                mRecent(n) = myLast
            End If

            Exit For

        End If

    Next

    ReDim TmpLast(mRecentMax)

    If CurRecent = vbNullString Then
        TmpLast = mRecent
    Else

        For n = 1 To mRecentMax
            myLast = mRecent(n - 1)

            If LenB(Trim$(myLast.fName)) Then
                TmpLast(n) = myLast
            End If

        Next

        TmpLast(0).fName = CurRecent
        TmpLast(0).fIndex = CurIndex
    End If

    For a% = 0 To mRecentMax
        For b% = 0 To mRecentMax

            If b% <> a% Then
                If LenB(TmpLast(a%).fName) Then
                    If TmpLast(a%).fName = TmpLast(b%).fName Then
                        TmpLast(b%).fName = vbNullString
                        b% = b% - 1
                    End If
                End If
            End If

        Next
    Next

    m = 0

    ReDim mRecent(mRecentMax)

    For n = 0 To mRecentMax - 1

        If LenB(Trim$(TmpLast(n).fName)) Then
            mRecent(m).fName = TmpLast(n).fName
            mRecent(m).fIndex = TmpLast(n).fIndex
            mRecent(m).fRecent = True
            m = m + 1
        End If

    Next

    mRecentCount = m
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function SetValue
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   MyHkey (HkeyLoc)
'                              myKey (String)
'                              MyValue (String)
'                              MyData (String)
'!--------------------------------------------------------------------------------
Private Function SetValue(MyHkey As HkeyLoc, myKey As String, MyValue As String, ByVal MyData As String)

    On Error GoTo SetValue_Error

    Resultat = 0
    Ident = 0
    TailleBuffer = 0
    Resultat = RegCreateKey(MyHkey, myKey, Ident)

    If Resultat = 0 Then
        Resultat = RegSetValueEx(Ident, MyValue, 0&, 1, ByVal MyData, Len(MyData) + 1)
    End If

    On Error GoTo 0

    Exit Function

SetValue_Error:
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ShowFont
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   fName (String)
'!--------------------------------------------------------------------------------
Private Sub ShowFont(fName As String)

    Dim TRC        As RECT
    Dim tStr       As String
    Dim Br         As Long
    Dim tC         As Long

    If fName = vbNullString Or mShowPreview = False Then

        Exit Sub

    End If

    If Trim$(mPreviewText) = vbNullString Then
        tStr = fName
    Else
        tStr = mPreviewText
    End If

    If fName <> OldFont Then
        OldFont = fName
    Else

        Exit Sub

    End If

    PicPreview.FontName = fName
    PicPreview.FontSize = mPreviewSize
    PicPreview.FontBold = False
    PicPreview.FontItalic = False
    PicPreview.Cls
    PicPreview.Height = (PicPreview.TextHeight(tStr) * Screen.TwipsPerPixelY) + 200
    PicPreview.Width = (PicPreview.TextWidth(tStr) * Screen.TwipsPerPixelX) + 200

    If PicPreview.Width > Screen.Width / 2 Then PicPreview.Width = Screen.Width / 2
    If Screen.Width - (PicList.Left + PicList.Width) < PicPreview.Width Then
        PicPreview.Left = PicList.Left - PicPreview.Width
    Else
        PicPreview.Left = PicList.Left + PicList.Width
    End If

    SetRect TRC, 0, 0, PicPreview.ScaleWidth, PicPreview.ScaleHeight
    DrawTxt PicPreview.hDC, tStr, TRC, MiddleCenter, False, True, True

    If mShowFontName = True Then
        OleTranslateColor mComboForeColor, 0, tC
        Br = CreateSolidBrush(vbBlack)
        PicPreview.FontName = "Tahoma"
        PicPreview.FontSize = 8
        PicPreview.FontBold = False
        PicPreview.FontItalic = False
        PicPreview.Height = PicPreview.Height + (PicPreview.TextHeight("X") * Screen.TwipsPerPixelY)
        SetRect TRC, -1, PicPreview.ScaleHeight - PicPreview.TextHeight("X") - 2, PicPreview.ScaleWidth + 1, PicPreview.ScaleHeight + 1
        DrawTxt PicPreview.hDC, fName, TRC, MiddleCenter
        FrameRect PicPreview.hDC, TRC, Br
        DeleteObject Br
    End If

    If PicPreview.Visible = False Then PicPreview.Visible = True
    PicPreview.Refresh
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ShowList
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ShowList()

    Dim cb As RECT

    CloseMe = False
    GetWindowRect UserControl.hWnd, cb
    tPos = mListPos
    PicList.Width = ScaleX(mComboWidth, vbPixels, vbTwips)
    PicList.Height = ScaleY(((mComboFontSize * 2) * (mComboFontCount + mRecentCount) + 2), vbPixels, vbTwips)
    VScroll1.Move PicList.ScaleWidth - 18, (mComboFontSize * 2) * mRecentCount, 18, PicList.ScaleHeight - ((mComboFontSize * 2) * mRecentCount)

    If cb.Bottom + (PicList.Height / Screen.TwipsPerPixelY) < Screen.Height / Screen.TwipsPerPixelY Then
        SetWindowPos PicList.hWnd, HWND_TOP, cb.Left, cb.Bottom, PicList.Width / Screen.TwipsPerPixelX, PicList.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Else
        SetWindowPos PicList.hWnd, HWND_TOP, cb.Left, cb.Top - (PicList.Height / Screen.TwipsPerPixelY), PicList.Width / Screen.TwipsPerPixelX, PicList.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    End If

    SetWindowPos PicPreview.hWnd, HWND_TOP, (PicList.Left + PicList.Width) / Screen.TwipsPerPixelX, (PicList.Top / Screen.TwipsPerPixelY), PicPreview.Width / Screen.TwipsPerPixelX, PicPreview.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE
    fPos = mListPos
    DrawList
    UserControl.SetFocus
    TmrFocus.Enabled = True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SortList
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SortList()

    Dim n      As Long
    Dim tStart As Long
    Dim tEnd   As Long
    Dim bStr1  As String
    Dim bStr2  As String
    Dim qRec   As Long

    mgSort 0, mListCount
    tStart = 0

    Do
        bStr1 = mListFont(tStart)
        qRec = 0

        For n = tStart To mListCount
            bStr2 = mListFont(n)

            If LCase$(bStr1) = LCase$(bStr2) Then
                qRec = qRec + 1
            Else

                Exit For

            End If

        Next

        tEnd = tStart + qRec
        mgSort tStart, tEnd - 1
        tStart = tEnd
    Loop While tEnd < mListCount

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SwapStrings
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   String1 (String)
'                              String2 (String)
'!--------------------------------------------------------------------------------
Private Sub SwapStrings(String1 As String, String2 As String)

    Dim tHold As Long

    CopyMemory tHold, ByVal VarPtr(String1), 4
    CopyMemory ByVal VarPtr(String1), ByVal VarPtr(String2), 4
    CopyMemory ByVal VarPtr(String2), tHold, 4
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub TmrAutoText_Timer
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub TmrAutoText_Timer()
    mAutoText = vbNullString
    TmrAutoText.Enabled = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub TmrFocus_Timer
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub TmrFocus_Timer()

    Dim Focus As Long

    Focus = GetFocus

    If (Focus <> UserControl.hWnd) Or CloseMe = True Then
        bCancel = True
        PicPreview.Visible = False
        PicList.Visible = False
        TmrFocus.Enabled = False
        CloseMe = True
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub TmrOver_Timer
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub TmrOver_Timer()

    Dim Pos As POINTAPI
    Dim WFP As Long

    GetCursorPos Pos
    WFP = WindowFromPoint(Pos.X, Pos.Y)

    If WFP <> Me.hWnd Then
        DrawControl bUp
        TmrOver.Enabled = False
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_DblClick
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_GotFocus()

    If mShowFocus = True Then
        FocusBox.Visible = True
    Else
        FocusBox.Visible = False
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_InitProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()
    mEnabled = True
    mPreviewText = Ambient.DisplayName
    mBorderStyle = sSunken
    mButtonBorderStyle = sRaised
    mShowPreview = True
    mShowFontName = True
    mPreviewSize = 36
    mSorted = True
    mShowFontInCombo = True
    mComboFontCount = 20
    mComboFontSize = 8
    mComboFontBold = False
    mComboFontItalic = False
    mComboWidth = 250
    mRecentMax = 4
    mRecentBackColor = vbWindowBackground
    mRecentForeColor = vbWindowText
    mForeColor = vbWindowText
    mBackColor = vbWindowBackground
    mComboForeColor = vbWindowText
    mComboBackColor = vbWindowBackground
    mComboSelectColor = vbHighlight
    mButtonBackColor = vbButtonFace
    mButtonForeColor = vbButtonText
    mUseMouseWheel = False
    Set UserControl.Font = Ambient.Font
    mUsedBackColor = vbInfoBackground
    mUsedForeColor = vbInfoText
    mXPStyle = True
    mShowFocus = True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_KeyDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim kCode As String
    Dim fI    As Integer
    Dim kC    As Boolean

    If PicList.Visible = True Then

        Select Case KeyCode

            Case vbKeyUp

                If VScroll1.Value Then
                    VScroll1.Value = VScroll1.Value - 1
                End If

            Case vbKeyDown

                If VScroll1.Value < VScroll1.Max Then
                    VScroll1.Value = VScroll1.Value + 1
                End If

            Case vbKeyPageUp

                If VScroll1.Value - VScroll1.LargeChange Then
                    VScroll1.Value = VScroll1.Value - VScroll1.LargeChange
                Else
                    VScroll1.Value = VScroll1.Min
                End If

            Case vbKeyPageDown

                If VScroll1.Value + VScroll1.LargeChange < VScroll1.Max Then
                    VScroll1.Value = VScroll1.Value + VScroll1.LargeChange
                Else
                    VScroll1.Value = VScroll1.Max
                End If

            Case vbKeyHome
                VScroll1.Value = 0

            Case vbKeyEnd
                VScroll1.Value = VScroll1.Max
        End Select

        If mSorted = False Then

            Exit Sub

        End If

        kCode = LCase$(Chr$(KeyCode))

        If Asc(kCode) >= 97 Then
            If Asc(kCode) <= 122 Then
                kC = mAutoText = kCode

                If kC = False Then
                    mAutoText = mAutoText & kCode
                End If

                fI = FontExist(mAutoText & "*", mListPos + IIf(kC = True, 1, 0))

                ' check from current position
                If fI >= 0 Then
                    TmrAutoText.Enabled = False
                    mListPos = fI

                    If fI <= VScroll1.Max Then
                        VScroll1.Value = fI
                    Else
                        VScroll1.Value = VScroll1.Max
                    End If

                    SelBox.Move 0, (fI - VScroll1.Value + mRecentCount) * (mComboFontSize * 2), PicList.ScaleWidth + 2, (mComboFontSize * 2) + 2

                    If kC = False Then
                        TmrAutoText.Interval = 1500
                    Else
                        TmrAutoText.Interval = 800
                    End If

                    TmrAutoText.Enabled = True
                Else
                    fI = FontExist(mAutoText & "*")

                    'check from position 0
                    If fI >= 0 Then
                        TmrAutoText.Enabled = False
                        mListPos = fI
                        VScroll1.Value = fI
                        TmrAutoText.Interval = 1500
                        TmrAutoText.Enabled = True
                    Else
                        mAutoText = vbNullString
                    End If
                End If
            End If
        End If
    End If

    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_KeyPress
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyAscii (Integer)
'!--------------------------------------------------------------------------------
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_KeyUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_LostFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_LostFocus()
    FocusBox.Visible = False
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
    On Local Error Resume Next

    If Button = 1 Then
        inRct = PtInRect(uRct, X, Y)

        If inRct = True Then
            DrawControl bDown, True
            DoEvents

            If PicList.Visible = False Then
                ShowList
            Else
                PicList.Visible = False
                PicPreview.Visible = False
                TmrFocus.Enabled = False
                CloseMe = True
            End If
        End If
    End If

    RaiseEvent MouseDown(Button, Shift, X, Y)
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

    If Button = 0 Then
        DrawControl bOver, True
        TmrOver.Enabled = True
    End If

    RaiseEvent MouseMove(Button, Shift, X, Y)
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

    If Button = 1 Then
        If inRct = True Then DrawControl bUp
        inRct = False
    End If

    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_ReadProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropBag (PropertyBag)
'!--------------------------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        mEnabled = .ReadProperty("Enabled", True)
        mPreviewText = .ReadProperty("PreviewText", Ambient.DisplayName)
        mBorderStyle = .ReadProperty("BorderStyle", sSunken)
        mButtonBorderStyle = .ReadProperty("ButtonBorderStyle", sRaised)
        mShowPreview = .ReadProperty("ShowPreview", True)
        mShowFontName = .ReadProperty("ShowFontName", True)
        mPreviewSize = .ReadProperty("PreviewSize", 36)
        mSorted = .ReadProperty("Sorted", True)
        mShowFontInCombo = .ReadProperty("ShowFontInCombo", True)
        mComboFontCount = .ReadProperty("ComboFontCount", 20)
        mComboFontSize = .ReadProperty("ComboFontSize", 8)
        mComboFontBold = .ReadProperty("ComboFontBold", False)
        mComboFontItalic = .ReadProperty("ComboFontItalic", False)
        mComboWidth = .ReadProperty("ComboWidth", 250)
        mRecentMax = .ReadProperty("RecentMax", 4)
        mRecentBackColor = .ReadProperty("RecentBackColor", vbWindowBackground)
        mRecentForeColor = .ReadProperty("RecentForeColor", vbWindowText)
        mForeColor = .ReadProperty("ForeColor", vbWindowText)
        mBackColor = .ReadProperty("BackColor", vbWindowBackground)
        mComboForeColor = .ReadProperty("ComboForeColor", vbWindowText)
        mComboBackColor = .ReadProperty("ComboBackColor", vbWindowBackground)
        mComboSelectColor = .ReadProperty("ComboSelectColor", vbHighlight)
        mButtonBackColor = .ReadProperty("ButtonBackColor", vbButtonFace)
        mButtonForeColor = .ReadProperty("ButtonForeColor", vbButtonText)
        mButtonOverColor = .ReadProperty("ButtonOverColor", vbBlue)
        mUseMouseWheel = .ReadProperty("UseMouseWheel", False)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        mUsedBackColor = .ReadProperty("UsedBackColor", vbInfoBackground)
        mUsedForeColor = .ReadProperty("UsedForeColor", vbInfoText)
        mXPStyle = .ReadProperty("XPStyle", True)
        mShowFocus = .ReadProperty("ShowFocus", True)
    End With

    UserControl.ForeColor = mForeColor
    UserControl.BackColor = mBackColor
    FocusBox.BackColor = mComboSelectColor

    ReDim mRecent(mRecentMax)

    If RunMode = True Then
        FillList

        If mSorted = True Then SortList
    End If

    DrawControl , True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Resize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Resize()

    Dim tBdr As Single
    Dim v    As Integer

    On Error Resume Next

    If mXPStyle = False Then
        v = 0

        Select Case mBorderStyle

            Case sNone
                tBdr = 0

            Case sSmoothRaised, sSmoothSunken
                tBdr = 1

            Case Else
                tBdr = 2
        End Select

    Else
        v = 2
        tBdr = 1
    End If

    UserControl.Height = ScaleY(TextHeight("X") + (tBdr * 2) + 4 + v, vbPixels, vbTwips)

    If UserControl.Width < 600 Then UserControl.Width = 600
    FocusBox.Move tBdr + 1, tBdr + 1, UserControl.ScaleWidth - tBdr - 20 + v, UserControl.ScaleHeight - (tBdr * 2) - 1
    SetRect uRct, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    SetRect Btn, UserControl.ScaleWidth - tBdr - 17, tBdr, UserControl.ScaleWidth - tBdr, UserControl.ScaleHeight - tBdr
    DrawControl bUp, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_WriteProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropBag (PropertyBag)
'!--------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Enabled", mEnabled, True
        .WriteProperty "PreviewText", mPreviewText, Ambient.DisplayName
        .WriteProperty "BorderStyle", mBorderStyle, sSunken
        .WriteProperty "ButtonBorderStyle", mButtonBorderStyle, sRaised
        .WriteProperty "ShowPreview", mShowPreview, True
        .WriteProperty "ShowFontName", mShowFontName, True
        .WriteProperty "PreviewSize", mPreviewSize, 36
        .WriteProperty "Sorted", mSorted, True
        .WriteProperty "ShowFontInCombo", mShowFontInCombo, True
        .WriteProperty "ComboFontCount", mComboFontCount, 20
        .WriteProperty "ComboFontSize", mComboFontSize, 8
        .WriteProperty "ComboFontBold", mComboFontBold, False
        .WriteProperty "ComboFontItalic", mComboFontItalic, False
        .WriteProperty "ComboWidth", mComboWidth, 250
        .WriteProperty "RecentMax", mRecentMax, 4
        .WriteProperty "RecentBackColor", mRecentBackColor, vbWindowBackground
        .WriteProperty "RecentForeColor", mRecentForeColor, vbWindowText
        .WriteProperty "ForeColor", mForeColor, vbWindowText
        .WriteProperty "BackColor", mBackColor, vbWindowBackground
        .WriteProperty "ComboForeColor", mComboForeColor, vbWindowText
        .WriteProperty "ComboBackColor", mComboBackColor, vbWindowBackground
        .WriteProperty "ComboSelectColor", mComboSelectColor, vbHighlight
        .WriteProperty "ButtonBackColor", mButtonBackColor, vbButtonFace
        .WriteProperty "ButtonForeColor", mButtonForeColor, vbButtonText
        .WriteProperty "ButtonOverColor", mButtonOverColor, vbBlue
        .WriteProperty "UseMouseWheel", mUseMouseWheel, False
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "UsedBackColor", mUsedBackColor, vbInfoBackground
        .WriteProperty "UsedForeColor", mUsedForeColor, vbInfoText
        .WriteProperty "XPStyle", mXPStyle, True
        .WriteProperty "ShowFocus", mShowFocus, True
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub VScroll1_Change
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub VScroll1_Change()

    Dim tFont As String

    If doNothing = True Then

        Exit Sub

    End If

    fPos = VScroll1.Value
    DrawList
    tFont = fList(fPos - VScroll1.Value).fName
    ShowFont tFont
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub VScroll1_GotFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub VScroll1_GotFocus()
    PicList.SetFocus
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub VScroll1_KeyDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub VScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
    UserControl_KeyDown KeyCode, Shift
End Sub
