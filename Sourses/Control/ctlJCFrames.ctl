VERSION 5.00
Begin VB.UserControl ctlJCFrames 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HitBehavior     =   0  'None
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ToolboxBitmap   =   "ctlJCFrames.ctx":0000
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4410
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "ctlJCFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'============================================================================================
'   jcFrames v 1.0 Copyright © 2005.All rights reserved.
'   Juan Carlos San Romбn Arias (sanroman2004@yahoo.com)
'
'   You may use this control in your applications free of charge,
'   provided that you do not redistribute this source code without
'   giving me credit for my work.  Of course, credit in your
'   applications is always welcome.
'
'   Thanks to Jim K for doing the initial idea of the usercontrol using
'   my job posted in PSC
'
'   Thanks to ElectroZ for his frame style used here as TextBox style
'============================================================================================
'
'   Modifications: Paul R. Territo, Ph.D
'
'   The following code is based on the above authors submission which
'   can be found at the follow URL:
'   http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=63827&lngWId=1
'
'   29Dec05 - Moved all external API drawing and Type structures into UserControl
'       eliminate the need for external dependancies (i.e. OCX). This provides
'       a single drop in place UserControl which follows the general rules of
'       encapsulation (i.e. self-contained).
'
'============================================================================================
'  -----------------------------
'  Version 1.1.0 - 29 Dec. 2005
'  -----------------------------
'   Thanks to Paul R. Territo, Ph.D for your advices and usercontrol modification.
'   - usercontrol includes now API drawing and type declaration (no more mods in usercontrol)
'   - Added icon alignment (left and right)
'   - caption alignment takes into consideration if icon picture exists and its alignment
'============================================================================================
'  -----------------------------
'  Version 1.2.0 - 04 Jan. 2006
'  -----------------------------
'   - Added different header styles for Windows frame style (txtboxcolor and gradient)
'   - Added different gradient styles for header gardient style for Windows frame style
'     (horizontal, vertical and cilinder)
'   - Caption is trimmed when its width exceeds control width
'============================================================================================
'  -----------------------------
'  Version 2.0.0 - 11 Jan. 2006
'  -----------------------------
'   - 4 new styles have been added: Inner widge, Outer widge, Header and Panel
'   - Header styles have been extended for other frame style (messenger, jcGradient
'     textbox and panel style)
'   - Control structure was reorganized
'   - Gradientframe style was renamed as jcGradient
'   - Added TxtBoxShadow property for textbox style
'   - Added multiline caption for Panel style
'============================================================================================
'  ----------------------------
'  Version 2.0.1 - 8 Feb. 2006
'  ----------------------------
'   - Added enabled property (it enables or disables all the controls in usercontrol)
'   - Added TransBlt from Chameleon button to draw grayscale image when control is disabled
'============================================================================================
Option Explicit

'Mudar a borda para mudar tamanho
Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1

'*************************************************************
'   Required Type Definitions
'*************************************************************
Public Enum jcStyleConst
    XPDefault = 0
    jcGradient = 1
    TextBox = 2
    Windows = 3
    Messenger = 4
    InnerWedge = 5
    OuterWedge = 6
    Header = 7
    Panel = 8
End Enum

#If False Then

    Private XPDefault, jcGradient, TextBox, Windows, Messenger, InnerWedge, OuterWedge, Header, Panel
#End If

'xp theme
Public Enum jcThemeConst
    Blue = 0
    Silver = 1
    Olive = 2
    Visual2005 = 3
    Norton2004 = 4
    Custom = 5
    xThemeDarkBlue = 6
    xThemeGreen = 7
    xThemeOffice2003Style2 = 8
    xThemeMetallic = 9
    xThemeOrange = 10
    xThemeTurquoise = 11
    xThemeGray = 12
    xThemeDarkBlue2 = 13
    xThemeMoney = 14
    xThemeOffice2003Style1 = 15
End Enum

'Responsбvel por mover o form
Public Enum jcResp
    jcTitulo = 0
    jcPainel = 1
    jcAmbos = 2
End Enum

'gradient type
Public Enum jcGradConst
    VerticalGradient = 0
    HorizontalGradient = 1
    VCilinderGradient = 2
    HCilinderGradient = 3
End Enum

#If False Then

    Private VerticalGradient, HorizontalGradient, VCilinderGradient, HCilinderGradient
#End If

'header style
Public Enum jcHeaderConst
    TxtBoxColor = 0
    Gradient = 1
End Enum

#If False Then

    Private TxtBoxColor, Gradient
#End If

'TxtBox style
Public Enum jcShadowConst
    [No shadow] = 0
    Shadow = 1
End Enum

#If False Then

    Private Shadow
#End If

'icon aligment
Public Enum IconAlignConst
    vbLeftAligment = 0
    vbRightAligment = 1
End Enum

#If False Then

    Private vbLeftAligment, vbRightAligment
#End If

Enum m_PanelArea
    xTitle = 0
    xPanel = 1
End Enum

Private useMask As Boolean
Private useGrey As Boolean

'*************************************************************
'   Required API Declarations
'*************************************************************
Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Type POINT
    X                                   As Long
    Y                                   As Long
End Type

Private Type RGB
    Red                                 As Byte
    Green                               As Byte
    Blue                                As Byte
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

Private Type BITMAPINFO
    bmiHeader                           As BITMAPINFOHEADER
    bmiColors                           As RGB
End Type

Private Declare Sub ReleaseCapture Lib "user32.dll" ()
Private Declare Function OleTranslateColor Lib "OlePro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CopyRect Lib "user32.dll" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal XLeft As Long, ByVal YTop As Long, ByVal hIcon As Long, ByVal CXWidth As Long, ByVal CYWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long

'FORM TRANSPARENTE
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Private Const TransColor       As Long = &H8000000F
Private Const GWL_STYLE        As Long = -16
Private Const GWL_EXSTYLE      As Long = (-20)
Private Const WS_CAPTION       As Long = &HC00000

'*************************************************************
'   DRAW TEXT
'*************************************************************
' --Formatting Text Consts
Private Const DT_LEFT          As Long = &H0
Private Const DT_CENTER        As Long = &H1
Private Const DT_RIGHT         As Long = &H2
Private Const DT_NOCLIP        As Long = &H100
Private Const DT_WORDBREAK     As Long = &H10
Private Const DT_CALCRECT      As Long = &H400
Private Const DT_RTLREADING    As Long = &H20000
Private Const DT_DRAWFLAG      As Long = DT_CENTER Or DT_WORDBREAK
Private Const DT_TOP           As Long = &H0
Private Const DT_BOTTOM        As Long = &H8
Private Const DT_VCENTER       As Long = &H4
Private Const DT_SINGLELINE    As Long = &H20
Private Const DT_WORD_ELLIPSIS As Long = &H40000

Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Private Declare Function DrawTextExW Lib "user32.dll" (ByVal hDC As Long, ByVal lpsz As Long, ByVal n As Long, ByRef lpRect As RECT, ByVal dwDTFormat As Long, ByRef lpDrawTextParams As DRAWTEXTPARAMS) As Long

'*************************************************************
'   FONT PROPERTIES
'*************************************************************
Private Const LF_FACESIZE     As Long = 32
Private Const FW_NORMAL       As Long = 400
Private Const FW_BOLD         As Long = 700
Private Const DEFAULT_QUALITY As Long = 0

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

Private Const WM_SETFONT       As Long = &H30
Private Const WS_EX_RTLREADING As Long = &H2000

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private FrameFontHandle       As Long
Private FrameLogFont          As LOGFONT
Private WithEvents PropFont   As StdFont
Attribute PropFont.VB_VarHelpID = -1

'*************************************************************
'   UPDATE WINDOW
'*************************************************************
Private Const RDW_UPDATENOW   As Long = &H100
Private Const RDW_INVALIDATE  As Long = &H1
Private Const RDW_ERASE       As Long = &H4
Private Const RDW_ALLCHILDREN As Long = &H80

Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

'*************************************************************
'   Members
'*************************************************************
Private m_FrameColor             As OLE_COLOR
Private m_TextBoxColor           As OLE_COLOR
Private m_BackColor              As OLE_COLOR
Private m_FillColor              As OLE_COLOR
Private m_FrameColorDis          As OLE_COLOR
Private m_TextBoxColorDis        As OLE_COLOR
Private m_FillColorDis           As OLE_COLOR
Private jcColorToDis             As OLE_COLOR
Private jcColorFromDis           As OLE_COLOR
Private jcColorBorderPicDis      As OLE_COLOR
Private m_FrameColorIni          As OLE_COLOR
Private m_TextBoxColorIni        As OLE_COLOR
Private m_FillColorIni           As OLE_COLOR
Private jcColorToIni             As OLE_COLOR
Private jcColorFromIni           As OLE_COLOR
Private jcColorBorderPicIni      As OLE_COLOR
Private m_Caption                As String
Private m_Enabled                As Boolean
Private m_TextBoxHeight          As Long
Private m_TextHeight             As Long
Private m_TextWidth              As Long
Private m_Height                 As Long
Private m_TextColor              As Long
Private m_Alignment              As Long
Private m_RoundedCorner          As Boolean
Private m_RoundedCornerTxtBox    As Boolean
Private m_Style                  As jcStyleConst
Private m_HeaderStyle            As jcHeaderConst
Private m_GradientHeaderStyle    As jcGradConst
Private m_Icon                   As StdPicture
Private m_IconSize               As Integer
Private m_IconAlignment          As IconAlignConst
Private m_ThemeColor             As jcThemeConst
Private m_ColorTo                As OLE_COLOR
Private m_ColorFrom              As OLE_COLOR
Private m_Indentation            As Integer
Private m_Space                  As Integer
Private m_TxtBoxShadow           As jcShadowConst
Private jcTextBoxCenter          As Long
Private jcTextDrawParams         As Long
Private jcColorTo                As OLE_COLOR
Private jcColorFrom              As OLE_COLOR
Private jcColorBorderPic         As OLE_COLOR
Private jcLpp                    As POINT

Private Const TEXT_INACTIVE      As Long = &H80000011    '&H6A6A6A
Private Const m_Border_Inactive  As Long = &H8000000B
Private Const m_BtnFace_Inactive As Long = &H8000000F
Private Const m_BtnFace          As Long = &H80000016    '&H8000000F '&H80000016&

'*************************************************************
'   Constants
'*************************************************************
Private Const ALTERNATE          As Integer = 1    ' ALTERNATE and WINDING are
Private Const WINDING            As Integer = 2    ' constants for FillMode.
Private Const BLACKBRUSH         As Integer = 4    ' Constant for brush type.
Private Const WHITE_BRUSH        As Integer = 0    ' Constant for brush type.
Private Const RGN_AND            As Integer = 1
Private Const RGN_COPY           As Integer = 5
Private Const RGN_OR             As Integer = 2
Private Const RGN_XOR            As Integer = 3
Private Const RGN_DIFF           As Integer = 4
Private Const m_def_Responsavel = 0
Private Const m_def_AllowDraging = 0
Private Const m_def_AtivarResizeDoForm = False
Private Const m_def_Collapsar = False

Dim temp_height              As Integer

Private m_Responsavel        As jcResp
Private m_AllowDraging       As Boolean
Private m_AllowParentDraging As Boolean
Private m_AtivarResizeDoForm As Boolean
Private m_Collapsar          As Boolean
Private m_Collapsado         As Boolean
Private m_bIsWinXpOrLater    As Boolean

'*************************************************************
'   events
'*************************************************************
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event TileClick()
Event PanelClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, PanelArea As m_PanelArea)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, PanelArea As m_PanelArea)
Event CollapseClick(Button As Integer)

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Alignment
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Alignment
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_Alignment (AlignmentConstants)
'!--------------------------------------------------------------------------------
Public Property Let Alignment(ByRef new_Alignment As AlignmentConstants)
    m_Alignment = new_Alignment
    SetjcTextDrawParams
    PropertyChanged "Alignment"
    PaintFrame
End Property

'==========================================================================
' API Functions and subroutines
'==========================================================================
' full version of APILine
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub APILineEx
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lhdcEx (Long)
'                              X1 (Long)
'                              Y1 (Long)
'                              X2 (Long)
'                              Y2 (Long)
'                              lColor (Long)
'!--------------------------------------------------------------------------------
Private Sub APILineEx(lhdcEx As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lColor As Long)

    'Use the API LineTo for Fast Drawing
    Dim PT      As POINT
    Dim hPen    As Long
    Dim hPenOld As Long

    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(lhdcEx, hPen)
    MoveToEx lhdcEx, X1, Y1, PT
    LineTo lhdcEx, X2, Y2
    SelectObject lhdcEx, hPenOld
    DeleteObject hPen
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function APIRectangle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lngHDc (Long)
'                              X (Long)
'                              Y (Long)
'                              W (Long)
'                              H (Long)
'                              lColor (OLE_COLOR = -1)
'!--------------------------------------------------------------------------------
Private Function APIRectangle(ByVal lngHDc As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional lColor As OLE_COLOR = -1) As Long

    'Draw an api rectangle
    Dim hPen    As Long
    Dim hPenOld As Long
    Dim PT      As POINT

    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(lngHDc, hPen)
    MoveToEx lngHDc, X, Y, PT
    LineTo lngHDc, X + W, Y
    LineTo lngHDc, X + W, Y + H
    LineTo lngHDc, X, Y + H
    LineTo lngHDc, X, Y
    SelectObject lngHDc, hPenOld
    DeleteObject hPen
End Function

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
'! Parameters  (Переменные):   new_BackColor (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let BackColor(ByRef new_BackColor As OLE_COLOR)
    m_BackColor = TranslateColor(new_BackColor)
    UserControl.BackColor = m_BackColor
    PropertyChanged "BackColor"
    PaintFrame
End Property

'Blend two colors
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function BlendColors
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lColor1 (Long)
'                              lColor2 (Long)
'!--------------------------------------------------------------------------------
Private Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long) As Single
    BlendColors = RGB(((lColor1 And &HFF) + (lColor2 And &HFF)) / 2, (((lColor1 \ &H100) And &HFF) + ((lColor2 \ &H100) And &HFF)) / 2, (((lColor1 \ &H10000) And &HFF) + ((lColor2 \ &H10000) And &HFF)) / 2)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Caption
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Caption() As String
    Caption = m_Caption
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Caption
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_Caption (String)
'!--------------------------------------------------------------------------------
Public Property Let Caption(ByRef New_Caption As String)
    m_Caption = New_Caption
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ColorFrom
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ColorFrom() As OLE_COLOR
    ColorFrom = m_ColorFrom
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ColorFrom
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_ColorFrom (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ColorFrom(ByRef new_ColorFrom As OLE_COLOR)
Attribute ColorFrom.VB_Description = "Returns/Sets the Start color for gradient"
    m_ColorFrom = TranslateColor(new_ColorFrom)

    If m_ThemeColor = Custom Then
        jcColorFromIni = m_ColorFrom
    End If

    PropertyChanged "ColorFrom"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ColorTo
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ColorTo() As OLE_COLOR
    ColorTo = m_ColorTo
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ColorTo
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_ColorTo (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ColorTo(ByRef new_ColorTo As OLE_COLOR)
Attribute ColorTo.VB_Description = "Returns/Sets the End color for gradient"
    m_ColorTo = TranslateColor(new_ColorTo)

    If m_ThemeColor = Custom Then
        jcColorToIni = m_ColorTo
    End If

    PropertyChanged "ColorTo"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Draw_Header
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   R_Caption (RECT)
'!--------------------------------------------------------------------------------
Private Sub Draw_Header(R_Caption As RECT)

    Dim p_left As Long

    APILineEx UserControl.hDC, 0&, jcTextBoxCenter, UserControl.ScaleWidth, jcTextBoxCenter, IIf(m_Enabled, TranslateColor(&H80000015), TranslateColor(TEXT_INACTIVE))
    'TranslateColor(&H80000015)&H808080
    APILineEx UserControl.hDC, 0&, jcTextBoxCenter + 1, UserControl.ScaleWidth, jcTextBoxCenter + 1, vbWhite

    If LenB(m_Caption) <> 0 Then
        If m_Alignment = vbLeftJustify Then
            'm_Indentation
        ElseIf m_Alignment = vbRightJustify Then
            p_left = UserControl.ScaleWidth - m_TextWidth - m_Space
        Else
            p_left = (UserControl.ScaleWidth - m_TextWidth) / 2
        End If

        'Draw a line
        APILineEx UserControl.hDC, p_left, jcTextBoxCenter, p_left + m_TextWidth + m_Space, jcTextBoxCenter, m_FillColor
        'TranslateColor(Ambient.BackColor)
        APILineEx UserControl.hDC, p_left, jcTextBoxCenter + 1, p_left + m_TextWidth + m_Space, jcTextBoxCenter + 1, m_FillColor
        'TranslateColor(Ambient.BackColor)
        'set caption rect
        SetRect R_Caption, p_left + m_Space / 2, 0, m_TextWidth + p_left + m_Space / 2, m_TextHeight
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Draw_InnerWedge
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   R_Caption (RECT)
'!--------------------------------------------------------------------------------
Private Sub Draw_InnerWedge(R_Caption As RECT)

    Dim txtWidth        As Integer
    Dim txtHeight       As Integer
    Dim R               As RECT
    Dim m_roundedRadius As Long
    Dim hFRgn           As Long
    Dim poly(1 To 4)    As POINT
    Dim NumCoords       As Long
    Dim hBrush          As Long
    Dim hRgn            As Long

    m_roundedRadius = IIf(m_RoundedCorner = False, 0&, 10&)
    txtWidth = m_TextWidth + 10

    If txtWidth < 100 Then
        txtWidth = 100
    End If

    txtHeight = m_TextHeight + 5
    NumCoords = 4
    SetRect R, 0&, 0&, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1

    If (txtWidth + R.Left + txtHeight / 2) >= R.Right - m_Indentation Then
        txtWidth = R.Right - txtHeight / 2 - R.Left - m_Indentation - 1
    End If

    'Assign values to points.
    poly(1).X = R.Left
    poly(1).Y = R.Top
    poly(2).X = R.Left
    poly(2).Y = R.Top + txtHeight
    poly(3).X = R.Left + txtWidth
    poly(3).Y = R.Top + txtHeight
    poly(4).X = R.Left + txtWidth + txtHeight / 2
    poly(4).Y = R.Top
    'Creates first region to fill with color.
    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
    'Creates second region to fill with color.
    hFRgn = CreateRoundRectRgn(R.Left, R.Top, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius)
    'Combine our two regions
    CombineRgn hRgn, hRgn, hFRgn, RGN_AND
    'delete second region
    DeleteObject hFRgn
    'fill frame
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FillColor, R
    'If the creation of the region was successful then color.
    hBrush = CreateSolidBrush(m_TextBoxColor)

    If hRgn Then
        FillRgn UserControl.hDC, hRgn, hBrush
    End If

    'draw frame borders
    APILineEx UserControl.hDC, poly(2).X, poly(2).Y, poly(3).X, poly(3).Y, m_FrameColor
    APILineEx UserControl.hDC, poly(3).X, poly(3).Y, poly(4).X, poly(4).Y, m_FrameColor
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R, True
    'delete created region
    DeleteObject hRgn
    DeleteObject hBrush
    'set caption rectangle
    SetRect R_Caption, poly(1).X + m_Indentation / 2, poly(1).Y, txtWidth + poly(1).X, txtHeight + poly(1).Y + 2
    '    'set icon coordinates
    '   iY = (txtHeight - m_IconSize) / 2
    UserControl.FillStyle = 0
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Draw_jcGradient
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   R_Caption (RECT)
'                              iY (Integer)
'!--------------------------------------------------------------------------------
Private Sub Draw_jcGradient(R_Caption As RECT, iY As Integer)

    Dim R As RECT

    jcTextBoxCenter = m_TextBoxHeight / 2
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, BlendColors(jcColorFrom, vbWhite), IIf(m_ThemeColor = Custom, m_FrameColor, jcColorBorderPic), R
    'Draw header
    SetRect R, 0, 0, UserControl.ScaleWidth - 2, m_Height
    DrawGradientInRectangle UserControl.hDC, jcColorTo, jcColorFrom, R, VCilinderGradient, True, jcColorBorderPic

    If m_HeaderStyle = Gradient Then
        SetRect R, 0, m_Height, UserControl.ScaleWidth - 2, m_TextBoxHeight
        DrawGradientInRectangle UserControl.hDC, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, jcColorBorderPic
    Else
        SetRect R, 0, m_Height, UserControl.ScaleWidth - 1, m_TextBoxHeight + m_Height + 2
        DrawAPIRoundRect False, 0&, m_FillColor, m_FrameColor, R
    End If

    With UserControl
        SetRect R, 0, m_Height + m_TextBoxHeight, .ScaleWidth - 2, m_Height
        DrawGradientInRectangle .hDC, jcColorTo, jcColorFrom, R, VCilinderGradient, True, jcColorBorderPic
        SetRect R, 1, m_Height * 2 + m_TextBoxHeight, .ScaleWidth - 3, .ScaleHeight - (2 + m_Height * 2 + m_TextBoxHeight) - .ScaleHeight * 0.2
        DrawGradientInRectangle .hDC, BlendColors(jcColorFrom, vbWhite), BlendColors(jcColorTo, vbWhite), R, VerticalGradient, False, m_TextBoxColor
        'set caption rect
        SetRect R_Caption, m_Space, m_Height + 1, .ScaleWidth - 2 - m_Space, m_TextBoxHeight + 2
        'set icon Y coordinate
    End With

    iY = (m_Height * 2 + m_TextBoxHeight - m_IconSize) / 2
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Draw_Messenger
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   R_Caption (RECT)
'                              iY (Integer)
'!--------------------------------------------------------------------------------
Private Sub Draw_Messenger(R_Caption As RECT, iY As Integer)

    Dim R As RECT

    jcTextBoxCenter = 0
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, BlendColors(jcColorFrom, vbWhite), IIf(m_ThemeColor = Custom, m_FrameColor, jcColorBorderPic), R
    'Draw header
    SetRect R, 0, 0, UserControl.ScaleWidth - 2, m_Height * 2
    DrawGradientInRectangle UserControl.hDC, jcColorFrom, vbWhite, R, VerticalGradient, True, jcColorBorderPic, 2.01
    PaintShpInBar vbWhite, BlendColors(vbBlack, jcColorFrom), m_Height * 2

    If m_HeaderStyle = Gradient Or m_Enabled = False Then
        SetRect R, 0&, m_Height * 2, UserControl.ScaleWidth - 2, m_TextBoxHeight + 1
        DrawGradientInRectangle UserControl.hDC, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, jcColorBorderPic
    Else
        SetRect R, 0, m_Height * 2 + m_TextBoxHeight + 1, UserControl.ScaleWidth - 2, m_Height * 2 + m_TextBoxHeight + 1
        APILineEx UserControl.hDC, R.Left, R.Top, R.Right, R.Bottom, jcColorBorderPic
        'vbBlack
    End If

    SetRect R, 1, 1 + m_Height * 2 + m_TextBoxHeight, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - (2 + m_Height * 2 + m_TextBoxHeight) - UserControl.ScaleHeight * 0.2
    DrawGradientInRectangle UserControl.hDC, BlendColors(jcColorFrom, vbWhite), BlendColors(jcColorTo, vbWhite), R, VerticalGradient, False, m_TextBoxColor
    'set caption rect
    SetRect R_Caption, m_Space, m_Height * 2 + 2, UserControl.ScaleWidth - 1 - m_Space, m_TextBoxHeight + 6
    'set icon coordinates
    iY = m_Height * 2 + (m_TextBoxHeight - m_IconSize) / 2
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Draw_OuterWedge
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   R_Caption (RECT)
'!--------------------------------------------------------------------------------
Private Sub Draw_OuterWedge(R_Caption As RECT)

    Dim txtWidth        As Integer
    Dim txtHeight       As Integer
    Dim R               As RECT
    Dim r1              As RECT
    Dim m_roundedRadius As Long
    Dim poly(1 To 4)    As POINT
    Dim NumCoords       As Long
    Dim hBrush          As Long
    Dim hRgn            As Long

    m_roundedRadius = IIf(m_RoundedCorner = False, 0&, 10&)
    txtWidth = m_TextWidth + 10

    If txtWidth < 100 Then
        txtWidth = 100
    End If

    txtHeight = m_TextHeight + 5
    NumCoords = 4
    SetRect R, 0&, 0&, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1

    If (txtWidth + R.Left + txtHeight / 2) >= R.Right - m_Indentation Then
        txtWidth = R.Right - txtHeight / 2 - R.Left - m_Indentation - 1
    End If

    'Assign values to points.
    poly(1).X = R.Left + 6
    poly(1).Y = R.Top
    poly(2).X = R.Left + 6
    poly(2).Y = R.Top + txtHeight
    poly(3).X = R.Left + txtWidth + txtHeight / 2
    poly(3).Y = R.Top + txtHeight
    poly(4).X = R.Left + txtWidth
    poly(4).Y = R.Top
    'Creates first region to fill with color.
    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
    'If the creation of the region was successful then color.
    hBrush = CreateSolidBrush(m_TextBoxColor)

    If hRgn Then
        FillRgn UserControl.hDC, hRgn, hBrush
    End If

    'fill frame
    SetRect r1, 0&, 0&, txtWidth * 0.9, txtHeight * 1.3
    DrawAPIRoundRect m_RoundedCorner, 10&, m_TextBoxColor, m_FrameColor, r1
    SetRect r1, txtWidth * 0.9 - 5, 1, txtWidth * 0.9 + 3, txtHeight * 1.3
    DrawAPIRoundRect m_RoundedCorner, 0&, m_TextBoxColor, m_TextBoxColor, r1
    SetRect r1, -1, -1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FillColor, r1, True

    'draw frame borders
    With UserControl
        .ForeColor = m_FrameColor
        APILineEx .hDC, poly(1).X, poly(1).Y, poly(4).X, poly(4).Y, .ForeColor
        APILineEx .hDC, poly(4).X, poly(4).Y, poly(3).X, poly(3).Y, .ForeColor
        RoundRect .hDC, R.Left, R.Top + txtHeight, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius
        RoundRect .hDC, R.Left, R.Top + txtHeight, R.Left + 10, R.Top + txtHeight + 10, 0, 0
        .ForeColor = m_FillColor
        RoundRect .hDC, R.Left + 1, R.Top + txtHeight + 1, R.Left + 10, R.Top + txtHeight + 10, 0, 0
        'delete created region
    End With

    DeleteObject hRgn
    DeleteObject hBrush
    'set caption rectangle
    SetRect R_Caption, poly(1).X + m_Indentation / 2 - 6, poly(1).Y, txtWidth + poly(1).X - 6, txtHeight + poly(1).Y + 2
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Draw_Panel
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   R_Caption (RECT)
'                              iY (Integer)
'!--------------------------------------------------------------------------------
Private Sub Draw_Panel(R_Caption As RECT, iY As Integer)

    Dim R               As RECT
    Dim m_roundedRadius As Long
    Dim hFRgn           As Long
    Dim hRgn            As Long

    jcTextBoxCenter = m_TextBoxHeight / 2
    'Draw border rectangle
    UserControl.FillColor = m_FillColor

    If m_ThemeColor = Custom Or m_HeaderStyle = TxtBoxColor Then
        UserControl.ForeColor = m_FrameColor
    Else
        UserControl.ForeColor = jcColorBorderPic
    End If

    'If m_Enabled = False Then UserControl.ForeColor = m_Border_Inactive
    m_roundedRadius = IIf(m_RoundedCorner = False, 0&, 9&)
    SetRect R, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight

    If m_HeaderStyle = Gradient Then
        DrawGradientInRectangle UserControl.hDC, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, False, UserControl.ForeColor, 2.03
    End If

    'Creates first region to fill with color.
    hRgn = CreateRoundRectRgn(R.Left, R.Top, R.Right, R.Bottom, 0&, 0&)
    'Creates second region to fill with color.
    hFRgn = CreateRoundRectRgn(R.Left, R.Top, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius)
    'Combine our two regions
    CombineRgn hRgn, hRgn, hFRgn, RGN_AND
    'delete second region
    DeleteObject hFRgn
    SetWindowRgn UserControl.hWnd, hRgn, True
    UserControl.FillStyle = IIf(m_HeaderStyle = Gradient, 1, 0)

    If UserControl.ForeColor <> UserControl.BackColor Or m_HeaderStyle = TxtBoxColor Then
        RoundRect UserControl.hDC, R.Left, R.Top, R.Right - 1, R.Bottom - 1, m_roundedRadius, m_roundedRadius
        UserControl.FillStyle = 0
        DrawCorners UserControl.ForeColor
    End If

    'set caption rect
    SetRect R_Caption, m_Space, 0&, UserControl.ScaleWidth - m_Space, UserControl.ScaleHeight - 2
    'set icon coordinates
    iY = (UserControl.ScaleHeight - m_IconSize) / 2
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Draw_TextBox
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   R_Caption (RECT)
'                              iX (Integer)
'                              iY (Integer)
'!--------------------------------------------------------------------------------
Private Sub Draw_TextBox(R_Caption As RECT, iX As Integer, iY As Integer)

    Dim R As RECT

    jcTextBoxCenter = m_TextBoxHeight / 2
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R

    'Draw textbox border rectangle
    If m_HeaderStyle = Gradient Then
        If m_TxtBoxShadow = Shadow Then
            SetRect R, m_Indentation, 0, UserControl.ScaleWidth - 1 - m_Indentation, m_TextBoxHeight
            OffsetRect R, 2, 2
            DrawAPIRoundRect False, m_TextBoxHeight, BlendColors(m_FillColor, &HA7A7A7), BlendColors(m_FillColor, &HA7A7A7), R
        End If

        SetRect R, m_Indentation, 0, UserControl.ScaleWidth - 2 - 2 * m_Indentation, m_TextBoxHeight - 1
        DrawGradientInRectangle UserControl.hDC, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, m_FrameColor
        ', 3.08
    Else
        SetRect R, m_Indentation, 0, UserControl.ScaleWidth - 1 - m_Indentation, m_TextBoxHeight

        If m_TxtBoxShadow = Shadow Then
            OffsetRect R, 2, 2
            DrawAPIRoundRect m_RoundedCornerTxtBox, m_TextBoxHeight, BlendColors(m_FillColor, &HA7A7A7), BlendColors(m_FillColor, &HA7A7A7), R
            OffsetRect R, -2, -2
        End If

        DrawAPIRoundRect m_RoundedCornerTxtBox, m_TextBoxHeight, m_TextBoxColor, m_FrameColor, R
    End If

    'set caption rect
    SetRect R_Caption, m_Indentation + m_Space * 1.5, 0, UserControl.ScaleWidth - 1 - m_Indentation - m_Space * 1.5, m_TextBoxHeight - 1
    'set icon coordinates
    iX = m_Indentation + m_Space * 2
    iY = (m_TextBoxHeight - m_IconSize) / 2
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Draw_Windows
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   R_Caption (RECT)
'                              iY (Integer)
'!--------------------------------------------------------------------------------
Private Sub Draw_Windows(R_Caption As RECT, iY As Integer)

    Dim R As RECT

    jcTextBoxCenter = m_TextBoxHeight / 2
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R

    'Draw text box borders
    If m_HeaderStyle = Gradient Then
        SetRect R, 0&, 0&, UserControl.ScaleWidth - 2, m_TextBoxHeight - 1
        DrawGradientInRectangle UserControl.hDC, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, m_FrameColor
        ', 3.08
    Else
        SetRect R, 0&, 0&, UserControl.ScaleWidth - 1, m_TextBoxHeight
        DrawAPIRoundRect m_RoundedCornerTxtBox, 10&, m_TextBoxColor, m_FrameColor, R
    End If

    'set caption rect
    SetRect R_Caption, m_Space, 0, UserControl.ScaleWidth - m_Space, m_TextBoxHeight
    '- 1
    'set icon coordinates
    iY = (m_TextBoxHeight - m_IconSize) / 2
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Draw_XPDefault
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   R_Caption (RECT)
'!--------------------------------------------------------------------------------
Private Sub Draw_XPDefault(R_Caption As RECT)

    Dim p_left As Long
    Dim R      As RECT

    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R

    If LenB(m_Caption) <> 0 Then
        If m_Alignment = vbLeftJustify Then
            p_left = m_Indentation
        ElseIf m_Alignment = vbRightJustify Then
            p_left = UserControl.ScaleWidth - m_TextWidth - m_Indentation - m_Space - 1
        Else
            p_left = (UserControl.ScaleWidth - 1 - m_TextWidth) / 2
        End If

        'Draw a line
        APILineEx UserControl.hDC, p_left, jcTextBoxCenter, p_left + m_TextWidth + m_Space, jcTextBoxCenter, m_FillColor
        'set caption rect
        SetRect R_Caption, p_left + m_Space / 2, 0, m_TextWidth + p_left + m_Space / 2, m_TextHeight
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawAPIRoundRect
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   blnRounded (Boolean)
'                              LngRoundValue (Long)
'                              MyFillColor (Long)
'                              MyBorderColor (Long)
'                              R (RECT)
'                              blnTransparent (Boolean = False)
'!--------------------------------------------------------------------------------
Private Sub DrawAPIRoundRect(ByVal blnRounded As Boolean, ByVal LngRoundValue As Long, ByVal MyFillColor As Long, ByVal MyBorderColor As Long, R As RECT, Optional ByVal blnTransparent As Boolean = False)

    Dim m_roundedRadius As Long

    With UserControl
        .FillColor = MyFillColor
        .ForeColor = MyBorderColor
        .FillStyle = IIf(blnTransparent, 1, 0)
    End With

    m_roundedRadius = IIf(blnRounded = False, 0&, LngRoundValue)
    RoundRect UserControl.hDC, R.Left, R.Top, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius
    UserControl.FillStyle = 0
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawCorners
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PenColor (Long)
'!--------------------------------------------------------------------------------
Private Sub DrawCorners(PenColor As Long)

    With UserControl
        'left top corner
        SetPixel .hDC, 0, 4, PenColor
        SetPixel .hDC, 4, 0, PenColor
        'left bottom corner
        SetPixel .hDC, .ScaleWidth - 5, 0, PenColor
        SetPixel .hDC, .ScaleWidth - 1, 4, PenColor
        'right top corner
        SetPixel .hDC, 0, .ScaleHeight - 5, PenColor
        SetPixel .hDC, 4, .ScaleHeight - 1, PenColor
        'right bottom corner
        SetPixel .hDC, .ScaleWidth - 5, .ScaleHeight - 1, PenColor
        SetPixel .hDC, .ScaleWidth - 1, .ScaleHeight - 5, PenColor
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawGradCilinder
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lhdcEx (Long)
'                              lStartColor (Long)
'                              lEndColor (Long)
'                              R (RECT)
'                              blnVertical (Boolean = True)
'                              LightCenter (Double = 2.01)
'!--------------------------------------------------------------------------------
Private Sub DrawGradCilinder(lhdcEx As Long, lStartColor As Long, lEndColor As Long, R As RECT, Optional ByVal blnVertical As Boolean = True, Optional ByVal LightCenter As Double = 2.01)

    If LightCenter <= 1# Then
        LightCenter = 1.01
    End If

    If blnVertical Then
        DrawGradientEx lhdcEx, lStartColor, lEndColor, R.Left, R.Top, R.Right + R.Left, R.Bottom / LightCenter, True
        DrawGradientEx lhdcEx, lEndColor, lStartColor, R.Left, R.Top + R.Bottom / LightCenter - 1, R.Right + R.Left, (LightCenter - 1) * R.Bottom / LightCenter + 1, True
    Else
        DrawGradientEx lhdcEx, lStartColor, lEndColor, R.Left, R.Top, R.Right / LightCenter, R.Bottom + R.Top, False
        DrawGradientEx lhdcEx, lEndColor, lStartColor, R.Left + R.Right / LightCenter - 1, R.Top, (LightCenter - 1) * R.Right / LightCenter + 1, R.Bottom + R.Top, False
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawGradientEx
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lhdcEx (Long)
'                              lEndColor (Long)
'                              lStartColor (Long)
'                              X (Long)
'                              Y (Long)
'                              X2 (Long)
'                              Y2 (Long)
'                              blnVertical (Long)
'!--------------------------------------------------------------------------------
Private Sub DrawGradientEx(lhdcEx As Long, ByVal lEndColor As Long, ByVal lStartColor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional blnVertical = True)

    On Error Resume Next

    'Draw a Vertical or horizontal Gradient in the current HDC
    Dim dR As Single
    Dim dG As Single
    Dim dB As Single
    Dim sR As Single
    Dim sG As Single
    Dim sB As Single
    Dim er As Single
    Dim eG As Single
    Dim eB As Single
    Dim ni As Long

    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    er = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000

    If blnVertical Then
        dR = (sR - er) / Y2
        dG = (sG - eG) / Y2
        dB = (sB - eB) / Y2

        For ni = 1 To Y2 - 1
            APILineEx lhdcEx, X, Y + ni, X2, Y + ni, RGB(er + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next

    Else
        dR = (sR - er) / X2
        dG = (sG - eG) / X2
        dB = (sB - eB) / X2

        For ni = 1 To X2 - 1
            APILineEx lhdcEx, X + ni, Y, X + ni, Y2, RGB(er + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawGradientInRectangle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lhdcEx (Long)
'                              lStartColor (Long)
'                              lEndColor (Long)
'                              R (RECT)
'                              GradientType (jcGradConst)
'                              blnDrawBorder (Boolean = False)
'                              lBorderColor (Long = vbBlack)
'                              LightCenter (Double = 2.01)
'!--------------------------------------------------------------------------------
Private Sub DrawGradientInRectangle(lhdcEx As Long, lStartColor As Long, lEndColor As Long, R As RECT, GradientType As jcGradConst, Optional ByVal blnDrawBorder As Boolean = False, Optional lBorderColor As Long = vbBlack, Optional LightCenter As _
                            Double = 2.01)

    Select Case GradientType

        Case VerticalGradient
            DrawGradientEx lhdcEx, lEndColor, lStartColor, R.Left, R.Top, R.Right + R.Left, R.Bottom, True

        Case HorizontalGradient
            DrawGradientEx lhdcEx, lEndColor, lStartColor, R.Left, R.Top, R.Right, R.Bottom + R.Top, False

        Case VCilinderGradient
            DrawGradCilinder lhdcEx, lStartColor, lEndColor, R, True, LightCenter

        Case HCilinderGradient
            DrawGradCilinder lhdcEx, lStartColor, lEndColor, R, False, LightCenter
    End Select

    If blnDrawBorder Then
        APIRectangle lhdcEx, R.Left, R.Top, R.Right, R.Bottom, lBorderColor
    End If

End Sub

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
'! Parameters  (Переменные):   New_Enabled (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let Enabled(ByRef New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    PaintFrame
    FrameEnabled m_Enabled
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub EraseRegion
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub EraseRegion()

    Dim hRgn As Long

    'Creates second region to fill with color.
    hRgn = CreateRoundRectRgn(0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, 0&)
    SetWindowRgn UserControl.hWnd, hRgn, True
    'delete our elliptical region
    DeleteObject hRgn
    UserControl.FillStyle = 0
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property FillColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get FillColor() As OLE_COLOR
    FillColor = m_FillColorIni
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property FillColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_FillColor (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let FillColor(ByRef new_FillColor As OLE_COLOR)
Attribute FillColor.VB_Description = "Returns/Sets the Fill color for TextBox and Windows style"
    m_FillColorIni = TranslateColor(new_FillColor)
    PropertyChanged "FillColor"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Font
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = PropFont
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Font
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewFont (StdFont)
'!--------------------------------------------------------------------------------
Public Property Let Font(ByVal NewFont As StdFont)
    Set Me.Font = NewFont
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Font
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewFont (StdFont)
'!--------------------------------------------------------------------------------
Public Property Set Font(ByVal NewFont As StdFont)

    Dim OldFontHandle As Long

    Set PropFont = NewFont
    Call OLEFontToLogFont(NewFont, FrameLogFont)
    OldFontHandle = FrameFontHandle
    FrameFontHandle = CreateFontIndirect(FrameLogFont)

    If UserControl.hDC <> 0 Then SendMessage UserControl.hDC, WM_SETFONT, FrameFontHandle, ByVal 1&
    If OldFontHandle <> 0 Then DeleteObject OldFontHandle
    Me.Refresh
    'PaintFrame
    UserControl.PropertyChanged "Font"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PropFont_FontChanged
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropertyName (String)
'!--------------------------------------------------------------------------------
Private Sub PropFont_FontChanged(ByVal PropertyName As String)

    Dim OldFontHandle As Long

    Call OLEFontToLogFont(PropFont, FrameLogFont)
    OldFontHandle = FrameFontHandle
    FrameFontHandle = CreateFontIndirect(FrameLogFont)

    If UserControl.hDC <> 0 Then SendMessage UserControl.hDC, WM_SETFONT, FrameFontHandle, ByVal 1&
    If OldFontHandle <> 0 Then DeleteObject OldFontHandle
    Me.Refresh
    UserControl.PropertyChanged "Font"
    'PaintFrame
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub OLEFontToLogFont
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Font (StdFont)
'                              LF (LOGFONT)
'!--------------------------------------------------------------------------------
Private Sub OLEFontToLogFont(ByVal Font As StdFont, ByRef LF As LOGFONT)

    Dim FontName As String

    With LF
        FontName = Left$(Font.Name, LF_FACESIZE)
        CopyMemory .LFFaceName(0), ByVal StrPtr(FontName), LenB(FontName)
        .LFHeight = -MulDiv(CLng(Font.Size), DPI_Y(), 72)

        If Font.Bold = True Then
            .LFWeight = FW_BOLD
        Else
            .LFWeight = FW_NORMAL
        End If

        .LFItalic = IIf(Font.Italic = True, 1, 0)
        .LFStrikeOut = IIf(Font.Strikethrough = True, 1, 0)
        .LFUnderline = IIf(Font.Underline = True, 1, 0)
        .LFQuality = DEFAULT_QUALITY
        .LFCharset = CByte(Font.Charset And &HFF)
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property FrameColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get FrameColor() As OLE_COLOR
    FrameColor = m_FrameColorIni
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property FrameColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_FrameColor (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let FrameColor(ByRef new_FrameColor As OLE_COLOR)
    m_FrameColorIni = TranslateColor(new_FrameColor)

    If m_ThemeColor = Custom Then
        jcColorBorderPic = m_FrameColor
    End If

    PropertyChanged "FrameColor"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FrameEnabled
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   blnValor (Boolean)
'!--------------------------------------------------------------------------------
Private Sub FrameEnabled(ByVal blnValor As Boolean)

    Dim C As Control

    On Error Resume Next

    For Each C In UserControl.ContainedControls

        C.Enabled = blnValor
    Next

    On Error GoTo 0

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property GradientHeaderStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get GradientHeaderStyle() As jcGradConst
    GradientHeaderStyle = m_GradientHeaderStyle
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property GradientHeaderStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_GradientHeaderStyle (jcGradConst)
'!--------------------------------------------------------------------------------
Public Property Let GradientHeaderStyle(ByRef new_GradientHeaderStyle As jcGradConst)
    m_GradientHeaderStyle = new_GradientHeaderStyle
    PropertyChanged "GradientHeaderStyle"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property AtivarResizeDoForm
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get AtivarResizeDoForm() As Boolean
    AtivarResizeDoForm = m_AtivarResizeDoForm
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property AtivarResizeDoForm
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_AtivarResizeDoForm (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let AtivarResizeDoForm(ByVal New_AtivarResizeDoForm As Boolean)
    m_AtivarResizeDoForm = New_AtivarResizeDoForm
    PropertyChanged "AtivarResizeDoForm"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MoverResponsavel
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get MoverResponsavel() As jcResp
    MoverResponsavel = m_Responsavel
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MoverResponsavel
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_Responsavel (jcResp)
'!--------------------------------------------------------------------------------
Public Property Let MoverResponsavel(ByVal New_Responsavel As jcResp)
    m_Responsavel = New_Responsavel
    PropertyChanged "MoverResponsavel"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MoverControle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get MoverControle() As Boolean
    MoverControle = m_AllowDraging
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MoverControle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_AllowDraging (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let MoverControle(ByVal New_AllowDraging As Boolean)
    m_AllowDraging = New_AllowDraging
    PropertyChanged "MoverControle"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MoverForm
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get MoverForm() As Boolean
    MoverForm = m_AllowParentDraging
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MoverForm
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_Drag (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let MoverForm(ByVal New_Drag As Boolean)
    m_AllowParentDraging = New_Drag
    PropertyChanged "MoverForm"
    Call PaintFrame
End Property

'Faz o form aumentar e diminuir o seu tamanho (Collapsar)
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Collapsar
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Collapsar() As Boolean
    Collapsar = m_Collapsar
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Collapsar
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_Drag (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let Collapsar(ByVal New_Drag As Boolean)

    Select Case m_Style

        Case Header, Panel, TextBox:
            m_Collapsar = False

        Case Else
            m_Collapsar = New_Drag
    End Select

    If m_Collapsar Then
        Label.Visible = True
    Else
        Label.Visible = False
    End If

    'PaintFrame
    PropertyChanged "Collapsar"
End Property

'Faz o form aumentar e diminuir o seu tamanho (Collapsar)
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Collapsado
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Collapsado() As Boolean
    Collapsado = m_Collapsado
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Collapsado
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_Drag (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let Collapsado(ByVal New_Drag As Boolean)

    If m_Collapsar Then

        Select Case m_Style

            Case Header, Panel, TextBox:
                m_Collapsado = False

            Case Else:
                m_Collapsado = New_Drag

                Dim iY As Integer

                Select Case m_Style

                    Case Messenger
                        iY = 9

                    Case jcGradient
                        iY = 8

                    Case XPDefault
                        iY = 6

                    Case InnerWedge
                        iY = -1

                    Case OuterWedge
                        iY = -2

                    Case Windows
                        iY = 1
                End Select

                If New_Drag Then
                    Label.Caption = "6"
                    UserControl.Height = temp_height
                Else
                    Label.Caption = "5"
                    temp_height = UserControl.Height
                    UserControl.Height = ScaleY(m_TextBoxHeight + iY, vbPixels, vbTwips)
                End If

        End Select

    Else
        m_Collapsado = False
    End If

    PropertyChanged "Collapsado"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property HeaderStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get HeaderStyle() As jcHeaderConst
    HeaderStyle = m_HeaderStyle
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property HeaderStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_HeaderStyle (jcHeaderConst)
'!--------------------------------------------------------------------------------
Public Property Let HeaderStyle(ByRef new_HeaderStyle As jcHeaderConst)
    m_HeaderStyle = new_HeaderStyle
    PropertyChanged "HeaderStyle"
    PaintFrame
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
'! Procedure   (Функция)   :   Property IconAlignment
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get IconAlignment() As IconAlignConst
    IconAlignment = m_IconAlignment
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property IconAlignment
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_IconAlignment (IconAlignConst)
'!--------------------------------------------------------------------------------
Public Property Let IconAlignment(ByRef new_IconAlignment As IconAlignConst)
    m_IconAlignment = new_IconAlignment
    PropertyChanged "IconAlignment"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property IconSize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get IconSize() As Integer
    IconSize = m_IconSize
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property IconSize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_Value (Integer)
'!--------------------------------------------------------------------------------
Public Property Let IconSize(ByVal New_Value As Integer)
    m_IconSize = New_Value
    PropertyChanged "IconSize"
    PaintFrame
End Property

'==================
' Main drawing sub
'==================
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PaintFrame
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub PaintFrame()

    Dim R_Caption        As RECT
    Dim RC               As RECT
    Dim iX               As Integer
    Dim iY               As Integer
    Dim m_caption_aux    As String
    Dim lpDrawTextParams As DRAWTEXTPARAMS

    m_Height = 3
    m_Indentation = 15
    m_Space = 6
    EraseRegion
    'Clear user control
    UserControl.Cls

    'Set caption height and width
    If LenB(m_Caption) <> 0 Then
        m_TextWidth = UserControl.TextWidth(m_Caption)
        m_TextHeight = UserControl.TextHeight(m_Caption)
        jcTextBoxCenter = m_TextHeight / 2
    Else
        jcTextBoxCenter = 0
    End If

    'Select colors according to enabled property
    If Not m_Enabled Then
        m_FrameColor = m_FrameColorDis
        m_TextBoxColor = m_TextBoxColorDis
        m_FillColor = m_FillColorDis
        jcColorTo = jcColorToDis
        jcColorFrom = jcColorFromDis
        jcColorBorderPic = jcColorBorderPicDis
    Else
        m_FrameColor = m_FrameColorIni
        m_TextBoxColor = m_TextBoxColorIni
        m_FillColor = m_FillColorIni
        jcColorTo = jcColorToIni
        jcColorFrom = jcColorFromIni
        jcColorBorderPic = jcColorBorderPicIni
    End If

    'select frame style
    Select Case m_Style

        Case XPDefault
            Draw_XPDefault R_Caption

        Case jcGradient
            Draw_jcGradient R_Caption, iY

        Case TextBox
            Draw_TextBox R_Caption, iX, iY

        Case Windows
            Draw_Windows R_Caption, iY

        Case Messenger
            Draw_Messenger R_Caption, iY

        Case InnerWedge
            Draw_InnerWedge R_Caption

        Case OuterWedge
            Draw_OuterWedge R_Caption

        Case Header
            Draw_Header R_Caption

        Case Panel
            Draw_Panel R_Caption, iY

        Case Else
            Draw_jcGradient R_Caption, iY
    End Select

    'caption and icon alignments
    If Not (m_Icon Is Nothing Or m_Style = XPDefault) Then
        If m_IconAlignment = vbLeftAligment Then
            If m_Alignment = vbLeftJustify Then
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
            ElseIf m_Alignment = vbRightJustify Then
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
            Else
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            End If

            If m_Style = TextBox Then
                iX = m_Indentation + m_Space * 2
            Else
                iX = m_Space
            End If

        ElseIf m_IconAlignment = vbRightAligment Then

            If m_Alignment = vbLeftJustify Then
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            ElseIf m_Alignment = vbRightJustify Then
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            Else
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            End If

            If m_Style = TextBox Then
                iX = UserControl.ScaleWidth - m_Space * 2 - m_IconSize - m_Indentation
            Else
                iX = UserControl.ScaleWidth - m_Space - m_IconSize
            End If
        End If
    End If

    'Draw caption
    If LenB(m_Caption) <> 0 Then
        'Set text color
        m_caption_aux = TrimWord(m_Caption, R_Caption.Right - R_Caption.Left)
        'Draw text
        UserControl.ForeColor = IIf(m_Enabled, m_TextColor, TranslateColor(TEXT_INACTIVE))
        'Set Me.Font = PropFont
        lpDrawTextParams.cbSize = Len(lpDrawTextParams)

        If m_Style = Panel Then
            CopyRect RC, R_Caption
            'DrawTextEx UserControl.hDC, m_Caption, Len(m_Caption), RC, DT_CALCRECT Or DT_WORDBREAK, ByVal 0&
            DrawTextExW UserControl.hDC, StrPtr(m_Caption & vbNullChar), -1, RC, DT_CALCRECT Or DT_WORDBREAK, lpDrawTextParams
            OffsetRect RC, (R_Caption.Right - RC.Right) \ 2, (R_Caption.Bottom - RC.Bottom) \ 2
            'DrawTextEx UserControl.hDC, m_Caption, Len(m_Caption), RC, jcTextDrawParams, ByVal 0&
            DrawTextExW UserControl.hDC, StrPtr(m_Caption & vbNullChar), -1, RC, jcTextDrawParams, lpDrawTextParams
        Else
            'DrawTextEx UserControl.hDC, m_caption_aux, Len(m_caption_aux), R_Caption, jcTextDrawParams, ByVal 0&
            DrawTextExW UserControl.hDC, StrPtr(m_caption_aux & vbNullChar), -1, R_Caption, jcTextDrawParams, lpDrawTextParams
        End If
    End If

    'draw picture
    If Not (m_Icon Is Nothing Or m_Style = XPDefault Or m_Style = InnerWedge Or m_Style = OuterWedge) Then
        If m_Style = Messenger Then
            If iY < m_Height * 2 + 2 Then
                iY = m_Height * 2 + 2
            End If

        ElseIf m_Style = jcGradient Then

            If iY < m_Height + 2 Then
                iY = m_Height + 2
            End If

        Else

            If iY < 0 Then
                iY = m_Space / 2
            End If
        End If

        If m_Enabled Then
            UserControl.PaintPicture m_Icon, iX, iY, m_IconSize, m_IconSize
            'TransBlt UserControl.hDC,iX,iY, m_IconSize, m_IconSize, m_Icon, vbBlack, , , False, False
        Else
            TransBlt UserControl.hDC, iX, iY, m_IconSize, m_IconSize, m_Icon, vbBlack, , , True, False
        End If
    End If

    Select Case m_Style

        Case Messenger
            iY = 5

        Case jcGradient
            iY = 8

        Case XPDefault
            iY = 6

        Case InnerWedge
            iY = 11

        Case OuterWedge
            iY = 11

        Case Windows
            iY = 11
    End Select

    Label.Move UserControl.ScaleWidth - 30, CInt(ScaleY((ScaleY(m_TextBoxHeight, vbPixels, vbTwips) - Label.Height) / 2, vbTwips, vbPixels)) - iY
    Set UserControl.Picture = UserControl.Image
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PaintShpInBar
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   iColorA (Long)
'                              iColorB (Long)
'                              m_Height (Long)
'!--------------------------------------------------------------------------------
Private Sub PaintShpInBar(iColorA As Long, iColorB As Long, ByVal m_Height As Long)

    Dim i            As Integer
    Dim x_left       As Integer
    Dim y_top        As Integer
    Dim SpaceBtwnShp As Integer
    Dim NumShp       As Integer
    Dim RectHeight   As Long
    Dim RectWidth    As Long
    Dim R            As RECT

    SpaceBtwnShp = 2
    'space between shapes
    NumShp = 9
    'number of points
    RectHeight = 2
    'shape height
    RectWidth = 2
    'shape width
    'x and y shape  coordinates
    x_left = (UserControl.ScaleWidth - NumShp * RectWidth - (NumShp - 1) * SpaceBtwnShp) / 2
    y_top = (m_Height - RectHeight) / 2

    For i = 0 To NumShp - 1
        SetRect R, x_left + i * SpaceBtwnShp + i * RectWidth + 1, y_top + 1, 1, 1
        APIRectangle UserControl.hDC, R.Left, R.Top, R.Right, R.Bottom, iColorA
        SetRect R, x_left + i * SpaceBtwnShp + i * RectWidth, y_top, 1, 1
        APIRectangle UserControl.hDC, R.Left, R.Top, R.Right, R.Bottom, iColorB
    Next

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Picture
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Picture() As StdPicture
    Set Picture = m_Icon
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Picture
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_Picture (StdPicture)
'!--------------------------------------------------------------------------------
Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Icon = New_Picture
    PropertyChanged "Picture"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property RoundedCorner
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get RoundedCorner() As Boolean
    RoundedCorner = m_RoundedCorner
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property RoundedCorner
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_RoundedCorner (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let RoundedCorner(ByRef new_RoundedCorner As Boolean)
    m_RoundedCorner = new_RoundedCorner
    PropertyChanged "RoundedCorner"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property RoundedCornerTxtBox
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get RoundedCornerTxtBox() As Boolean
    RoundedCornerTxtBox = m_RoundedCornerTxtBox
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property RoundedCornerTxtBox
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_RoundedCornerTxtBox (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let RoundedCornerTxtBox(ByRef new_RoundedCornerTxtBox As Boolean)
    m_RoundedCornerTxtBox = new_RoundedCornerTxtBox
    PropertyChanged "RoundedCornerTxtBox"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetDefault
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SetDefault()

    Select Case m_Style

        Case XPDefault
            m_TextColor = &HCF3603
            m_FrameColorIni = RGB(195, 195, 195)
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams

        Case jcGradient
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_ThemeColor = Blue
            SetjcTextDrawParams

        Case TextBox
            m_TextColor = vbBlack
            m_FrameColorIni = &H6A6A6A
            m_TextBoxColorIni = &HB0EFF0
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_RoundedCornerTxtBox = True
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams

        Case Windows
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = &HB0EFF0
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_RoundedCorner = True
            m_RoundedCornerTxtBox = False
            m_FillColorIni = &HE0FFFF
            m_GradientHeaderStyle = HorizontalGradient
            m_HeaderStyle = TxtBoxColor
            SetjcTextDrawParams

        Case Messenger
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_ThemeColor = Blue
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = TxtBoxColor
            SetjcTextDrawParams

        Case InnerWedge
            m_TextColor = vbWhite
            m_FrameColorIni = 192
            m_TextBoxColorIni = 192
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams

        Case OuterWedge
            m_TextColor = vbWhite
            m_FrameColorIni = 10878976
            m_TextBoxColorIni = 10878976
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams

        Case Header
            m_TextColor = &HCF3603
            m_FrameColorIni = RGB(195, 195, 195)
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams

        Case Panel
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_ThemeColor = Blue
            m_GradientHeaderStyle = VCilinderGradient
            m_HeaderStyle = Gradient
            SetjcTextDrawParams
    End Select

End Sub

'==========================================================================
' Functions and subroutines
'==========================================================================
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetDefaultThemeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   ThemeType (Long)
'!--------------------------------------------------------------------------------
Private Sub SetDefaultThemeColor(ByVal ThemeType As Long)

    Select Case ThemeType

        Case 0
            '"NormalColor" - "Blue"
            jcColorFromIni = RGB(129, 169, 226)
            jcColorToIni = RGB(221, 236, 254)
            jcColorBorderPicIni = RGB(0, 0, 128)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(0, 0, 0)
                m_BackColor = RGB(192, 212, 240)
                m_FillColorIni = RGB(192, 212, 240)
            End If

        Case 1
            '"Metallic" - "Silver"
            jcColorFromIni = RGB(153, 151, 180)
            jcColorToIni = RGB(244, 244, 251)
            jcColorBorderPicIni = RGB(75, 75, 111)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(0, 0, 0)
                m_BackColor = RGB(204, 203, 218)
                m_FillColorIni = RGB(204, 203, 218)
            End If

        Case 2
            '"HomeStead" - "Olive"
            jcColorFromIni = RGB(181, 197, 143)
            jcColorToIni = RGB(247, 249, 225)
            jcColorBorderPicIni = RGB(63, 93, 56)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(0, 0, 0)
                m_BackColor = RGB(224, 224, 213)
                m_FillColorIni = RGB(224, 224, 213)
            End If

        Case 3
            '"Visual2005"
            jcColorFromIni = RGB(194, 194, 171)
            jcColorToIni = RGB(248, 248, 242)
            jcColorBorderPicIni = RGB(145, 145, 115)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(0, 0, 0)
                m_BackColor = RGB(224, 224, 213)
                m_FillColorIni = RGB(224, 224, 213)
            End If

        Case 4
            '"Norton2004"
            jcColorFromIni = RGB(217, 172, 1)
            jcColorToIni = RGB(255, 239, 165)
            jcColorBorderPicIni = RGB(117, 91, 30)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(0, 0, 0)
                m_BackColor = RGB(236, 214, 128)
                m_FillColorIni = RGB(236, 214, 128)
            End If

        Case 5
            'Custom
            jcColorFromIni = m_ColorFrom
            jcColorToIni = m_ColorTo
            jcColorBorderPicIni = m_FrameColor

        Case 6
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = Gradient
            jcColorFromIni = RGB(137, 170, 224)
            jcColorToIni = RGB(7, 33, 100)
            jcColorBorderPicIni = RGB(100, 144, 88)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(215, 230, 251)
                m_FillColorIni = RGB(142, 179, 231)
                '=
                m_BackColor = RGB(142, 179, 231)
                '=
            End If

        Case 7
            'xThemeGreen
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = Gradient
            jcColorFromIni = RGB(228, 235, 200)
            jcColorToIni = RGB(175, 194, 142)
            jcColorBorderPicIni = RGB(100, 144, 88)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(100, 144, 88)
                m_FillColorIni = RGB(233, 244, 207)
                '=
                m_BackColor = RGB(233, 244, 207)
                '=
            End If

        Case 8
            'xThemeOffice2003Style2
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = Gradient
            jcColorFromIni = RGB(249, 249, 255)
            jcColorToIni = RGB(159, 157, 185)
            jcColorBorderPicIni = RGB(124, 124, 148)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(110, 109, 143)
                m_FillColorIni = RGB(253, 250, 255)
                '=
                m_BackColor = RGB(253, 250, 255)
                '=
            End If

        Case 9
            'xThemeMetallic
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = Gradient
            jcColorFromIni = RGB(219, 220, 232)
            jcColorToIni = RGB(149, 147, 177)
            jcColorBorderPicIni = RGB(119, 118, 151)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(119, 118, 151)
                m_FillColorIni = RGB(232, 232, 232)
                '=
                m_BackColor = RGB(232, 232, 232)
                '=
            End If

        Case 10
            'xThemeOrange
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = Gradient
            jcColorFromIni = RGB(255, 122, 0)
            jcColorToIni = RGB(130, 0, 0)
            jcColorBorderPicIni = RGB(139, 0, 0)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(255, 222, 173)
                m_FillColorIni = RGB(255, 222, 173)
                '=
                m_BackColor = RGB(255, 222, 173)
                '=
            End If

        Case 11
            'xThemeTurquoise
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = Gradient
            jcColorFromIni = RGB(72, 209, 204)
            jcColorToIni = RGB(43, 103, 109)
            jcColorBorderPicIni = RGB(65, 131, 111)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(233, 250, 248)
                m_FillColorIni = RGB(224, 255, 255)
                '=
                m_BackColor = RGB(224, 255, 255)
                '=
            End If

        Case 12
            'xThemeGray
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = Gradient
            jcColorFromIni = RGB(192, 192, 192)
            jcColorToIni = RGB(51, 51, 51)
            jcColorBorderPicIni = RGB(51, 51, 51)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(235, 235, 235)
                m_FillColorIni = RGB(235, 235, 235)
                '=
                m_BackColor = RGB(235, 235, 235)
                '=
            End If

        Case 13
            'xThemeDarkBlue2
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = Gradient
            jcColorFromIni = RGB(81, 128, 208)
            jcColorToIni = dBlendColor(RGB(11, 63, 153), vbBlack, 230)
            jcColorBorderPicIni = RGB(0, 45, 150)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = vbRed
                m_FillColorIni = RGB(142, 179, 231)
                '=
                m_BackColor = RGB(142, 179, 231)
                '=
            End If

        Case 14
            'xThemeMoney
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = Gradient
            jcColorFromIni = RGB(160, 160, 160)
            jcColorToIni = dBlendColor(RGB(90, 90, 90), vbBlack, 230)
            jcColorBorderPicIni = RGB(68, 68, 68)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = vbWhite
                m_FillColorIni = RGB(112, 112, 112)
                '=
                m_BackColor = RGB(112, 112, 112)
                '=
            End If

        Case 15
            'xThemeOffice2003Style1
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = Gradient
            jcColorFromIni = RGB(209, 227, 251)
            jcColorToIni = RGB(106, 140, 203)
            jcColorBorderPicIni = RGB(0, 0, 128)

            If (Style = jcGradient) Or (Style = Messenger) Or (Style = Windows) Or (Style = TextBox) Then
                m_TextColor = RGB(110, 109, 143)
                m_FillColorIni = RGB(255, 255, 255)
                '=
                m_BackColor = RGB(255, 255, 255)
                '=
            End If

        Case Else
            jcColorFromIni = RGB(153, 151, 180)
            jcColorToIni = RGB(244, 244, 251)
            jcColorBorderPicIni = RGB(75, 75, 111)
    End Select

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property dBlendColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   oColorFrom (OLE_COLOR)
'                              oColorTo (OLE_COLOR)
'                              Alpha (Long = 128)
'!--------------------------------------------------------------------------------
Public Property Get dBlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long

    Dim lSrcR  As Long
    Dim lSrcG  As Long
    Dim lSrcB  As Long
    Dim lDstR  As Long
    Dim lDstG  As Long
    Dim lDstB  As Long
    Dim lCFrom As Long
    Dim lCTo   As Long

    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    dBlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetDisabledColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SetDisabledColor()
    m_FrameColorDis = TranslateColor(m_Border_Inactive)
    m_TextBoxColorDis = TranslateColor(m_BtnFace)
    '_Inactive)
    m_FillColorDis = TranslateColor(Ambient.BackColor)
    jcColorToDis = TranslateColor(m_BtnFace_Inactive)
    jcColorFromDis = TranslateColor(m_BtnFace_Inactive)
    jcColorBorderPicDis = TranslateColor(m_Border_Inactive)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetjcTextDrawParams
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SetjcTextDrawParams()

    'Set text draw params using m_Alignment
    If m_Style = Panel Then
        If m_Alignment = vbLeftJustify Then
            jcTextDrawParams = DT_LEFT Or DT_WORDBREAK Or DT_VCENTER
        ElseIf m_Alignment = vbRightJustify Then
            jcTextDrawParams = DT_RIGHT Or DT_WORDBREAK Or DT_VCENTER
        Else
            jcTextDrawParams = DT_CENTER Or DT_WORDBREAK Or DT_VCENTER
        End If

    Else

        If m_Alignment = vbLeftJustify Then
            jcTextDrawParams = DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
        ElseIf m_Alignment = vbRightJustify Then
            jcTextDrawParams = DT_RIGHT Or DT_SINGLELINE Or DT_VCENTER
        Else
            jcTextDrawParams = DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
        End If
    End If

    If Ambient.RightToLeft = True Then jcTextDrawParams = jcTextDrawParams Or WS_EX_RTLREADING
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Style
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Style() As jcStyleConst

    If (m_Style = Header) Or (m_Style = Panel) Or (m_Style = TextBox) Then
        m_Collapsar = False
        Label.Visible = False
    End If

    Style = m_Style
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Style
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_Style (jcStyleConst)
'!--------------------------------------------------------------------------------
Public Property Let Style(ByRef new_Style As jcStyleConst)
    m_Style = new_Style
    PropertyChanged "Style"

    If (new_Style = Header) Or (new_Style = Panel) Or (new_Style = TextBox) Then m_Collapsar = False
    Label.Visible = False
    SetDefault
    ' m_ThemeColor
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property TextBoxColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get TextBoxColor() As OLE_COLOR
    TextBoxColor = m_TextBoxColorIni
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property TextBoxColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_TextBoxColor (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let TextBoxColor(ByRef new_TextBoxColor As OLE_COLOR)
    m_TextBoxColorIni = TranslateColor(new_TextBoxColor)
    PropertyChanged "TextBoxColor"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property TextBoxHeight
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get TextBoxHeight() As Long
    TextBoxHeight = m_TextBoxHeight
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property TextBoxHeight
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_TextBoxHeight (Long)
'!--------------------------------------------------------------------------------
Public Property Let TextBoxHeight(ByRef new_TextBoxHeight As Long)
    m_TextBoxHeight = new_TextBoxHeight
    PropertyChanged "TextBoxHeight"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property TextColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property TextColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_TextColor (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let TextColor(ByRef new_TextColor As OLE_COLOR)
    m_TextColor = TranslateColor(new_TextColor)
    PropertyChanged "TextColor"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ThemeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ThemeColor() As jcThemeConst
    ThemeColor = m_ThemeColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ThemeColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   vData (jcThemeConst)
'!--------------------------------------------------------------------------------
Public Property Let ThemeColor(ByVal vData As jcThemeConst)

    If m_ThemeColor <> vData Then
        m_ThemeColor = vData
        SetDefaultThemeColor m_ThemeColor
        PaintFrame
        PropertyChanged "ThemeColor"
    End If

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub TransBlt
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   DstDC (Long)
'                              DstX (Long)
'                              DstY (Long)
'                              DstW (Long)
'                              DstH (Long)
'                              SrcPic (StdPicture)
'                              TransColor (Long = -1)
'                              BrushColor (Long = -1)
'                              MonoMask (Boolean = False)
'                              isGreyscale (Boolean = False)
'                              XPBlend (Boolean = False)
'!--------------------------------------------------------------------------------
Private Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal TransColor As Long = -1, Optional ByVal BrushColor As Long = -1, Optional ByVal _
                            MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False, Optional ByVal XPBlend As Boolean = False)

    Dim B        As Long
    Dim H        As Long
    Dim F        As Long
    Dim i        As Long
    Dim newW     As Long
    Dim TmpDC    As Long
    Dim TmpBmp   As Long
    Dim TmpObj   As Long
    Dim Sr2DC    As Long
    Dim Sr2Bmp   As Long
    Dim Sr2Obj   As Long
    Dim Data1()  As RGB
    Dim Data2()  As RGB
    Dim Info     As BITMAPINFO
    Dim BrushRGB As RGB
    Dim gCol     As Long
    Dim SrcDC    As Long
    Dim tObj     As Long
    Dim hBrush   As Long

    If Not DstW = 0 Or DstH = 0 Then
        SrcDC = CreateCompatibleDC(hDC)

        If DstW < 0 Then
            DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
        End If

        If DstH < 0 Then
            DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)
        End If

        If SrcPic.Type = 1 Then
            'check if it's an icon or a bitmap
            tObj = SelectObject(SrcDC, SrcPic)
        Else
            tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
            hBrush = CreateSolidBrush(TransColor)
            'MaskColor)
            DrawIconEx SrcDC, 0, 0, SrcPic.Handle, DstW, DstH, 0, hBrush, &H1 Or &H2
            DeleteObject hBrush
        End If

        TmpDC = CreateCompatibleDC(SrcDC)
        Sr2DC = CreateCompatibleDC(SrcDC)
        TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
        Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
        TmpObj = SelectObject(TmpDC, TmpBmp)
        Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)

        ReDim Data1(DstW * DstH * 3 - 1) As RGB
        ReDim Data2(UBound(Data1)) As RGB

        With Info.bmiHeader
            .biSize = Len(Info.bmiHeader)
            .biWidth = DstW
            .biHeight = DstH
            .biPlanes = 1
            .biBitCount = 24
        End With

        'INFO.BMIHEADER
        BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
        BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
        GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
        GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0

        If BrushColor > 0 Then

            With BrushRGB
                .Blue = (BrushColor \ &H10000) Mod &H100
                .Green = (BrushColor \ &H100) Mod &H100
                .Red = BrushColor And &HFF
            End With

        End If

        useMask = True

        If Not useMask Then
            TransColor = -1
        End If

        newW = DstW - 1

        For H = 0 To DstH - 1
            F = H * DstW

            For B = 0 To newW
                i = F + B

                If GetNearestColor(hDC, CLng(Data2(i).Red) + 256& * Data2(i).Green + 65536 * Data2(i).Blue) <> TransColor Then

                    With Data1(i)

                        If BrushColor > -1 Then
                            If MonoMask Then
                                If (CLng(Data2(i).Red) + Data2(i).Green + Data2(i).Blue) <= 384 Then
                                    Data1(i) = BrushRGB
                                End If

                            Else
                                Data1(i) = BrushRGB
                            End If

                        Else

                            If isGreyscale Then
                                gCol = CLng(Data2(i).Red * 0.3) + Data2(i).Green * 0.59 + Data2(i).Blue * 0.11
                                .Red = gCol
                                .Green = gCol
                                .Blue = gCol
                            Else

                                If XPBlend Then
                                    .Red = (CLng(.Red) + Data2(i).Red * 2) \ 3
                                    .Green = (CLng(.Green) + Data2(i).Green * 2) \ 3
                                    .Blue = (CLng(.Blue) + Data2(i).Blue * 2) \ 3
                                Else
                                    Data1(i) = Data2(i)
                                End If
                            End If
                        End If

                    End With

                End If

            Next
        Next

        SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0
        Erase Data1, Data2
        DeleteObject SelectObject(TmpDC, TmpObj)
        DeleteObject SelectObject(Sr2DC, Sr2Obj)

        If SrcPic.Type = 3 Then
            DeleteObject SelectObject(SrcDC, tObj)
        End If

        DeleteDC TmpDC
        DeleteDC Sr2DC
        DeleteObject tObj
        DeleteDC SrcDC
    End If

End Sub

'System color code to long rgb
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function TranslateColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lColor (Long)
'!--------------------------------------------------------------------------------
Private Function TranslateColor(ByVal lColor As Long) As Long

    If OleTranslateColor(lColor, 0, TranslateColor) Then
        TranslateColor = -1
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function TrimWord
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strCaption (String)
'                              lngWidth (Long)
'!--------------------------------------------------------------------------------
Private Function TrimWord(strCaption As String, lngWidth As Long) As String

    Dim lngLenOfText As Long

    TrimWord = strCaption

    If TextWidth(strCaption) > lngWidth Then
        lngLenOfText = Len(strCaption)

        Do Until TextWidth(TrimWord & "...") <= lngWidth Or lngLenOfText = 0
            lngLenOfText = lngLenOfText - 1
            TrimWord = Left$(TrimWord, lngLenOfText)
        Loop

        If lngLenOfText = 0 Then
            TrimWord = Empty
        Else
            TrimWord = TrimWord & "..."
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property TxtBoxShadow
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get TxtBoxShadow() As jcShadowConst
    TxtBoxShadow = m_TxtBoxShadow
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property TxtBoxShadow
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_TxtBoxShadow (jcShadowConst)
'!--------------------------------------------------------------------------------
Public Property Let TxtBoxShadow(ByRef new_TxtBoxShadow As jcShadowConst)
    m_TxtBoxShadow = new_TxtBoxShadow
    PropertyChanged "TxtBoxShadow"
    PaintFrame
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Initialize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Initialize()
    m_bIsWinXpOrLater = IsWinXPOrLater
    m_IconSize = 16
    m_ColorFrom = 10395391
    m_ColorTo = 15790335
    m_TxtBoxShadow = [No shadow]
    m_ThemeColor = Blue
    m_Enabled = True
    SetDefaultThemeColor m_ThemeColor
    m_TextBoxHeight = 22
    m_Alignment = vbCenter
    m_IconAlignment = vbLeftAligment
End Sub

'==========================================================================
' Init, Initialize, Read & Write UserControl
'==========================================================================
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_InitProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()

    With Ambient
        Set PropFont = .Font
        Set UserControl.Font = PropFont
        m_Caption = .DisplayName
        m_BackColor = TranslateColor(.BackColor)
        m_FillColorIni = TranslateColor(.BackColor)
    End With

    m_Responsavel = m_def_Responsavel
    m_AllowDraging = m_def_AllowDraging
    m_AtivarResizeDoForm = m_def_AtivarResizeDoForm
    m_AllowParentDraging = False
    m_Collapsar = m_def_Collapsar
    m_Collapsado = False
    m_RoundedCorner = True
    m_RoundedCornerTxtBox = False
    m_Style = jcGradient
    m_ThemeColor = Blue
    m_TextColor = TranslateColor(vbBlack)
    m_FrameColorIni = TranslateColor(vbBlack)
    m_TextBoxColorIni = TranslateColor(vbWhite)
    m_TxtBoxShadow = [No shadow]
    m_TextBoxHeight = 22
    m_HeaderStyle = Gradient
    m_GradientHeaderStyle = VerticalGradient
    SetjcTextDrawParams
End Sub

'==========================================================================
' Usercontrol events
'==========================================================================
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_MouseMove
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_ReadProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropBag (PropertyBag)
'!--------------------------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set PropFont = .ReadProperty("Font", Ambient.Font)
        Set UserControl.Font = PropFont
        m_FillColorIni = .ReadProperty("FillColor", Ambient.BackColor)
        m_TextBoxColorIni = .ReadProperty("TextBoxColor", vbWhite)
        m_AtivarResizeDoForm = .ReadProperty("AtivarResizeDoForm", m_def_AtivarResizeDoForm)
        m_Responsavel = .ReadProperty("MoverResponsavel", m_def_Responsavel)
        m_AllowDraging = .ReadProperty("MoverControle", m_def_AllowDraging)
        m_Collapsar = .ReadProperty("Collapsar", m_def_Collapsar)
        m_Collapsado = .ReadProperty("Collapsado", False)
        m_AllowParentDraging = PropBag.ReadProperty("MoverForm", False)
        m_TxtBoxShadow = .ReadProperty("TxtBoxShadow", [No shadow])
        m_Style = .ReadProperty("Style", jcGradient)
        m_RoundedCorner = .ReadProperty("RoundedCorner", True)
        m_Enabled = .ReadProperty("Enabled", True)
        m_RoundedCornerTxtBox = .ReadProperty("RoundedCornerTxtBox", False)
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_TextBoxHeight = .ReadProperty("TextBoxHeight", 22)
        m_TextColor = .ReadProperty("TextColor", vbBlack)
        m_Alignment = .ReadProperty("Alignment", vbCenter)
        m_IconAlignment = .ReadProperty("IconAlignment", vbLeftAligment)
        m_FrameColorIni = .ReadProperty("FrameColor", vbBlack)
        Set m_Icon = .ReadProperty("Picture", Nothing)
        m_IconSize = .ReadProperty("IconSize", 16)
        m_ThemeColor = .ReadProperty("ThemeColor", Blue)
        m_ColorFrom = .ReadProperty("ColorFrom", 10395391)
        m_ColorTo = .ReadProperty("ColorTo", 15790335)
        m_HeaderStyle = .ReadProperty("HeaderStyle", TxtBoxColor)
        m_GradientHeaderStyle = .ReadProperty("GradientHeaderStyle", VerticalGradient)
        'Add properties
        UserControl.BackColor = TranslateColor(m_BackColor)
        SetjcTextDrawParams
        m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
    End With

    If m_Collapsar Then

        Select Case m_Style

            Case Header, Panel, TextBox
                Label.Visible = False

            Case Else
                Label.Visible = True
        End Select

    End If

    SetDefaultThemeColor m_ThemeColor
    SetDisabledColor
    'Paint control
    PaintFrame

    If m_AtivarResizeDoForm Then
        If Ambient.UserMode Then
            Set frm = Parent
            SetWindowLong frm.hWnd, GWL_STYLE, GetWindowLong(frm.hWnd, GWL_STYLE) And Not (WS_CAPTION)
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Resize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Resize()

    On Error Resume Next

    If (m_Collapsar = False) Then
        If UserControl.Width < 700 Then
            UserControl.Width = 700
        End If

        If UserControl.Height < 400 Then
            UserControl.Height = 400
        End If
    End If

    PaintFrame
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Terminate
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Terminate()

    On Error Resume Next

    'Clean up Font (StdFont)
    Set PropFont = Nothing
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_WriteProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropBag (PropertyBag)
'!--------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Font", PropFont, Ambient.Font
        .WriteProperty "FrameColor", m_FrameColorIni, vbBlack
        .WriteProperty "BackColor", m_BackColor, Ambient.BackColor
        .WriteProperty "FillColor", m_FillColorIni, Ambient.BackColor
        .WriteProperty "TextBoxColor", m_TextBoxColorIni, vbWhite
        .WriteProperty "MoverForm", m_AllowParentDraging, False
        .WriteProperty "MoverControle", m_AllowDraging, m_def_AllowDraging
        .WriteProperty "AtivarResizeDoForm", m_AtivarResizeDoForm, m_def_AtivarResizeDoForm
        .WriteProperty "MoverResponsavel", m_Responsavel, m_def_Responsavel
        .WriteProperty "Collapsar", m_Collapsar, m_def_Collapsar
        .WriteProperty "Collapsado", m_Collapsado, False
        .WriteProperty "TxtBoxShadow", m_TxtBoxShadow, [No shadow]
        .WriteProperty "Style", m_Style, jcGradient
        .WriteProperty "RoundedCorner", m_RoundedCorner, True
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "RoundedCornerTxtBox", m_RoundedCornerTxtBox, False
        .WriteProperty "Caption", m_Caption, Ambient.DisplayName
        .WriteProperty "TextBoxHeight", m_TextBoxHeight, 22
        .WriteProperty "TextColor", m_TextColor, vbBlack
        .WriteProperty "Alignment", m_Alignment, vbCenter
        .WriteProperty "IconAlignment", m_IconAlignment, vbLeftAligment
        .WriteProperty "Picture", m_Icon, Nothing
        .WriteProperty "IconSize", m_IconSize, 16
        .WriteProperty "ThemeColor", m_ThemeColor, Blue
        .WriteProperty "ColorFrom", m_ColorFrom, 10395391
        .WriteProperty "ColorTo", m_ColorTo, 15790335
        .WriteProperty "HeaderStyle", m_HeaderStyle, TxtBoxColor
        .WriteProperty "GradientHeaderStyle", m_GradientHeaderStyle, VerticalGradient
    End With

End Sub

'Eventos
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function PanelOrTitle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Y (Single)
'!--------------------------------------------------------------------------------
Function PanelOrTitle(Y As Single) As m_PanelArea

    If (Y <= 0) Or (Y < m_TextBoxHeight) Then
        PanelOrTitle = xTitle
    Else
        PanelOrTitle = xPanel
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Label_MouseUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub Label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent CollapseClick(Button)

    If (Button = 1) Then
        If m_Collapsar Then
            Collapsado = Not m_Collapsado
        End If

        PaintFrame
    End If

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
    RaiseEvent MouseDown(Button, Shift, X, Y, PanelOrTitle(Y))

    If (Button <> vbLeftButton) Then

        Exit Sub

    End If

    Select Case PanelOrTitle(Y)

        Case xTitle
            RaiseEvent TileClick

        Case xPanel
            RaiseEvent PanelClick
    End Select

    Dim iHwnd As Long

    If m_AllowDraging Then
        If m_AllowParentDraging Then
            If (m_Responsavel = jcAmbos) Or (m_Responsavel = PanelOrTitle(Y)) Then
                iHwnd = UserControl.Parent.hWnd
                jcTransp 70
            End If

        Else
            iHwnd = UserControl.hWnd
        End If

        Call ReleaseCapture
        Call SendMessage(iHwnd, &HA1, 2, 0&)

        If m_AllowParentDraging Then jcTransp 255
    End If

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
    RaiseEvent MouseUp(Button, Shift, X, Y, PanelOrTitle(Y))
End Sub

'FORM TRANSPARENTE
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub jcTransp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   TranslucenceLevel (Byte)
'!--------------------------------------------------------------------------------
Private Sub jcTransp(TranslucenceLevel As Byte)

    If m_bIsWinXpOrLater Then
        SetWindowLong UserControl.Parent.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
        SetLayeredWindowAttributes UserControl.Parent.hWnd, 0, TranslucenceLevel, LWA_ALPHA
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Refresh
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
    UserControl.Refresh

    If UserControl.hDC <> 0 Then RedrawWindow UserControl.hDC, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub
