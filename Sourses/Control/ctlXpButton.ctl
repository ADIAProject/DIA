VERSION 5.00
Begin VB.UserControl ctlXpButton 
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000015&
   HasDC           =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   111
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   ToolboxBitmap   =   "ctlXpButton.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   1395
      ScaleHeight     =   825
      ScaleWidth      =   1245
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "mnuGeneral"
      Visible         =   0   'False
      Begin VB.Menu mnuMenu 
         Caption         =   "#"
         Index           =   0
      End
   End
End
Attribute VB_Name = "ctlXpButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

' ==========================================================
' Copyright © 2003 Alexander Drovosekov (apexsun@narod.ru) '
' Visit apexsun.narod.ru                                   '
' ==========================================================
Private UserScaleW As Long
Private UserScaleH As Long

Public Enum XBPicturePosition
    gbTOP = 0
    gbLEFT = 1
    gbRIGHT = 2
    gbBOTTOM = 3
End Enum

#If False Then

    Private gbTOP, gbLEFT, gbRIGHT, gbBOTTOM
#End If

Public Enum XBButtonStyle
    gbStandard = 0
    gbFlat = 1
    gbWinXP = 3
End Enum

#If False Then

    Private gbStandard, gbFlat, gbWinXP
#End If

Private mvarClientRect     As RECT
Private mvarPictureRect    As RECT
Private mvarCaptionRect    As RECT
Private mvarOrgRect        As RECT
Private g_FocusRect        As RECT
Private alan               As RECT
Private m_Picture          As Picture
Private m_PicturePosition  As XBPicturePosition
Private m_ButtonStyle      As XBButtonStyle
Private mvarDrawTextParams As DRAWTEXTPARAMS
Private m_Caption          As String
Private m_PictureWidth     As Long
Private m_PictureHeight    As Long
Private g_HasFocus         As Byte
Private g_MouseDown        As Byte
Private g_MouseIn          As Byte
Private m_ShowFocusRect    As Boolean
Private m_MaskColor        As Long
Private m_UseMaskColor     As Long
Private m_XPDefaultColors  As Boolean
Private m_MenuExist        As Boolean
Private m_CheckExist       As Boolean
Private m_XPColor_Pressed  As Long
Private m_XPColor_Hover    As Long
Private m_TextColor        As OLE_COLOR

Private Const mvarPadding  As Byte = 4

Dim dtDefTextDrawParams    As Long

Private Type RGB
    Red                                 As Byte
    Green                               As Byte
    Blue                                As Byte
End Type

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

Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINT) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

'*************************************************************
'   DRAW TEXT
'*************************************************************
Private Const DT_WORDBREAK  As Long = &H10
Private Const DT_CENTER     As Long = &H1
Private Const DT_VCENTER    As Long = &H4
Private Const DT_CALCRECT   As Long = &H400
Private Const DT_RTLREADING As Long = &H20000

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
Private ButtonFontHandle      As Long
Private ButtonLogFont         As LOGFONT
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
'   events
'*************************************************************
Public Event Click()
Public Event ClickMenu(mnuIndex As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseIn(Shift As Integer)
Public Event MouseOut(Shift As Integer)

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub AddMenu
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sCaption (String)
'!--------------------------------------------------------------------------------
Public Sub AddMenu(ByVal sCaption As String)

    Dim iCount As Integer

    On Error Resume Next

    'проверяем количество меню
    iCount = mnuMenu.Count
    'загружаем данные и показываем меню
    mnuMenu(iCount - 1).Caption = sCaption
    mnuMenu(iCount - 1).Visible = True
    'загрузка следующего меню, но невидимая
    Load mnuMenu(iCount)
    mnuMenu(iCount).Visible = False

    On Error GoTo 0

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   new_BackColor (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let BackColor(ByVal new_BackColor As OLE_COLOR)
    UserControl.BackColor = new_BackColor
    PropertyChanged "BackColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ButtonStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ButtonStyle() As XBButtonStyle
    ButtonStyle = m_ButtonStyle
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ButtonStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_ButtonStyle (XBButtonStyle)
'!--------------------------------------------------------------------------------
Public Property Let ButtonStyle(ByVal New_ButtonStyle As XBButtonStyle)
    m_ButtonStyle = New_ButtonStyle
    Refresh
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CalcRECTs
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub CalcRECTs()

    Dim picWidth  As Long
    Dim picHeight As Long
    Dim capWidth  As Long
    Dim capHeight As Long

    With alan
        .Left = 0
        .Top = 0
        .Right = UserScaleW - 1
        .Bottom = UserScaleH - 1
    End With

    With mvarClientRect
        .Left = alan.Left + mvarPadding
        .Top = alan.Top + mvarPadding
        .Right = alan.Right - mvarPadding + 1
        .Bottom = alan.Bottom - mvarPadding + 1
    End With

    If LenB(m_Caption) = 0 Then

        With mvarPictureRect
            .Left = (((mvarClientRect.Right - mvarClientRect.Left) - m_PictureWidth) * 0.5) + mvarClientRect.Left
            .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - m_PictureHeight) * 0.5) + mvarClientRect.Top
            .Right = .Left + m_PictureWidth
            .Bottom = .Top + m_PictureHeight
        End With

    Else

        With mvarCaptionRect
            .Left = mvarClientRect.Left
            .Top = mvarClientRect.Top

            If m_Picture Is Nothing Then
                .Right = mvarClientRect.Right
            Else

                If m_CheckExist Then
                    .Right = mvarClientRect.Right - m_PictureWidth - 13
                Else
                    .Right = mvarClientRect.Right - m_PictureWidth
                End If
            End If

            .Bottom = mvarClientRect.Bottom
        End With

        CalculateCaptionRect

        If m_Picture Is Nothing Then

            Exit Sub

        End If

        picWidth = m_PictureWidth
        picHeight = m_PictureHeight

        With mvarCaptionRect
            capWidth = .Right - .Left
            capHeight = .Bottom - .Top
        End With

        If m_PicturePosition = gbLEFT Then

            With mvarPictureRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - picHeight) * 0.5) + mvarClientRect.Top

                '.Left = (((mvarClientRect.Right - mvarClientRect.Left) - (picWidth + mvarPadding + capWidth)) * 0.5) + mvarClientRect.Left
                If m_CheckExist Then
                    .Left = 18
                Else
                    .Left = 3
                End If

                .Bottom = .Top + picHeight
                .Right = .Left + picWidth
            End With

            With mvarCaptionRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - capHeight) * 0.5) + mvarClientRect.Top
                .Left = mvarPictureRect.Right + mvarPadding
                .Bottom = .Top + capHeight
                .Right = .Left + capWidth
            End With

        ElseIf m_PicturePosition = gbRIGHT Then

            With mvarCaptionRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - capHeight) * 0.5) + mvarClientRect.Top
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - (picWidth + mvarPadding + capWidth)) * 0.5) + mvarClientRect.Left
                .Bottom = .Top + capHeight
                .Right = .Left + capWidth
            End With

            With mvarPictureRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - picHeight) * 0.5) + mvarClientRect.Top
                .Left = mvarCaptionRect.Right + mvarPadding
                .Bottom = .Top + picHeight
                .Right = .Left + picWidth
            End With

        ElseIf m_PicturePosition = gbTOP Then

            With mvarPictureRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - (picHeight + mvarPadding + capHeight)) * 0.5) + mvarClientRect.Top
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - picWidth) * 0.5) + mvarClientRect.Left
                .Bottom = .Top + picHeight
                .Right = .Left + picWidth
            End With

            With mvarCaptionRect
                .Top = mvarPictureRect.Bottom + mvarPadding
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - capWidth) * 0.5) + mvarClientRect.Left
                .Bottom = .Top + capHeight
                .Right = .Left + capWidth
            End With

        ElseIf m_PicturePosition = gbBOTTOM Then

            With mvarCaptionRect
                .Top = (((mvarClientRect.Bottom - mvarClientRect.Top) - (picHeight + mvarPadding + capHeight)) * 0.5) + mvarClientRect.Top
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - capWidth) * 0.5) + mvarClientRect.Left
                .Bottom = .Top + capHeight
                .Right = .Left + capWidth
            End With

            With mvarPictureRect
                .Top = mvarCaptionRect.Bottom + mvarPadding
                .Left = (((mvarClientRect.Right - mvarClientRect.Left) - picWidth) * 0.5) + mvarClientRect.Left
                .Bottom = .Top + picHeight
                .Right = .Left + picWidth
            End With

        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CalculateCaptionRect
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub CalculateCaptionRect()

    Dim mvarWidth        As Long
    Dim mvarHeight       As Long
    Dim dtTextDrawParams As Long

    With mvarDrawTextParams
        .iLeftMargin = 1
        .iRightMargin = 1
        .iTabLength = 1
        .cbSize = Len(mvarDrawTextParams)
    End With

    dtTextDrawParams = dtDefTextDrawParams
    '    If Ambient.RightToLeft = True Then
    '        dtTextDrawParams = dtTextDrawParams Or WS_EX_RTLREADING
    '    End If
    DrawTextExW hDC, StrPtr(m_Caption & vbNullChar), -1, mvarCaptionRect, DT_CALCRECT Or dtTextDrawParams, mvarDrawTextParams

    With mvarCaptionRect
        mvarWidth = .Right - .Left
        mvarHeight = .Bottom - .Top
        .Left = mvarClientRect.Left + (((mvarClientRect.Right - mvarClientRect.Left) - (.Right - .Left)) * 0.5)
        .Top = mvarClientRect.Top + (((mvarClientRect.Bottom - mvarClientRect.Top) - (.Bottom - .Top)) * 0.5)
        .Right = .Left + mvarWidth
        .Bottom = .Top + mvarHeight
    End With

End Sub

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
Public Property Let Caption(ByVal New_Caption As String)
Attribute Caption.VB_UserMemId = -518
    m_Caption = New_Caption
    SetAccessKeys
    Refresh
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   clr (OLE_COLOR)
'!--------------------------------------------------------------------------------
Private Function CColor(ByVal clr As OLE_COLOR) As Long

    ' If it's a system color, get the RGB value.
    If clr And &H80000000 Then
        CColor = GetSysColor(clr And (Not &H80000000))
    Else
        CColor = clr
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property CheckExist
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get CheckExist() As Boolean
    CheckExist = m_MenuExist
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property CheckExist
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_CheckExist (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let CheckExist(ByVal New_CheckExist As Boolean)
    m_CheckExist = New_CheckExist
    PropertyChanged "CheckExist"
    Refresh
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function COLOR_DarkenColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Color (Long)
'                              Value (Long)
'!--------------------------------------------------------------------------------
Private Function COLOR_DarkenColor(ByVal Color As Long, ByVal Value As Long) As Long

    Dim cc As RGB
    Dim R  As Integer
    Dim G  As Integer
    Dim B  As Integer

    CopyMemory ByVal VarPtr(cc), ByVal VarPtr(Color), 3

    With cc
        B = .Blue + Value
        G = .Green + Value
        R = .Red + Value
    End With

    If R < 0 Then
        R = 0
    End If

    If R > 255 Then
        R = 255
    End If

    If G < 0 Then
        G = 0
    End If

    If G > 255 Then
        G = 255
    End If

    If B < 0 Then
        B = 0
    End If

    If B > 255 Then
        B = 255
    End If

    COLOR_DarkenColor = RGB(R, G, B)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawCaption
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub DrawCaption()

    Dim g_tmpFontColor      As OLE_COLOR
    Dim mvarCaptionRect_Iki As RECT
    Dim dtTextDrawParams    As Long

    dtTextDrawParams = dtDefTextDrawParams

    '    If Ambient.RightToLeft = True Then
    '        dtTextDrawParams = dtTextDrawParams Or WS_EX_RTLREADING
    '    End If
    If Enabled Then
        SetTextColor hDC, CColor(m_TextColor)
        'CColor(&H80000012)
        mvarOrgRect = mvarCaptionRect

        If g_MouseDown = 1 Then

            With mvarCaptionRect
                .Left = .Left + 1
                .Top = .Top + 1
                .Right = .Right + 1
                .Bottom = .Bottom + 1
            End With

        End If

        DrawTextExW hDC, StrPtr(m_Caption & vbNullChar), -1, mvarCaptionRect, dtTextDrawParams, mvarDrawTextParams
        mvarCaptionRect = mvarOrgRect
    Else
        g_tmpFontColor = m_TextColor
        SetTextColor hDC, CColor(&H80000014)

        With mvarCaptionRect_Iki
            .Bottom = mvarCaptionRect.Bottom
            .Left = mvarCaptionRect.Left + 1
            .Right = mvarCaptionRect.Right + 1
            .Top = mvarCaptionRect.Top + 1
        End With

        DrawTextExW hDC, StrPtr(m_Caption & vbNullChar), -1, mvarCaptionRect_Iki, dtTextDrawParams, mvarDrawTextParams
        SetTextColor hDC, CColor(&H80000010)
        DrawTextExW hDC, StrPtr(m_Caption & vbNullChar), -1, mvarCaptionRect, dtTextDrawParams, mvarDrawTextParams
        SetTextColor hDC, CColor(g_tmpFontColor)
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawPicture
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub DrawPicture()
    mvarOrgRect = mvarPictureRect

    With mvarPictureRect
        .Left = .Left + g_MouseDown
        .Top = .Top + g_MouseDown
        .Right = .Right + g_MouseDown
        .Bottom = .Bottom + g_MouseDown

        If m_Picture.Type = 1 Then
            Picture1.AutoRedraw = True
            Picture1.Cls
            Picture1.PaintPicture m_Picture, 0, 0

            If Not m_UseMaskColor Then
                m_MaskColor = CColor(GetPixel(Picture1.hDC, 0, 0))
            End If

            DoEvents
            TransparentBlt UserControl.hDC, .Left, .Top, .Right - .Left, .Bottom - .Top, Picture1.hDC, 0, 0, .Right - .Left, .Bottom - .Top, m_MaskColor
            UserControl.Refresh
            Picture1.AutoRedraw = False
        ElseIf m_Picture.Type = 3 Then
            UserControl.PaintPicture m_Picture, .Left, .Top, .Right - .Left, .Bottom - .Top, 0, 0, m_PictureWidth, m_PictureHeight
        End If

    End With

    mvarPictureRect = mvarOrgRect
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawRect
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   DestHDC (Long)
'                              RectLEFT (Long)
'                              RectTOP (Long)
'                              RectRIGHT (Long)
'                              RectBOTTOM (Long)
'                              MyColor (Long)
'                              FillRectWithColor (Byte = 0)
'!--------------------------------------------------------------------------------
Private Sub DrawRect(DestHDC As Long, ByVal RectLEFT As Long, ByVal RectTOP As Long, ByVal RectRIGHT As Long, ByVal RectBOTTOM As Long, ByVal MyColor As Long, Optional ByVal FillRectWithColor As Byte = 0)

    Dim MyRect As RECT
    Dim Firca  As Long

    Firca = CreateSolidBrush(CColor(MyColor))

    With MyRect
        .Left = RectLEFT
        .Top = RectTOP
        .Right = RectRIGHT
        .Bottom = RectBOTTOM
    End With

    If FillRectWithColor = 1 Then
        FillRect DestHDC, MyRect, Firca
    Else
        FrameRect DestHDC, MyRect, Firca
    End If

    DeleteObject Firca
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawWinXPButton
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Press (Byte)
'                              HOVERING (Byte)
'!--------------------------------------------------------------------------------
Private Sub DrawWinXPButton(ByVal Press As Byte, Optional HOVERING As Byte)

    Dim X            As Long
    Dim Intg         As Single
    Dim curBackColor As Long

    curBackColor = COLOR_DarkenColor(CColor(&H8000000F), 48)

    If Enabled Then
        If m_XPDefaultColors Then
            m_XPColor_Pressed = RGB(140, 170, 230)
            m_XPColor_Hover = RGB(225, 153, 71)
        End If

        If UserScaleH = 0 Then

            Exit Sub

        End If

        If Press = 0 Then
            Intg = 50 / UserScaleH

            For X = 1 To UserScaleH
                Line (0, X)-(UserScaleW, X), COLOR_DarkenColor(vbWhite, -Intg * X)
            Next

            DrawRect hDC, 0, 0, UserScaleW, UserScaleH, &H80000015

            If HOVERING = 1 Or g_HasFocus = 1 Then
                Intg = CColor(IIf(HOVERING, m_XPColor_Hover, m_XPColor_Pressed))
                DrawRect hDC, 1, 2, UserScaleW - 1, UserScaleH - 2, Intg
                Line (2, UserScaleH - 2)-(UserScaleW - 2, UserScaleH - 2), COLOR_DarkenColor(Intg, -40)
                Line (2, 1)-(UserScaleW - 2, 1), COLOR_DarkenColor(Intg, 90)
                Line (1, 2)-(UserScaleW - 1, 2), COLOR_DarkenColor(Intg, 35)
                curBackColor = COLOR_DarkenColor(Intg, 20)
                Line (2, 3)-(2, UserScaleH - 3), curBackColor
                Line (UserScaleW - 3, 3)-(UserScaleW - 3, UserScaleH - 3), curBackColor
                SetPixel hDC, 3, UserScaleH - 4, Intg
                SetPixel hDC, UserScaleW - 4, UserScaleH - 4, Intg
                Intg = COLOR_DarkenColor(Intg, 35)
                SetPixel hDC, UserScaleW - 4, 3, Intg
                SetPixel hDC, 3, 3, Intg
            End If

        Else
            Intg = 25 / UserScaleH
            curBackColor = COLOR_DarkenColor(curBackColor, -32)

            For X = 1 To UserScaleH
                Line (0, UserScaleH - X)-(UserScaleW, UserScaleH - X), COLOR_DarkenColor(curBackColor, -Intg * X)
            Next

            DrawRect hDC, 0, 0, UserScaleW, UserScaleH, &H80000015
        End If

        Intg = &H80000015
    Else
        DrawRect hDC, 0, 0, UserScaleW, UserScaleH, COLOR_DarkenColor(curBackColor, -24), 1
        DrawRect hDC, 0, 0, UserScaleW, UserScaleH, COLOR_DarkenColor(curBackColor, -84)
        Intg = COLOR_DarkenColor(curBackColor, -72)
    End If

    curBackColor = CColor(&H8000000F)
    Line (0, 0)-(1, 1), curBackColor, BF
    SetPixel hDC, 1, 1, Intg
    Line (0, UserScaleH - 2)-(1, UserScaleH), curBackColor, BF
    SetPixel hDC, 1, UserScaleH - 2, Intg
    Line (UserScaleW - 2, 0)-(UserScaleW, 1), curBackColor, BF
    SetPixel hDC, UserScaleW - 2, 1, Intg
    Line (UserScaleW - 2, UserScaleH - 2)-(UserScaleW, UserScaleH), curBackColor, BF
    SetPixel hDC, UserScaleW - 2, UserScaleH - 2, Intg
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property EnabledCtrl
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get EnabledCtrl() As Boolean
    EnabledCtrl = Enabled
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property EnabledCtrl
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_Enabled (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let EnabledCtrl(ByVal New_Enabled As Boolean)
    Enabled = New_Enabled
    g_HasFocus = 0
    g_MouseDown = 0
    g_MouseIn = 0
    Refresh
End Property

'Public Property Get Font() As Font
'    Set Font = m_TextFont
'
'End Property
'Public Property Set Font(ByVal New_Font As Font)
'    Set m_TextFont = New_Font
'    Set UserControl.Font = New_Font
'    PropertyChanged "Font"
'    UserControl_Resize
'
'End Property
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
    Call OLEFontToLogFont(NewFont, ButtonLogFont)
    OldFontHandle = ButtonFontHandle
    ButtonFontHandle = CreateFontIndirect(ButtonLogFont)

    If UserControl.hDC <> 0 Then SendMessage UserControl.hDC, WM_SETFONT, ButtonFontHandle, ByVal 1&
    If OldFontHandle <> 0 Then DeleteObject OldFontHandle
    Me.Refresh
    UserControl.PropertyChanged "Font"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PropFont_FontChanged
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropertyName (String)
'!--------------------------------------------------------------------------------
Private Sub PropFont_FontChanged(ByVal PropertyName As String)

    Dim OldFontHandle As Long

    Call OLEFontToLogFont(PropFont, ButtonLogFont)
    OldFontHandle = ButtonFontHandle
    ButtonFontHandle = CreateFontIndirect(ButtonLogFont)

    If UserControl.hDC <> 0 Then SendMessage UserControl.hDC, WM_SETFONT, ButtonFontHandle, ByVal 1&
    If OldFontHandle <> 0 Then DeleteObject OldFontHandle
    Me.Refresh
    UserControl.PropertyChanged "Font"
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
'! Procedure   (Функция)   :   Property MaskColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get MaskColor() As OLE_COLOR
    MaskColor = m_MaskColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MaskColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_MaskColor (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    Refresh
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MenuCaption
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Public Property Get MenuCaption(Index As Integer) As String
    MenuCaption = mnuMenu(Index).Caption
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MenuCaption
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'                              New_MenuCaption (String)
'!--------------------------------------------------------------------------------
Public Property Let MenuCaption(Index As Integer, ByVal New_MenuCaption As String)
    mnuMenu(Index).Caption = New_MenuCaption
    PropertyChanged "MenuCaption"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MenuChecked
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Public Property Get MenuChecked(Index As Integer) As Boolean
    MenuChecked = mnuMenu(Index).Checked
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MenuChecked
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'                              New_MenuChecked (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let MenuChecked(Index As Integer, ByVal New_MenuChecked As Boolean)
    mnuMenu(Index).Checked = New_MenuChecked
    PropertyChanged "MenuChecked"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function MenuCount
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function MenuCount() As Long
    MenuCount = mnuMenu.UBound
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MenuEnabled
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Public Property Get MenuEnabled(Index As Integer) As Boolean
    MenuEnabled = mnuMenu(Index).Enabled
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MenuEnabled
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'                              New_MenuEnabled (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let MenuEnabled(Index As Integer, ByVal New_MenuEnabled As Boolean)
    mnuMenu(Index).Enabled = New_MenuEnabled
    PropertyChanged "MenuEnabled"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MenuExist
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get MenuExist() As Boolean
    MenuExist = m_MenuExist
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MenuExist
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_MenuExist (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let MenuExist(ByVal New_MenuExist As Boolean)
    m_MenuExist = New_MenuExist
    PropertyChanged "MenuExist"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MenuVisible
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Public Property Get MenuVisible(Index As Integer) As Boolean
    MenuVisible = mnuMenu(Index).Visible
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property MenuVisible
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'                              New_MenuVisible (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let MenuVisible(Index As Integer, ByVal New_MenuVisible As Boolean)
    mnuMenu(Index).Visible = New_MenuVisible
    PropertyChanged "MenuVisible"
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
'! Procedure   (Функция)   :   Sub mnuMenu_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuMenu_Click(Index As Integer)
    RaiseEvent ClickMenu(Index)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Picture
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PicturePosition
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get PicturePosition() As XBPicturePosition
    PicturePosition = m_PicturePosition
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property PicturePosition
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_PicturePosition (XBPicturePosition)
'!--------------------------------------------------------------------------------
Public Property Let PicturePosition(ByVal New_PicturePosition As XBPicturePosition)
    m_PicturePosition = New_PicturePosition
    Refresh
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Picture
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_Picture (Picture)
'!--------------------------------------------------------------------------------
Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture

    With UserControl

        If m_Picture Is Nothing = False Then
            m_PictureWidth = .ScaleX(m_Picture.Width, 8, 3)
            m_PictureHeight = .ScaleY(m_Picture.Height, 8, 3)
        End If

    End With

    Refresh
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub pInitialize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub pInitialize()
    dtDefTextDrawParams = DT_WORDBREAK Or DT_VCENTER Or DT_CENTER
    ScaleMode = 3
    PaletteMode = 3
    UserScaleW = UserControl.ScaleWidth
    UserScaleH = UserControl.ScaleHeight

    If UserScaleW < 10 Then
        UserControl.Width = 150
    End If

    If UserScaleH < 10 Then
        UserControl.Height = 150
    End If

    With g_FocusRect
        .Left = 4
        .Right = UserScaleW - 4
        .Top = 4
        .Bottom = UserScaleH - 4
    End With

    Refresh
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Refresh
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Refresh()
    AutoRedraw = True
    UserControl.Cls

    If m_ButtonStyle = gbWinXP Then
        DrawWinXPButton g_MouseDown, g_MouseIn
    Else

        If g_MouseDown = 1 Then
            DrawRect hDC, 0, 0, UserScaleW, UserScaleH, &H80000014
            DrawRect hDC, 0, 0, UserScaleW + 1, UserScaleH + 1, 0
        ElseIf g_MouseIn = 1 Or m_ButtonStyle = gbStandard Then
            DrawRect hDC, 0, 0, UserScaleW, UserScaleH, 0
            DrawRect hDC, 0, 0, UserScaleW + 1, UserScaleH + 1, &H80000014
        End If
    End If

    CalcRECTs

    If LenB(m_Caption) > 0 Then
        DrawCaption
    End If

    If m_Picture Is Nothing = False Then
        DrawPicture
    End If

    If g_HasFocus = 1 Then
        If m_ShowFocusRect Then
            If m_ButtonStyle <> gbWinXP Then
                DrawFocusRect hDC, g_FocusRect
            End If
        End If
    End If

    UserControl.Refresh
    AutoRedraw = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetAccessKeys
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SetAccessKeys()

    Dim ampersandPos As Long

    With UserControl

        If Len(m_Caption) > 1 Then
            ampersandPos = InStr(m_Caption, "&")

            If ampersandPos < Len(m_Caption) Then
                If ampersandPos > 0 Then
                    If Mid$(m_Caption, ampersandPos + 1, 1) <> "&" Then
                        .AccessKeys = LCase$(Mid$(m_Caption, ampersandPos + 1, 1))
                    Else
                        ampersandPos = InStr(ampersandPos + 2, m_Caption, "&", vbTextCompare)

                        If Mid$(m_Caption, ampersandPos + 1, 1) <> "&" Then
                            .AccessKeys = LCase$(Mid$(m_Caption, ampersandPos + 1, 1))
                        Else
                            .AccessKeys = vbNullString
                        End If
                    End If
                End If

            Else
                .AccessKeys = vbNullString
            End If

        Else
            .AccessKeys = vbNullString
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ShowFocusRect
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowFocusRect
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ShowFocusRect
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_ShowFocusRect (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
    m_ShowFocusRect = New_ShowFocusRect
    Refresh
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
Public Property Let TextColor(ByVal new_TextColor As OLE_COLOR)
    m_TextColor = new_TextColor
    PropertyChanged "TextColor"
    Refresh
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseMaskColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get UseMaskColor() As Boolean
    UseMaskColor = m_UseMaskColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseMaskColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_V (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let UseMaskColor(ByVal New_V As Boolean)
    m_UseMaskColor = New_V
    Refresh
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_AccessKeyPress
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyAscii (Integer)
'!--------------------------------------------------------------------------------
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

    If Enabled Then
        If m_MenuExist Then
            If mnuMenu.Count > 1 Then
                PopupMenu mnuGeneral, 2, ScaleLeft, ScaleHeight + ScaleTop
            Else
                RaiseEvent Click
            End If

        Else
            RaiseEvent Click
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_AmbientChanged
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropertyName (String)
'!--------------------------------------------------------------------------------
Private Sub UserControl_AmbientChanged(PropertyName As String)
    pInitialize
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_EnterFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_EnterFocus()
    g_MouseIn = 0
    g_HasFocus = 1
    Refresh
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_ExitFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_ExitFocus()
    g_HasFocus = 0
    g_MouseDown = 0
    g_MouseIn = 0
    Refresh
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_InitProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()
    Set PropFont = Ambient.Font
    Set UserControl.Font = PropFont
    BackColor = vbButtonFace
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_KeyDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If g_MouseDown = 0 Then
        If KeyCode = 32 Then
            g_MouseDown = 1
            g_MouseIn = 1
            Refresh
        End If
    End If

    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_KeyUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 32 Then
        g_MouseDown = 0
        g_MouseIn = 0
        Refresh
        UserControl_MouseUp 1, Shift, 0, 0
    End If

    RaiseEvent KeyUp(KeyCode, Shift)
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

    If Button < 2 Then
        g_MouseDown = 1
        Refresh
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

    Dim P As POINT

    GetCursorPos P

    If g_MouseIn = 0 Then
        ReleaseCapture
        g_MouseDown = Button
        g_MouseIn = 1
        RaiseEvent MouseIn(Shift)
        Refresh
        SetCapture UserControl.hWnd
    ElseIf hWnd <> WindowFromPoint(P.X, P.Y) Then
        ReleaseCapture
        g_MouseIn = 0
        g_MouseDown = 0
        RaiseEvent MouseOut(Shift)
        Refresh
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

    Dim P As POINT

    ReleaseCapture

    If Button < 2 Then
        GetCursorPos P

        If hWnd = WindowFromPoint(P.X, P.Y) Then
            If m_MenuExist Then
                If mnuMenu.Count > 1 Then
                    PopupMenu mnuGeneral, 2, ScaleLeft, ScaleHeight + ScaleTop
                Else
                    RaiseEvent Click
                End If

            Else
                RaiseEvent Click
            End If
        End If
    End If

    RaiseEvent MouseUp(Button, Shift, X, Y)
    g_MouseDown = 0
    g_MouseIn = 0
    Refresh
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_ReadProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropBag (PropertyBag)
'!--------------------------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Dim Index As Integer

    With PropBag
        Set PropFont = .ReadProperty("Font", Ambient.Font)
        Set UserControl.Font = PropFont
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_PicturePosition = .ReadProperty("PicturePosition", 1)
        m_ButtonStyle = .ReadProperty("ButtonStyle", 2)
        m_PictureWidth = .ReadProperty("PictureWidth", 16)
        m_PictureHeight = .ReadProperty("PictureHeight", 16)
        Set m_Picture = .ReadProperty("Picture", Nothing)
        Enabled = .ReadProperty("Enabled", True)
        m_XPColor_Pressed = .ReadProperty("XPColor_Pressed", &H80000014)
        m_XPColor_Hover = .ReadProperty("XPColor_Hover", &H80000016)
        m_XPDefaultColors = .ReadProperty("XPDefaultColors", 1)
        m_MaskColor = .ReadProperty("MaskColor", 0)
        m_UseMaskColor = .ReadProperty("UseMaskColor", 0)
        m_ShowFocusRect = .ReadProperty("ShowFocusRect", 1)
        m_TextColor = .ReadProperty("TextColor", Ambient.ForeColor)
        UserControl.BackColor = .ReadProperty("BackColor", vbButtonFace)
        mnuMenu(Index).Caption = .ReadProperty("MenuCaption" & Index, vbNullString)
        mnuMenu(Index).Checked = .ReadProperty("MenuChecked" & Index, False)
        mnuMenu(Index).Enabled = .ReadProperty("MenuEnabled" & Index, True)
        mnuMenu(Index).Visible = .ReadProperty("MenuVisible" & Index, True)
        m_MenuExist = .ReadProperty("MenuExist", False)
        m_CheckExist = .ReadProperty("CheckExist", False)
    End With

    SetAccessKeys
    pInitialize
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Resize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Resize()

    On Error Resume Next

    pInitialize
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
    'Clean up Picture
    Set m_Picture = Nothing
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_WriteProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropBag (PropertyBag)
'!--------------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Dim Index As Integer

    With PropBag
        .WriteProperty "Font", PropFont, Ambient.Font
        .WriteProperty "Caption", m_Caption, Ambient.DisplayName
        .WriteProperty "PicturePosition", m_PicturePosition, 1
        .WriteProperty "ButtonStyle", m_ButtonStyle, 2
        .WriteProperty "Picture", m_Picture, Nothing
        .WriteProperty "PictureWidth", m_PictureWidth, 16
        .WriteProperty "PictureHeight", m_PictureHeight, 16
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "ShowFocusRect", m_ShowFocusRect, 1
        .WriteProperty "XPColor_Pressed", m_XPColor_Pressed, &H80000014
        .WriteProperty "XPColor_Hover", m_XPColor_Hover, &H80000016
        .WriteProperty "XPDefaultColors", m_XPDefaultColors, 1
        .WriteProperty "MaskColor", m_MaskColor, 0
        .WriteProperty "UseMaskColor", m_UseMaskColor, 0
        .WriteProperty "TextColor", m_TextColor, Ambient.ForeColor
        .WriteProperty "BackColor", UserControl.BackColor, vbButtonFace
        .WriteProperty "MenuCaption" & Index, mnuMenu(Index).Caption, vbNullString
        .WriteProperty "MenuChecked" & Index, mnuMenu(Index).Checked, False
        .WriteProperty "MenuEnabled" & Index, mnuMenu(Index).Enabled, True
        .WriteProperty "MenuVisible" & Index, mnuMenu(Index).Visible, True
        .WriteProperty "MenuExist", m_MenuExist, False
        .WriteProperty "CheckExist", m_CheckExist, False
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property XPColor_Hover
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get XPColor_Hover() As OLE_COLOR
    XPColor_Hover = m_XPColor_Hover
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property XPColor_Hover
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_XPColor_Hover (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let XPColor_Hover(ByVal New_XPColor_Hover As OLE_COLOR)
    m_XPColor_Hover = New_XPColor_Hover
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property XPColor_Pressed
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get XPColor_Pressed() As OLE_COLOR
    XPColor_Pressed = m_XPColor_Pressed
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property XPColor_Pressed
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_XPColor_Pressed (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let XPColor_Pressed(ByVal New_XPColor_Pressed As OLE_COLOR)
    m_XPColor_Pressed = New_XPColor_Pressed
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property XPDefaultColors
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get XPDefaultColors() As Boolean
    XPDefaultColors = m_XPDefaultColors
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property XPDefaultColors
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   New_XPDefaultColors (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let XPDefaultColors(ByVal New_XPDefaultColors As Boolean)
    m_XPDefaultColors = New_XPDefaultColors
End Property
