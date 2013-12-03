VERSION 5.00
Begin VB.UserControl ctlXpButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   ClipControls    =   0   'False
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
Private UserScaleW                      As Long
Private UserScaleH                      As Long

Private Type textparametreleri
    cbSize                                  As Long
    iTabLength                          As Long
    iLeftMargin                         As Long
    iRightMargin                        As Long
    uiLengthDrawn                       As Long

End Type

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

Private mvarClientRect                  As RECT
Private mvarPictureRect                 As RECT
Private mvarCaptionRect                 As RECT
Private mvarOrgRect                     As RECT
Private g_FocusRect                     As RECT
Private alan                            As RECT
Private m_Picture                       As Picture
Private m_PicturePosition               As XBPicturePosition
Private m_ButtonStyle                   As XBButtonStyle
Private mvarDrawTextParams              As textparametreleri
Private m_Caption                       As String
Private m_PictureWidth                  As Long
Private m_PictureHeight                 As Long
Private g_HasFocus                      As Byte
Private g_MouseDown                     As Byte
Private g_MouseIn                       As Byte
Private m_ShowFocusRect                 As Boolean
Private m_MaskColor                     As Long
Private m_UseMaskColor                  As Long
Private m_XPDefaultColors               As Boolean
Private m_MenuExist                     As Boolean
Private m_CheckExist                    As Boolean
Private m_XPColor_Pressed               As Long
Private m_XPColor_Hover                 As Long
Private m_TextFont                      As Font
Private m_TextColor                     As OLE_COLOR

Private Const mvarPadding               As Byte = 4

Public Event ClickMenu(mnuIndex As Integer)
Public Event Click()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseIn(Shift As Integer)
Public Event MouseOut(Shift As Integer)

Private Declare Function DrawTextEx _
                          Lib "user32.dll" _
                              Alias "DrawTextExA" (ByVal hDC As Long, _
                                                   ByVal lpsz As String, _
                                                   ByVal n As Long, _
                                                   lpRect As RECT, _
                                                   ByVal un As Long, _
                                                   lpDrawTextParams As textparametreleri) As Long

Public Sub AddMenu(ByVal sCaption As String)

Dim iCount                              As Integer

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

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal new_BackColor As OLE_COLOR)
    UserControl.BackColor = new_BackColor
    PropertyChanged "BackColor"

End Property

Public Property Get ButtonStyle() As XBButtonStyle
    ButtonStyle = m_ButtonStyle

End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As XBButtonStyle)
    m_ButtonStyle = New_ButtonStyle
    Refresh

End Property

Private Sub CalcRECTs()

Dim picWidth                            As Long
Dim picHeight                           As Long
Dim capWidth                            As Long
Dim capHeight                           As Long

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

Private Sub CalculateCaptionRect()

Dim mvarWidth                           As Long
Dim mvarHeight                          As Long

    With mvarDrawTextParams
        .iLeftMargin = 1
        .iRightMargin = 1
        .iTabLength = 1
        .cbSize = Len(mvarDrawTextParams)

    End With

    DrawTextEx hDC, m_Caption, Len(m_Caption), mvarCaptionRect, 1045, mvarDrawTextParams

    With mvarCaptionRect
        mvarWidth = .Right - .Left
        mvarHeight = .Bottom - .Top
        .Left = mvarClientRect.Left + (((mvarClientRect.Right - mvarClientRect.Left) - (.Right - .Left)) * 0.5)
        .Top = mvarClientRect.Top + (((mvarClientRect.Bottom - mvarClientRect.Top) - (.Bottom - .Top)) * 0.5)
        .Right = .Left + mvarWidth
        .Bottom = .Top + mvarHeight

    End With

End Sub

Public Property Get Caption() As String
    Caption = m_Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)
Attribute Caption.VB_UserMemId = -518
    m_Caption = New_Caption
    SetAccessKeys
    Refresh

End Property

Private Function CColor(ByVal clr As OLE_COLOR) As Long

' If it's a system color, get the RGB value.
    If clr And &H80000000 Then
        CColor = GetSysColor(clr And (Not &H80000000))
    Else
        CColor = clr

    End If

End Function

Public Property Get CheckExist() As Boolean
    CheckExist = m_MenuExist

End Property

Public Property Let CheckExist(ByVal New_CheckExist As Boolean)
    m_CheckExist = New_CheckExist
    PropertyChanged "CheckExist"
    Refresh

End Property

Private Function COLOR_DarkenColor(ByVal Color As Long, ByVal Value As Long) As Long

Dim cc                                  As RGB
Dim R                                   As Integer
Dim G                                   As Integer
Dim B                                   As Integer

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

Private Sub DrawCaption()

Dim g_tmpFontColor                      As OLE_COLOR
Dim mvarCaptionRect_Iki                 As RECT

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

        DrawTextEx hDC, m_Caption, Len(m_Caption), mvarCaptionRect, 21, mvarDrawTextParams
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

        DrawTextEx hDC, m_Caption, Len(m_Caption), mvarCaptionRect_Iki, 21, mvarDrawTextParams
        SetTextColor hDC, CColor(&H80000010)
        DrawTextEx hDC, m_Caption, Len(m_Caption), mvarCaptionRect, 21, mvarDrawTextParams
        SetTextColor hDC, CColor(g_tmpFontColor)

    End If

End Sub

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

Private Sub DrawRect(DestHDC As Long, _
                     ByVal RectLEFT As Long, _
                     ByVal RectTOP As Long, _
                     ByVal RectRIGHT As Long, _
                     ByVal RectBOTTOM As Long, _
                     ByVal MyColor As Long, _
                     Optional ByVal FillRectWithColor As Byte = 0)

Dim MyRect                              As RECT
Dim Firca                               As Long

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

Private Sub DrawWinXPButton(ByVal Press As Byte, Optional HOVERING As Byte)

Dim X                                   As Long
Dim Intg                                As Single
Dim curBackColor                        As Long

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

Public Property Get EnabledCtrl() As Boolean
    EnabledCtrl = Enabled

End Property

Public Property Let EnabledCtrl(ByVal New_Enabled As Boolean)
    Enabled = New_Enabled
    g_HasFocus = 0
    g_MouseDown = 0
    g_MouseIn = 0
    Refresh

End Property

Public Property Get Font() As Font
    Set Font = m_TextFont

End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_TextFont = New_Font
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    UserControl_Resize

End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = m_MaskColor

End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    Refresh

End Property

Public Property Get MenuCaption(Index As Integer) As String
    MenuCaption = mnuMenu(Index).Caption

End Property

Public Property Let MenuCaption(Index As Integer, ByVal New_MenuCaption As String)
    mnuMenu(Index).Caption = New_MenuCaption
    PropertyChanged "MenuCaption"

End Property

Public Property Get MenuChecked(Index As Integer) As Boolean
    MenuChecked = mnuMenu(Index).Checked

End Property

Public Property Let MenuChecked(Index As Integer, ByVal New_MenuChecked As Boolean)
    mnuMenu(Index).Checked = New_MenuChecked
    PropertyChanged "MenuChecked"

End Property

Public Function MenuCount() As Long
    MenuCount = mnuMenu.UBound

End Function

Public Property Get MenuEnabled(Index As Integer) As Boolean
    MenuEnabled = mnuMenu(Index).Enabled

End Property

Public Property Let MenuEnabled(Index As Integer, ByVal New_MenuEnabled As Boolean)
    mnuMenu(Index).Enabled = New_MenuEnabled
    PropertyChanged "MenuEnabled"

End Property

Public Property Get MenuExist() As Boolean
    MenuExist = m_MenuExist

End Property

Public Property Let MenuExist(ByVal New_MenuExist As Boolean)
    m_MenuExist = New_MenuExist
    PropertyChanged "MenuExist"

End Property

Public Property Get MenuVisible(Index As Integer) As Boolean
    MenuVisible = mnuMenu(Index).Visible

End Property

Public Property Let MenuVisible(Index As Integer, ByVal New_MenuVisible As Boolean)
    mnuMenu(Index).Visible = New_MenuVisible
    PropertyChanged "MenuVisible"

End Property

Public Property Get mhwnd() As Long
    mhwnd = UserControl.hWnd

End Property

Private Sub mnuMenu_Click(Index As Integer)
    RaiseEvent ClickMenu(Index)

End Sub

Public Property Get Picture() As Picture
    Set Picture = m_Picture

End Property

Public Property Get PicturePosition() As XBPicturePosition
    PicturePosition = m_PicturePosition

End Property

Public Property Let PicturePosition(ByVal New_PicturePosition As XBPicturePosition)
    m_PicturePosition = New_PicturePosition
    Refresh

End Property

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

Private Sub pInitialize()
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

Private Sub SetAccessKeys()

Dim ampersandPos                        As Long

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

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowFocusRect

End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
    m_ShowFocusRect = New_ShowFocusRect
    Refresh

End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor

End Property

Public Property Let TextColor(ByVal new_TextColor As OLE_COLOR)
    m_TextColor = new_TextColor
    PropertyChanged "TextColor"
    Refresh

End Property

Public Property Get UseMaskColor() As Boolean
    UseMaskColor = m_UseMaskColor

End Property

Public Property Let UseMaskColor(ByVal New_V As Boolean)
    m_UseMaskColor = New_V
    Refresh

End Property

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

Private Sub UserControl_AmbientChanged(PropertyName As String)
    pInitialize

End Sub

Private Sub UserControl_EnterFocus()
    g_MouseIn = 0
    g_HasFocus = 1
    Refresh

End Sub

Private Sub UserControl_ExitFocus()
    g_HasFocus = 0
    g_MouseDown = 0
    g_MouseIn = 0
    Refresh

End Sub

Private Sub UserControl_InitProperties()
    BackColor = vbButtonFace

End Sub

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

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 32 Then
        g_MouseDown = 0
        g_MouseIn = 0
        Refresh
        UserControl_MouseUp 1, Shift, 0, 0

    End If

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    If Button < 2 Then
        g_MouseDown = 1
        Refresh

    End If

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

Dim P                                   As POINT

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

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

Dim P                                   As POINT

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

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Dim Index                               As Integer

    With PropBag
        Set Font = .ReadProperty("Font", Ambient.Font)
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

Private Sub UserControl_Resize()

    On Error Resume Next

    pInitialize

End Sub

Private Sub UserControl_Terminate()

    On Error Resume Next

    Set m_Picture = Nothing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Dim Index                               As Integer

    With PropBag
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
        .WriteProperty "Font", m_TextFont, Ambient.Font
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

Public Property Get XPColor_Hover() As OLE_COLOR
    XPColor_Hover = m_XPColor_Hover

End Property

Public Property Let XPColor_Hover(ByVal New_XPColor_Hover As OLE_COLOR)
    m_XPColor_Hover = New_XPColor_Hover

End Property

Public Property Get XPColor_Pressed() As OLE_COLOR
    XPColor_Pressed = m_XPColor_Pressed

End Property

Public Property Let XPColor_Pressed(ByVal New_XPColor_Pressed As OLE_COLOR)
    m_XPColor_Pressed = New_XPColor_Pressed

End Property

Public Property Get XPDefaultColors() As Boolean
    XPDefaultColors = m_XPDefaultColors

End Property

Public Property Let XPDefaultColors(ByVal New_XPDefaultColors As Boolean)
    m_XPDefaultColors = New_XPDefaultColors

End Property
