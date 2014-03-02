VERSION 5.00
Begin VB.UserControl ctlColorButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   302
   ToolboxBitmap   =   "ctlColorButton.ctx":0000
   Begin VB.PictureBox picDropDown 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   0
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   154
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   0
      Picture         =   "ctlColorButton.ctx":00FA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDropDown 
      Height          =   45
      Left            =   360
      Picture         =   "ctlColorButton.ctx":0244
      Top             =   120
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "ctlColorButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------
'Excel Style Color Picker
'Version 1.0
'
'Copyright © 2004 by Grid2000.com. All Rights Reserved.
'
'Copyright notice:
'This can be considered to be Version 2.0 :))
'The code of this module was downloaded from the following
'link: http://www.freevbcode.com/ShowCode.Asp?ID=6638
'
'After that it was sufficiently improved:
'- events TrackColor, MouseIn, MouseOut, DropDownOpen and
'  DropDownClose were added;
'- properties BackColor, ColorPalette, ForbiddenColor,
'  UseForbiddenColor, DropDownCaption and Style were added;
'- sub DropDown was made Public - so now it's a METHOD.
'
'The most pleasant things are: support of 16-color (VGA)
'palette; possibility to track color while mouse moves in
'the dropdown part of the control; possibility to forbid
'selection of a certain color (it's "cell" simply can't be
'clicked in the dropdown part of the control).
'
'Also I got rid of the CommonDialog control - I use API
'ChooseColor (from "comdlg32.dll") instead.
'
'Author of all these improvements is Yumashin Alexey,
'a.k.a. Comanche, e-mail: pc-er@mail.ru.
'
'---------------------------------------------------------
Option Explicit
Private Type RECT
    Left                                                As Long
    Top                                                 As Long
    Right                                               As Long
    Bottom                                              As Long
End Type
Private Const DT_CENTER                             As Long = &H1
Private Const DT_SINGLELINE                         As Long = &H20
Private Const DT_VCENTER                            As Long = &H4
Private Const DI_NORMAL                             As Long = &H3
Private Const BF_RECT                               As Long = &HF
Private Const BDR_SUNKENOUTER                       As Long = &H2
Private Const BDR_RAISEDINNER                       As Long = &H4
Private Const BDR_RAISED                            As Long = &H5
Private Const GWL_EXSTYLE                           As Integer = -20
Private Const WS_EX_TOPMOST                         As Long = &H8
Private Const WS_EX_TOOLWINDOW                      As Long = &H80
' *********** FOR COLOR SELECT DIALOG: *************************************************
Private Const CC_ANYCOLOR                           As Long = &H100
Private Const CC_RGBINIT                            As Long = &H1
Private Const CC_FULLOPEN                           As Long = &H2
Private Type CHOOSECOLOR_TYPE
    lStructSize                                         As Long
    hWndOwner                                           As Long
    hInstance                                           As Long
    rgbResult                                           As Long
    lpCustColors                                        As Long
    flags                                               As Long
    lCustData                                           As Long
    lpfnHook                                            As Long
    lpTemplateName                                      As String
End Type
' **************************************************************************************
'---------------------------------------------------------
'Events
'---------------------------------------------------------
Public Event Click()
Public Event TrackColor(ByVal HighlightedColor As Long)
Attribute TrackColor.VB_Description = "Fires when you move mouse over color cells in the dropdown part of the control."
Public Event MouseIn()
Public Event MouseOut()
Public Event DropDownOpen()
Public Event DropDownClose()
Private runOnce                                     As Boolean
Public Enum ColorButtonStyles
    ColorRectAndIconAbove = 0
    ColorRectOnly = 1
End Enum
#If False Then
Private ColorRectAndIconAbove, ColorRectOnly
#End If

Public Enum ColorPalettes
    System = 0  ' << usual color palette
    VGA = 1     ' << 16 color
End Enum
#If False Then
Private System, VGA
#End If
Private Const Def_DropDownCaption                   As String = "Additional colors..."
'---------------------------------------------------------
'Private variables
'---------------------------------------------------------
Private m_iXIndex                                   As Integer
Private m_iYIndex                                   As Integer
Private m_nBackColor                                As OLE_COLOR
Private m_nBorderColor                              As Long
Private m_nFillColor                                As Long
Private m_nDarkFillColor                            As Long
Private m_nShadowColor                              As Long
Private m_nSelectedColor                            As OLE_COLOR
Private m_arrColor()                                As Long
Private m_Style                                     As ColorButtonStyles
Private m_ColorPalette                              As ColorPalettes
Private m_DropDownCaption                           As String
Private m_nForbiddenColor                           As OLE_COLOR
Private m_UseForbiddenColor                         As Boolean
Private m_Step                                      As Long
Private m_RectSize                                  As Long
Private m_ColorsInRow                               As Long
Private m_ColorsInColumn                            As Long
Private m_OffsetTop                                 As Long
Private previousTrackedColor                        As Long
Private mouseIsIn                                   As Boolean
Private isDropped                                   As Boolean
Private flag                                        As Boolean
    ' see picDropDown_MouseDown for details; needed because repeated clicks on arrow area should
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, _
                                                ByVal X1 As Long, _
                                                ByVal Y1 As Long, _
                                                ByVal X2 As Long, _
                                                ByVal Y2 As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, _
                                                qrc As RECT, _
                                                ByVal edge As Long, _
                                                ByVal grfFlags As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
                                                  ByVal xLeft As Long, _
                                                  ByVal yTop As Long, _
                                                  ByVal hIcon As Long, _
                                                  ByVal cxWidth As Long, _
                                                  ByVal cyWidth As Long, _
                                                  ByVal istepIfAniCur As Long, _
                                                  ByVal hbrFlickerFreeDraw As Long, _
                                                  ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, _
                                               ByVal X1 As Long, _
                                               ByVal Y1 As Long, _
                                               ByVal X2 As Long, _
                                               ByVal Y2 As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                 ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, _
                                                               ByVal lHPalette As Long, _
                                                               lColorRef As Long) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpcc As CHOOSECOLOR_TYPE) As Long
'BackColor Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_nBackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    m_nBackColor = NewValue
    Call Redraw
    PropertyChanged ("BackColor")
    picDropDown.BackColor = m_nBackColor
End Property
'Palette Property
Public Property Get ColorPalette() As ColorPalettes
    ColorPalette = m_ColorPalette
End Property
Public Property Let ColorPalette(ByVal NewValue As ColorPalettes)
Dim nW As Single
Dim nH As Single
    
    m_ColorPalette = NewValue
    Call Redraw
    PropertyChanged ("ColorPalette")
' Appearance of picDropDown greatly depends upon this property:
    Call InitColorArray
    m_Step = IIf((m_ColorPalette = System), 18, 20)
    m_RectSize = IIf((m_ColorPalette = System), 12, 16)
    m_ColorsInRow = IIf((m_ColorPalette = System), 8, 4)
    m_ColorsInColumn = IIf((m_ColorPalette = System), 5, 4)
    m_OffsetTop = IIf((m_ColorPalette = System), 32, 7)
    nW = IIf((m_ColorPalette = System), 154, 92)
    nH = IIf((m_ColorPalette = System), 124, 90)
    If runOnce Then
' don't ask me why these two lines are needed :)) remove the whole IF - and you'll get complete
' shit if changing ColorPalette property at run-time!
        nW = nW * Screen.TwipsPerPixelX
        nH = nH * Screen.TwipsPerPixelY
    End If
    picDropDown.Move picDropDown.Left, picDropDown.Top, nW, nH
End Property
Private Sub DrawAllColors()
Dim i  As Integer
Dim j  As Integer
Dim rc As RECT
    
    picDropDown.Cls
'User-defined color
    If m_ColorPalette = System Then
        Call DrawRectangle(picDropDown.hDC, 8, 8, 138, 18, &H808080, , True)
        Call SetRect(rc, 8, 8, 138 + 8, 18 + 8)
        Call DrawText(picDropDown.hDC, m_DropDownCaption, Len(m_DropDownCaption), rc, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE)
    End If
'Selected color
    If m_iXIndex >= 0 And m_iXIndex <= m_ColorsInRow - 1 And m_iYIndex >= 0 And m_iYIndex <= m_ColorsInColumn - 1 Then
        Call DrawSelectedColorBackground(picDropDown.hDC, 8 + m_iXIndex * m_Step - 3, m_OffsetTop + m_iYIndex * m_Step - 3, m_RectSize + 6, m_RectSize + 6)
    End If
'Other colors
    For i = 0 To m_ColorsInColumn - 1
        For j = 0 To m_ColorsInRow - 1
            Call DrawRectangle(picDropDown.hDC, 8 + j * m_Step, m_OffsetTop + i * m_Step, m_RectSize, m_RectSize, &H808080, m_arrColor(i, j))
        Next j
    Next i
'Window border
    Call SetRect(rc, 0, 0, picDropDown.ScaleWidth, picDropDown.ScaleHeight)
    Call DrawEdge(picDropDown.hDC, rc, BDR_RAISED, BF_RECT)
End Sub

Public Sub DrawRectangle(ByVal lngHDC As Long, _
                         ByVal x As Long, _
                         ByVal y As Long, _
                         ByVal cx As Long, _
                         ByVal cy As Long, _
                         Optional ByVal PenColor As Long = 0, _
                         Optional ByVal BrushColor As Long = &HFFFFFF, _
                         Optional ByVal Transparent As Boolean)
Dim hPen   As Long
Dim hBrush As Long
    If cx < 0 Or cy < 0 Then
        Exit Sub
    End If
    hPen = SelectObject(lngHDC, CreatePen(0, 1, PenColor))
    If hPen Then
        If Not Transparent Then
            hBrush = SelectObject(lngHDC, CreateSolidBrush(BrushColor))
        End If
        Call Rectangle(lngHDC, x, y, x + cx, y + cy)
        If Not Transparent Then
            Call DeleteObject(SelectObject(lngHDC, hBrush))
        End If
        Call DeleteObject(SelectObject(lngHDC, hPen))
    End If
End Sub
Private Sub DrawSelectedColor()
    If m_Style = ColorRectAndIconAbove Then
        Call DrawRectangle(UserControl.hDC, 3, 15, 16, 3, VBColorToRGB(m_nSelectedColor), VBColorToRGB(m_nSelectedColor))
        Call DrawIconEx(UserControl.hDC, 3, 1, imgIcon.Picture, 16, 16, 0, 0, DI_NORMAL)
    Else
    'NOT M_STYLE...
        Call DrawRectangle(UserControl.hDC, 3, 3, 17, 16, VBColorToRGB(m_nSelectedColor), VBColorToRGB(m_nSelectedColor))
    End If
    UserControl.PaintPicture imgDropDown.Picture, 26, 10, 5, 3, 0, 0, 5, 3, vbSrcAnd
End Sub

Private Sub DrawSelectedColorBackground(lngHDC As Long, _
                                        x As Long, _
                                        y As Long, _
                                        cx As Long, _
                                        cy As Long)

Dim i      As Long
Dim j      As Long
Dim rc     As RECT
Dim hBrush As Long
    
    hBrush = CreateSolidBrush(&HFFFFFF)
    For i = x To x + cx - 1
        If i Mod 2 = 0 Then
            For j = y + 1 To y + cy - 1 Step 2
                Call SetRect(rc, i, j, i + 1, j + 1)
                Call FillRect(lngHDC, rc, hBrush)
            Next j
        Else
        'NOT I...
            For j = y To y + cy - 1 Step 2
                Call SetRect(rc, i, j, i + 1, j + 1)
                Call FillRect(lngHDC, rc, hBrush)
            Next j
        End If
    Next i
    Call DeleteObject(hBrush)
    Call SetRect(rc, x, y, x + cx, y + cy)
    Call DrawEdge(lngHDC, rc, BDR_SUNKENOUTER, BF_RECT)
End Sub
Public Sub DropDown()
Dim ListTop  As Single
Dim ListLeft As Single
Dim rc       As RECT
Dim i        As Integer
Dim j        As Integer
'Get m_ixIndex and m_iyIndex
    m_iXIndex = -1
    m_iYIndex = -1
    For i = 0 To m_ColorsInColumn - 1
        For j = 0 To m_ColorsInRow - 1
            If m_arrColor(i, j) = m_nSelectedColor Then
                m_iXIndex = j
                m_iYIndex = i
                Exit For
            End If
        Next j
        If m_iXIndex > -1 Or m_iYIndex > -1 Then
            Exit For
        End If
    Next i
'
    Call GetWindowRect(UserControl.hwnd, rc)
    If rc.Bottom < (Screen.Height - picDropDown.Height) / Screen.TwipsPerPixelY Then
        ListTop = rc.Bottom * Screen.TwipsPerPixelY
    Else
    'NOT RC.BOTTOM...
        ListTop = rc.Top * Screen.TwipsPerPixelY - picDropDown.Height
    End If
    If rc.Right < (Screen.Width - picDropDown.Width) / Screen.TwipsPerPixelX Then
        ListLeft = (rc.Left - 1) * Screen.TwipsPerPixelX
    Else
    'NOT RC.RIGHT...
        ListLeft = (rc.Right + 1) * Screen.TwipsPerPixelX - picDropDown.Width
    End If
'
    Call SetWindowLong(picDropDown.hwnd, GWL_EXSTYLE, WS_EX_TOPMOST Or WS_EX_TOOLWINDOW)
    Call SetParent(picDropDown.hwnd, 0)
    If runOnce = False Then
        runOnce = True
    End If
    picDropDown.Move ListLeft, ListTop, picDropDown.Width, picDropDown.Height
'
    picDropDown.Visible = True
    Call DrawAllColors
    Call SetCapture(picDropDown.hwnd)
'
    Call DrawRectangle(UserControl.hDC, 0, 0, 23, 22, m_nBorderColor, m_nFillColor)
    Call DrawRectangle(UserControl.hDC, 22, 0, 13, 22, m_nBorderColor, m_nDarkFillColor)
    Call DrawSelectedColor
    UserControl.Refresh
    RaiseEvent DropDownOpen
    isDropped = True
    previousTrackedColor = -1
End Sub
'DropDownCaption property
Public Property Get DropDownCaption() As String
    DropDownCaption = m_DropDownCaption
End Property
Public Property Let DropDownCaption(ByVal NewValue As String)
    m_DropDownCaption = NewValue
    PropertyChanged ("DropDownCaption")
End Property
'ForbiddenColor property
Public Property Get ForbiddenColor() As OLE_COLOR
Attribute ForbiddenColor.VB_Description = "Defines color that can't be selected from the dropdown part. Has meaning only if UseForbiddenColor = True. TrackColor event isn't fired for such color."
    ForbiddenColor = m_nForbiddenColor
End Property
Public Property Let ForbiddenColor(ByVal NewValue As OLE_COLOR)
    m_nForbiddenColor = NewValue
    Call Redraw
    PropertyChanged ("ForbiddenColor")
End Property

' toogle dropped state of picDropDown - in the original version of this module picDropDown
' just remained dropped, what seems incorrect to me.
Public Function GetColorFromDialog(ByVal ownerHwnd As Long, _
                                   ByVal DefColor As Long) As Long
Dim cc                As CHOOSECOLOR_TYPE
Dim custcols(0 To 15) As Long
Dim c                 As Integer
    
    For c = 240 To 15 Step -15
        custcols((c \ 15) - 1) = RGB(c, c, c)
    Next c
    With cc
        .lStructSize = Len(cc)
' size of the structure
        .hWndOwner = ownerHwnd
' handle of form opening the Choose Color box
        .hInstance = 0
' not needed
        .rgbResult = DefColor
' set default selected color to DefColor
        .lpCustColors = VarPtr(custcols(0))
' pointer to list of custom colors
        .flags = CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN
' allow any color, use rgbResult as default selection, open "Define Custom Colors" part
        .lCustData = 0
' not needed
        .lpfnHook = 0
' not needed
        .lpTemplateName = vbNullString
'<:-) :WARNING: Empty String assignment updated to use vbNullString
'<:-) :PREVIOUS CODE : .lpTemplateName = ""
' not needed
        ChooseColor cc
'<:-) :WARNING: assigned only variable 'retVal' removed.
        GetColorFromDialog = .rgbResult
    End With
    'CC
End Function

Private Function GetColorIndex(ByVal x As Single, _
                               ByVal y As Single, _
                               ByRef XIndex As Integer, _
                               ByRef YIndex As Integer) As Boolean
Dim i As Integer
Dim j As Integer
    For i = 0 To m_ColorsInRow - 1
        If x >= 8 + i * m_Step - 3 And x <= 8 + i * m_Step + m_RectSize + 3 Then
            Exit For
        End If
    Next i
    For j = 0 To m_ColorsInColumn - 1
        If y >= m_OffsetTop + j * m_Step - 3 And y <= m_OffsetTop + j * m_Step + m_RectSize + 3 Then
            Exit For
        End If
    Next j
    If i >= m_ColorsInRow Or j >= m_ColorsInColumn Then
        GetColorIndex = False
    Else
    'NOT I...
        XIndex = i
        YIndex = j
        GetColorIndex = True
    End If
End Function
'Icon property
Public Property Get Icon() As StdPicture
Attribute Icon.VB_Description = "Property has meaning only if Style property is set to ColorRectAndIconAbove."
    Set Icon = imgIcon.Picture
End Property
Public Property Set Icon(ByVal NewValue As StdPicture)
    Set imgIcon.Picture = NewValue
    Call Redraw
    PropertyChanged ("Icon")
End Property
Private Sub InitColorArray()
'Initialize color array
    If ColorPalette = System Then
        ReDim m_arrColor(4, 7)
        m_arrColor(0, 0) = 0
        m_arrColor(0, 1) = 13209
        m_arrColor(0, 2) = 13107
        m_arrColor(0, 3) = 13056
        m_arrColor(0, 4) = 6697728
        m_arrColor(0, 5) = 8388608
        m_arrColor(0, 6) = 10040115
        m_arrColor(0, 7) = 3355443
        m_arrColor(1, 0) = 128
        m_arrColor(1, 1) = 26367
        m_arrColor(1, 2) = 32896
        m_arrColor(1, 3) = 32768
        m_arrColor(1, 4) = 8421376
        m_arrColor(1, 5) = 16711680
        m_arrColor(1, 6) = 10053222
        m_arrColor(1, 7) = 8421504
        m_arrColor(2, 0) = 255
        m_arrColor(2, 1) = 39423
        m_arrColor(2, 2) = 52377
        m_arrColor(2, 3) = 6723891
        m_arrColor(2, 4) = 13421619
        m_arrColor(2, 5) = 16737843
        m_arrColor(2, 6) = 8388736
        m_arrColor(2, 7) = 9868950
        m_arrColor(3, 0) = 16711935
        m_arrColor(3, 1) = 52479
        m_arrColor(3, 2) = 65535
        m_arrColor(3, 3) = 65280
        m_arrColor(3, 4) = 16776960
        m_arrColor(3, 5) = 16763904
        m_arrColor(3, 6) = 6697881
        m_arrColor(3, 7) = 12632256
        m_arrColor(4, 0) = 13408767
        m_arrColor(4, 1) = 10079487
        m_arrColor(4, 2) = 10092543
        m_arrColor(4, 3) = 13434828
        m_arrColor(4, 4) = 16777164
        m_arrColor(4, 5) = 16764057
        m_arrColor(4, 6) = 16751052
        m_arrColor(4, 7) = 16777215
    Else
    'NOT COLORPALETTE...
        ReDim m_arrColor(3, 3)
        m_arrColor(0, 0) = RGB(0, 0, 0)
        m_arrColor(0, 1) = RGB(128, 0, 0)
        m_arrColor(0, 2) = RGB(0, 128, 0)
        m_arrColor(0, 3) = RGB(128, 128, 0)
        m_arrColor(1, 0) = RGB(0, 0, 128)
        m_arrColor(1, 1) = RGB(128, 0, 128)
        m_arrColor(1, 2) = RGB(0, 128, 128)
        m_arrColor(1, 3) = RGB(128, 128, 128)
        m_arrColor(2, 0) = RGB(192, 192, 192)
        m_arrColor(2, 1) = RGB(255, 0, 0)
        m_arrColor(2, 2) = RGB(0, 255, 0)
        m_arrColor(2, 3) = RGB(255, 255, 0)
        m_arrColor(3, 0) = RGB(0, 0, 255)
        m_arrColor(3, 1) = RGB(255, 0, 255)
        m_arrColor(3, 2) = RGB(0, 255, 255)
        m_arrColor(3, 3) = RGB(255, 255, 255)
    End If
End Sub
Private Sub picDropDown_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
Dim rc                 As RECT
Dim i                  As Integer
Dim j                  As Integer
Dim clickedOnArrowArea As Boolean

    Call DrawAllColors
    If x < 0 Or x > picDropDown.ScaleWidth Or y < 0 Or y > picDropDown.ScaleHeight Then
        mouseIsIn = False
'Outside of dropdown window. But where?!
        clickedOnArrowArea = (Abs(y) <= UserControl.Height / Screen.TwipsPerPixelY) And (x <= UserControl.Width / Screen.TwipsPerPixelX) And (x >= 23)
        If clickedOnArrowArea Then
' arrow area was clicked
            flag = False
        Else
        'CLICKEDONARROWAREA = FALSE/0
            flag = True
        End If
        Call Redraw
        Call ReleaseCapture
        picDropDown.Visible = False
        RaiseEvent DropDownClose
        isDropped = False
    Else
    'NOT X...
        If Button = 1 Then
            If x >= 8 And x <= 8 + 138 And y >= 8 And y <= 8 + 18 And m_ColorPalette = System Then
'User-defined color
                If m_ColorPalette = System Then
                    Call SetCapture(picDropDown.hwnd)
                    Call SetRect(rc, 8 - 3, 8 - 3, 8 + 138 + 3, 8 + 18 + 3)
                    Call DrawEdge(picDropDown.hDC, rc, BDR_SUNKENOUTER, BF_RECT)
                    picDropDown.Refresh
                End If
            Else
            'NOT X...
'Other colors
                Call SetCapture(picDropDown.hwnd)
                If GetColorIndex(x, y, i, j) Then
                    If Not (m_UseForbiddenColor And (m_arrColor(j, i) = m_nForbiddenColor)) Then
                        Call SetRect(rc, 8 + i * m_Step - 3, m_OffsetTop + j * m_Step - 3, 8 + i * m_Step + m_RectSize + 3, m_OffsetTop + j * m_Step + m_RectSize + 3)
                        Call DrawEdge(picDropDown.hDC, rc, BDR_SUNKENOUTER, BF_RECT)
                        picDropDown.Refresh
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub picDropDown_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
Dim rc As RECT
Dim i  As Integer
Dim j  As Integer
    Call DrawAllColors
    If x < 0 Or y < 0 Or x > picDropDown.ScaleWidth Or y > picDropDown.ScaleHeight Then
'do nothing
    Else
    'NOT X...
        Call SetCapture(picDropDown.hwnd)
        If x >= 8 And x <= 8 + 138 And y >= 8 And y <= 8 + 18 And m_ColorPalette = System Then
'User-defined color
            If m_ColorPalette = System Then
                Call SetRect(rc, 8 - 3, 8 - 3, 8 + 138 + 3, 8 + 18 + 3)
                Call DrawEdge(picDropDown.hDC, rc, BDR_RAISEDINNER, BF_RECT)
                picDropDown.Refresh
            End If
        Else
        'NOT X...
'Other colors
            If GetColorIndex(x, y, i, j) Then
                If Not (m_UseForbiddenColor And (m_arrColor(j, i) = m_nForbiddenColor)) Then
                    Call SetRect(rc, 8 + i * m_Step - 3, m_OffsetTop + j * m_Step - 3, 8 + i * m_Step + m_RectSize + 3, m_OffsetTop + j * m_Step + m_RectSize + 3)
                    If m_iXIndex = i And m_iYIndex = j Then
                        Call DrawEdge(picDropDown.hDC, rc, BDR_SUNKENOUTER, BF_RECT)
                    Else
                    'NOT M_IXINDEX...
                        If Button = 0 Then
                            Call DrawEdge(picDropDown.hDC, rc, BDR_RAISEDINNER, BF_RECT)
                        ElseIf Button = 1 Then
                        'NOT BUTTON...
                            Call DrawEdge(picDropDown.hDC, rc, BDR_SUNKENOUTER, BF_RECT)
                        End If
                    End If
                    If m_arrColor(j, i) <> previousTrackedColor Then
                        RaiseEvent TrackColor(m_arrColor(j, i))
                        previousTrackedColor = m_arrColor(j, i)
                    End If
                    picDropDown.Refresh
                End If
            End If
        End If
    End If
End Sub
Private Sub picDropDown_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
Dim i As Integer
Dim j As Integer
    
    If x >= 8 And x <= 8 + 138 And y >= 8 And y <= 8 + 18 And m_ColorPalette = System Then
'User-defined color
        If m_ColorPalette = System Then
            m_iXIndex = -1
            m_iYIndex = -1
            Call ReleaseCapture
            picDropDown.Visible = False
            RaiseEvent DropDownClose
            flag = True
            isDropped = False
            m_nSelectedColor = GetColorFromDialog(UserControl.hwnd, m_nSelectedColor)
            RaiseEvent Click
            Call Redraw
        End If
    Else
    'NOT X...
'Other colors
        If GetColorIndex(x, y, i, j) Then
            If Not (m_UseForbiddenColor And (m_arrColor(j, i) = m_nForbiddenColor)) Then
                m_iXIndex = i
                m_iYIndex = j
                m_nSelectedColor = m_arrColor(j, i)
                Call ReleaseCapture
                Call Redraw
                picDropDown.Visible = False
                RaiseEvent DropDownClose
                flag = True
                isDropped = False
                RaiseEvent Click
            End If
        End If
    End If
End Sub
Private Sub Redraw()
    UserControl.Cls
    UserControl.BackColor = m_nBackColor
    Call DrawSelectedColor
    UserControl.Refresh
End Sub
'Style Property
Public Property Get Style() As ColorButtonStyles
Attribute Style.VB_Description = "Indicates whether the button part of the control has color rect only or also with icon above (icon is taken from the Icon property)."
    Style = m_Style
End Property
Public Property Let Style(ByVal NewValue As ColorButtonStyles)
    m_Style = NewValue
    Call Redraw
    PropertyChanged ("Style")
End Property
'UseForbiddenColor property
Public Property Get UseForbiddenColor() As Boolean
Attribute UseForbiddenColor.VB_Description = "If set to True, you can't select cell with ForbiddenColor in the dropdown part."
    UseForbiddenColor = m_UseForbiddenColor
End Property
Public Property Let UseForbiddenColor(ByVal NewValue As Boolean)
    m_UseForbiddenColor = NewValue
    Call Redraw
    PropertyChanged ("UseForbiddenColor")
End Property
Private Sub UserControl_ExitFocus()
'Hide dropdown window
    If picDropDown.Visible Then
        picDropDown.Visible = False
        Call ReleaseCapture
        RaiseEvent DropDownClose
        flag = True
        isDropped = False
    End If
    Call Redraw
End Sub
Private Sub UserControl_Initialize()
    m_nBorderColor = RGB(8, 36, 107)
    m_nFillColor = RGB(181, 190, 214)
    m_nDarkFillColor = RGB(132, 146, 181)
    m_nShadowColor = VBColorToRGB(vbButtonShadow)
    m_nBackColor = vbButtonFace
    m_nForbiddenColor = vbButtonFace
    m_UseForbiddenColor = False
    m_Style = ColorButtonStyles.ColorRectAndIconAbove
    m_ColorPalette = ColorPalettes.System
    m_DropDownCaption = Def_DropDownCaption
    flag = True
End Sub
Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    If x > 0 And x < 23 And y > 0 And y < 23 Then
'Draw Icon
        Call DrawRectangle(UserControl.hDC, 0, 0, 23, 22, m_nBorderColor, m_nDarkFillColor)
        Call DrawSelectedColor
        UserControl.Refresh
    ElseIf x > 23 And x < 35 And y > 0 And y < 23 Then
    'NOT X...
'Draw dropDown arrow
        Call DrawRectangle(UserControl.hDC, 22, 0, 13, 22, m_nBorderColor, m_nDarkFillColor)
        Call DrawSelectedColor
        UserControl.Refresh
    End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
    If Button <> 0 Then
        Exit Sub
    End If
    If x < 0 Or y < 0 Or x > UserControl.ScaleWidth Or y > UserControl.ScaleHeight Then
'MouseLeave
        Call ReleaseCapture
        Call Redraw
        RaiseEvent MouseOut
        mouseIsIn = False
    Else
    'NOT X...
'MouseOver

        With UserControl
            Call SetCapture(.hwnd)
            Call DrawRectangle(.hDC, 0, 0, 23, 22, m_nBorderColor, m_nFillColor)
            Call DrawRectangle(.hDC, 22, 0, 13, 22, m_nBorderColor, m_nFillColor)
        End With
        'UserControl
        Call DrawSelectedColor
        UserControl.Refresh
        If Not mouseIsIn Then
            RaiseEvent MouseIn
            mouseIsIn = True
        End If
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
    If Button <> 1 Then
        Exit Sub
    End If
    If x > 0 And x < 23 And y > 0 And y < 23 Then
'Click Icon
        RaiseEvent Click
    ElseIf x > 23 And x < 35 And y > 0 And y < 23 Then
    'NOT X...
'Click dropdown arrow
        Call ReleaseCapture
        If Not isDropped Then
            If flag Then
                Call DropDown
            Else
            'FLAG = FALSE/0
                flag = True
            End If
        Else
        'NOT NOT...
            Call ReleaseCapture
            Call Redraw
            RaiseEvent DropDownClose
            flag = True
            isDropped = False
        End If
    End If
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Value = .ReadProperty("Value", &H0)
        Set Icon = .ReadProperty("Icon", imgIcon.Picture)
        Style = .ReadProperty("Style", ColorButtonStyles.ColorRectAndIconAbove)
        BackColor = .ReadProperty("BackColor", vbButtonFace)
        ForbiddenColor = .ReadProperty("ForbiddenColor", vbButtonFace)
        UseForbiddenColor = .ReadProperty("UseForbiddenColor", False)
        ColorPalette = .ReadProperty("ColorPalette", ColorPalettes.System)
        DropDownCaption = .ReadProperty("DropDownCaption", Def_DropDownCaption)
    End With
    'PropBag
End Sub
Private Sub UserControl_Resize()
    UserControl.Width = 35 * Screen.TwipsPerPixelX
    UserControl.Height = 22 * Screen.TwipsPerPixelY
    Call Redraw
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Value", Value, &H0
        .WriteProperty "Icon", imgIcon.Picture
        .WriteProperty "Style", Style, ColorButtonStyles.ColorRectAndIconAbove
        .WriteProperty "BackColor", BackColor, vbButtonFace
        .WriteProperty "ForbiddenColor", ForbiddenColor, vbButtonFace
        .WriteProperty "UseForbiddenColor", UseForbiddenColor, False
        .WriteProperty "ColorPalette", ColorPalette, ColorPalettes.System
        .WriteProperty "DropDownCaption", DropDownCaption, Def_DropDownCaption
    End With
    'PropBag
End Sub
'Value property
Public Property Get Value() As OLE_COLOR
    Value = m_nSelectedColor
End Property
Public Property Let Value(ByVal NewValue As OLE_COLOR)
    m_nSelectedColor = NewValue
    Call Redraw
    PropertyChanged ("Value")
End Property
Public Function VBColorToRGB(ByVal VBColor As Long) As Long
    If OleTranslateColor(VBColor, 0, VBColorToRGB) Then
        VBColorToRGB = VBColor
    End If
End Function
