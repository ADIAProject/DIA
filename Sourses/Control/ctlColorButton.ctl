VERSION 5.00
Begin VB.UserControl ctlColorButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Tahoma"
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

Private Const BDR_RAISED  As Long = &H5

' *********** FOR COLOR SELECT DIALOG: *************************************************
Private Const CC_ANYCOLOR As Long = &H100
Private Const CC_RGBINIT  As Long = &H1
Private Const CC_FULLOPEN As Long = &H2

Private Type CHOOSECOLOR_TYPE
    lStructSize                             As Long
    hWndOwner                           As Long
    hInstance                           As Long
    RGBResult                           As Long
    lpCustColors                        As Long
    Flags                               As Long
    lCustData                           As Long
    lpfnHook                            As Long
    lpTemplateName                      As String
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

Private runOnce As Boolean

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

Private Const Def_DropDownCaption As String = "Additional colors..."

'---------------------------------------------------------
'Private variables
'---------------------------------------------------------
Private m_iXIndex                 As Integer
Private m_iYIndex                 As Integer
Private m_nBackColor              As OLE_COLOR
Private m_nBorderColor            As Long
Private m_nFillColor              As Long
Private m_nDarkFillColor          As Long
Private m_nShadowColor            As Long
Private m_nSelectedColor          As OLE_COLOR
Private m_arrColor()              As Long
Private m_Style                   As ColorButtonStyles
Private m_ColorPalette            As ColorPalettes
Private m_DropDownCaption         As String
Private m_nForbiddenColor         As OLE_COLOR
Private m_UseForbiddenColor       As Boolean
Private m_Step                    As Long
Private m_RectSize                As Long
Private m_ColorsInRow             As Long
Private m_ColorsInColumn          As Long
Private m_OffsetTop               As Long
Private previousTrackedColor      As Long
Private mouseIsIn                 As Boolean
Private isDropped                 As Boolean
Private Flag                      As Boolean

' see picDropDown_MouseDown for details; needed because repeated clicks on arrow area should
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal XLeft As Long, ByVal YTop As Long, ByVal hIcon As Long, ByVal CXWidth As Long, ByVal CYWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpcc As CHOOSECOLOR_TYPE) As Long

'BackColor Property
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_nBackColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BackColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    m_nBackColor = NewValue
    Redraw
    PropertyChanged ("BackColor")
    picDropDown.BackColor = m_nBackColor
End Property

'Palette Property
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ColorPalette
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ColorPalette() As ColorPalettes
    ColorPalette = m_ColorPalette
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ColorPalette
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (ColorPalettes)
'!--------------------------------------------------------------------------------
Public Property Let ColorPalette(ByVal NewValue As ColorPalettes)

    Dim nW As Single
    Dim nH As Single

    m_ColorPalette = NewValue
    Redraw
    PropertyChanged ("ColorPalette")
    ' Appearance of picDropDown greatly depends upon this property:
    InitColorArray
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawAllColors
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub DrawAllColors()

    Dim i  As Integer
    Dim j  As Integer
    Dim RC As RECT

    picDropDown.Cls

    'User-defined color
    If m_ColorPalette = System Then
        DrawRectangle picDropDown.hDC, 8, 8, 138, 18, &H808080, , True
        SetRect RC, 8, 8, 138 + 8, 18 + 8
        DrawText picDropDown.hDC, m_DropDownCaption, Len(m_DropDownCaption), RC, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE
    End If

    'Selected color
    If m_iXIndex >= 0 Then
        If m_iXIndex <= m_ColorsInRow - 1 Then
            If m_iYIndex >= 0 Then
                If m_iYIndex <= m_ColorsInColumn - 1 Then
                    DrawSelectedColorBackground picDropDown.hDC, 8 + m_iXIndex * m_Step - 3, m_OffsetTop + m_iYIndex * m_Step - 3, m_RectSize + 6, m_RectSize + 6
                End If
            End If
        End If
    End If

    'Other colors
    For i = 0 To m_ColorsInColumn - 1
        For j = 0 To m_ColorsInRow - 1
            DrawRectangle picDropDown.hDC, 8 + j * m_Step, m_OffsetTop + i * m_Step, m_RectSize, m_RectSize, &H808080, m_arrColor(i, j)
        Next
    Next

    'Window border
    SetRect RC, 0, 0, picDropDown.ScaleWidth, picDropDown.ScaleHeight
    DrawEdge picDropDown.hDC, RC, BDR_RAISED, BF_RECT
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawRectangle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lngHDc (Long)
'                              X (Long)
'                              Y (Long)
'                              CX (Long)
'                              CY (Long)
'                              PenColor (Long = 0)
'                              BrushColor (Long = &HFFFFFF)
'                              Transparent (Boolean)
'!--------------------------------------------------------------------------------
Public Sub DrawRectangle(ByVal lngHDc As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, Optional ByVal PenColor As Long = 0, Optional ByVal BrushColor As Long = &HFFFFFF, Optional ByVal Transparent As Boolean)

    Dim hPen   As Long
    Dim hBrush As Long

    If Not CX < 0 Or CY < 0 Then
        hPen = SelectObject(lngHDc, CreatePen(0, 1, PenColor))

        If hPen Then
            If Not Transparent Then
                hBrush = SelectObject(lngHDc, CreateSolidBrush(BrushColor))
            End If

            Rectangle lngHDc, X, Y, X + CX, Y + CY

            If Not Transparent Then
                DeleteObject SelectObject(lngHDc, hBrush)
            End If

            DeleteObject SelectObject(lngHDc, hPen)
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawSelectedColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub DrawSelectedColor()

    If m_Style = ColorRectAndIconAbove Then
        DrawRectangle UserControl.hDC, 3, 15, 16, 3, VBColorToRGB(m_nSelectedColor), VBColorToRGB(m_nSelectedColor)
        DrawIconEx UserControl.hDC, 3, 1, imgIcon.Picture, 16, 16, 0, 0, DI_NORMAL
    Else
        DrawRectangle UserControl.hDC, 3, 3, 17, 16, VBColorToRGB(m_nSelectedColor), VBColorToRGB(m_nSelectedColor)
    End If

    UserControl.PaintPicture imgDropDown.Picture, 26, 10, 5, 3, 0, 0, 5, 3, vbSrcAnd
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DrawSelectedColorBackground
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lngHDc (Long)
'                              X (Long)
'                              Y (Long)
'                              CX (Long)
'                              CY (Long)
'!--------------------------------------------------------------------------------
Private Sub DrawSelectedColorBackground(lngHDc As Long, X As Long, Y As Long, CX As Long, CY As Long)

    Dim i      As Long
    Dim j      As Long
    Dim RC     As RECT
    Dim hBrush As Long

    hBrush = CreateSolidBrush(&HFFFFFF)

    For i = X To X + CX - 1

        If i Mod 2 = 0 Then

            For j = Y + 1 To Y + CY - 1 Step 2
                SetRect RC, i, j, i + 1, j + 1
                FillRect lngHDc, RC, hBrush
            Next

        Else

            For j = Y To Y + CY - 1 Step 2
                SetRect RC, i, j, i + 1, j + 1
                FillRect lngHDc, RC, hBrush
            Next

        End If

    Next

    DeleteObject hBrush
    SetRect RC, X, Y, X + CX, Y + CY
    DrawEdge lngHDc, RC, BDR_SUNKENOUTER, BF_RECT
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DropDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub DropDown()

    Dim ListTop  As Single
    Dim ListLeft As Single
    Dim RC       As RECT
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

        Next

        If m_iXIndex > -1 Or m_iYIndex > -1 Then

            Exit For

        End If

    Next

    '
    GetWindowRect UserControl.hWnd, RC

    If RC.Bottom < (Screen.Height - picDropDown.Height) / Screen.TwipsPerPixelY Then
        ListTop = RC.Bottom * Screen.TwipsPerPixelY
    Else
        ListTop = RC.Top * Screen.TwipsPerPixelY - picDropDown.Height
    End If

    If RC.Right < (Screen.Width - picDropDown.Width) / Screen.TwipsPerPixelX Then
        ListLeft = (RC.Left - 1) * Screen.TwipsPerPixelX
    Else
        ListLeft = (RC.Right + 1) * Screen.TwipsPerPixelX - picDropDown.Width
    End If

    '
    SetWindowLong picDropDown.hWnd, GWL_EXSTYLE, WS_EX_TOPMOST Or WS_EX_TOOLWINDOW
    SetParent picDropDown.hWnd, 0

    If Not runOnce Then
        runOnce = True
    End If

    picDropDown.Move ListLeft, ListTop, picDropDown.Width, picDropDown.Height
    '
    picDropDown.Visible = True
    DrawAllColors
    SetCapture picDropDown.hWnd
    '
    DrawRectangle UserControl.hDC, 0, 0, 23, 22, m_nBorderColor, m_nFillColor
    DrawRectangle UserControl.hDC, 22, 0, 13, 22, m_nBorderColor, m_nDarkFillColor
    DrawSelectedColor
    UserControl.Refresh
    RaiseEvent DropDownOpen
    isDropped = True
    previousTrackedColor = -1
End Sub

'DropDownCaption property
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property DropDownCaption
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get DropDownCaption() As String
    DropDownCaption = m_DropDownCaption
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property DropDownCaption
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (String)
'!--------------------------------------------------------------------------------
Public Property Let DropDownCaption(ByVal NewValue As String)
    m_DropDownCaption = NewValue
    PropertyChanged ("DropDownCaption")
End Property

'ForbiddenColor property
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ForbiddenColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ForbiddenColor() As OLE_COLOR
    ForbiddenColor = m_nForbiddenColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ForbiddenColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let ForbiddenColor(ByVal NewValue As OLE_COLOR)
Attribute ForbiddenColor.VB_Description = "Defines color that can't be selected from the dropdown part. Has meaning only if UseForbiddenColor = True. TrackColor event isn't fired for such color."
    m_nForbiddenColor = NewValue
    Redraw
    PropertyChanged ("ForbiddenColor")
End Property

' toogle dropped state of picDropDown - in the original version of this module picDropDown
' just remained dropped, what seems incorrect to me.
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetColorFromDialog
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   ownerHwnd (Long)
'                              DefColor (Long)
'!--------------------------------------------------------------------------------
Public Function GetColorFromDialog(ByVal ownerHwnd As Long, ByVal DefColor As Long) As Long

    Dim cc                As CHOOSECOLOR_TYPE
    Dim custcols(0 To 15) As Long
    Dim C                 As Integer

    For C = 240 To 15 Step -15
        custcols((C \ 15) - 1) = RGB(C, C, C)
    Next

    With cc
        .lStructSize = Len(cc)
        ' size of the structure
        .hWndOwner = ownerHwnd
        ' handle of form opening the Choose Color box
        .hInstance = 0
        ' not needed
        .RGBResult = DefColor
        ' set default selected color to DefColor
        .lpCustColors = VarPtr(custcols(0))
        ' pointer to list of custom colors
        .Flags = CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN
        ' allow any color, use rgbResult as default selection, open "Define Custom Colors" part
        .lCustData = 0
        ' not needed
        .lpfnHook = 0
        ' not needed
        .lpTemplateName = vbNullString
        ' not needed
        ChooseColor cc
        GetColorFromDialog = .RGBResult
    End With

    'CC
    'CC
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetColorIndex
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   X (Single)
'                              Y (Single)
'                              XIndex (Integer)
'                              YIndex (Integer)
'!--------------------------------------------------------------------------------
Private Function GetColorIndex(ByVal X As Single, ByVal Y As Single, ByRef XIndex As Integer, ByRef YIndex As Integer) As Boolean

    Dim i As Integer
    Dim j As Integer

    For i = 0 To m_ColorsInRow - 1

        If X >= 8 + i * m_Step - 3 Then
            If X <= 8 + i * m_Step + m_RectSize + 3 Then

                Exit For

            End If
        End If

    Next

    For j = 0 To m_ColorsInColumn - 1

        If Y >= m_OffsetTop + j * m_Step - 3 Then
            If Y <= m_OffsetTop + j * m_Step + m_RectSize + 3 Then

                Exit For

            End If
        End If

    Next

    If i >= m_ColorsInRow Or j >= m_ColorsInColumn Then
        GetColorIndex = False
    Else
        XIndex = i
        YIndex = j
        GetColorIndex = True
    End If

End Function

'Icon property
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Icon
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Icon() As StdPicture
    Set Icon = imgIcon.Picture
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Icon
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (StdPicture)
'!--------------------------------------------------------------------------------
Public Property Set Icon(ByVal NewValue As StdPicture)
Attribute Icon.VB_Description = "Property has meaning only if Style property is set to ColorRectAndIconAbove."
    Set imgIcon.Picture = NewValue
    Redraw
    PropertyChanged ("Icon")
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub InitColorArray
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub picDropDown_MouseDown
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub picDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim RC                 As RECT
    Dim i                  As Integer
    Dim j                  As Integer
    Dim clickedOnArrowArea As Boolean

    DrawAllColors

    If X < 0 Or X > picDropDown.ScaleWidth Or Y < 0 Or Y > picDropDown.ScaleHeight Then
        mouseIsIn = False
        'Outside of dropdown window. But where?!
        clickedOnArrowArea = (Abs(Y) <= UserControl.Height / Screen.TwipsPerPixelY) And (X <= UserControl.Width / Screen.TwipsPerPixelX) And (X >= 23)
        Flag = Not (clickedOnArrowArea)
        ' arrow area was clicked
        Redraw
        ReleaseCapture
        picDropDown.Visible = False
        RaiseEvent DropDownClose
        isDropped = False
    Else

        If Button = 1 Then
            If X >= 8 And X <= 8 + 138 And Y >= 8 And Y <= 8 + 18 And m_ColorPalette = System Then

                'User-defined color
                If m_ColorPalette = System Then
                    SetCapture picDropDown.hWnd
                    SetRect RC, 8 - 3, 8 - 3, 8 + 138 + 3, 8 + 18 + 3
                    DrawEdge picDropDown.hDC, RC, BDR_SUNKENOUTER, BF_RECT
                    picDropDown.Refresh
                End If

            Else
                'Other colors
                SetCapture picDropDown.hWnd

                If GetColorIndex(X, Y, i, j) Then
                    If Not (m_UseForbiddenColor And (m_arrColor(j, i) = m_nForbiddenColor)) Then
                        SetRect RC, 8 + i * m_Step - 3, m_OffsetTop + j * m_Step - 3, 8 + i * m_Step + m_RectSize + 3, m_OffsetTop + j * m_Step + m_RectSize + 3
                        DrawEdge picDropDown.hDC, RC, BDR_SUNKENOUTER, BF_RECT
                        picDropDown.Refresh
                    End If
                End If
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub picDropDown_MouseMove
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub picDropDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim RC As RECT
    Dim i  As Integer
    Dim j  As Integer

    DrawAllColors

    If X < 0 Or Y < 0 Or X > picDropDown.ScaleWidth Or Y > picDropDown.ScaleHeight Then
        'do nothing
    Else
        SetCapture picDropDown.hWnd

        If X >= 8 And X <= 8 + 138 And Y >= 8 And Y <= 8 + 18 And m_ColorPalette = System Then

            'User-defined color
            If m_ColorPalette = System Then
                SetRect RC, 8 - 3, 8 - 3, 8 + 138 + 3, 8 + 18 + 3
                DrawEdge picDropDown.hDC, RC, BDR_RAISEDINNER, BF_RECT
                picDropDown.Refresh
            End If

        Else

            'Other colors
            If GetColorIndex(X, Y, i, j) Then
                If Not (m_UseForbiddenColor And (m_arrColor(j, i) = m_nForbiddenColor)) Then
                    SetRect RC, 8 + i * m_Step - 3, m_OffsetTop + j * m_Step - 3, 8 + i * m_Step + m_RectSize + 3, m_OffsetTop + j * m_Step + m_RectSize + 3

                    If m_iXIndex = i And m_iYIndex = j Then
                        DrawEdge picDropDown.hDC, RC, BDR_SUNKENOUTER, BF_RECT
                    Else

                        If Button = 0 Then
                            DrawEdge picDropDown.hDC, RC, BDR_RAISEDINNER, BF_RECT
                        ElseIf Button = 1 Then
                            DrawEdge picDropDown.hDC, RC, BDR_SUNKENOUTER, BF_RECT
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub picDropDown_MouseUp
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub picDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer

    If X >= 8 And X <= 8 + 138 And Y >= 8 And Y <= 8 + 18 And m_ColorPalette = System Then

        'User-defined color
        If m_ColorPalette = System Then
            m_iXIndex = -1
            m_iYIndex = -1
            ReleaseCapture
            picDropDown.Visible = False
            RaiseEvent DropDownClose
            Flag = True
            isDropped = False
            m_nSelectedColor = GetColorFromDialog(UserControl.hWnd, m_nSelectedColor)
            RaiseEvent Click
            Redraw
        End If

    Else

        'Other colors
        If GetColorIndex(X, Y, i, j) Then
            If Not (m_UseForbiddenColor And (m_arrColor(j, i) = m_nForbiddenColor)) Then
                m_iXIndex = i
                m_iYIndex = j
                m_nSelectedColor = m_arrColor(j, i)
                ReleaseCapture
                Redraw
                picDropDown.Visible = False
                RaiseEvent DropDownClose
                Flag = True
                isDropped = False
                RaiseEvent Click
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Redraw
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Redraw()
    UserControl.Cls
    UserControl.BackColor = m_nBackColor
    DrawSelectedColor
    UserControl.Refresh
End Sub

'Style Property
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Style
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Style() As ColorButtonStyles
    Style = m_Style
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Style
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (ColorButtonStyles)
'!--------------------------------------------------------------------------------
Public Property Let Style(ByVal NewValue As ColorButtonStyles)
Attribute Style.VB_Description = "Indicates whether the button part of the control has color rect only or also with icon above (icon is taken from the Icon property)."
    m_Style = NewValue
    Redraw
    PropertyChanged ("Style")
End Property

'UseForbiddenColor property
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseForbiddenColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get UseForbiddenColor() As Boolean
Attribute UseForbiddenColor.VB_Description = "If set to True, you can't select cell with ForbiddenColor in the dropdown part."
    UseForbiddenColor = m_UseForbiddenColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseForbiddenColor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let UseForbiddenColor(ByVal NewValue As Boolean)
    m_UseForbiddenColor = NewValue
    Redraw
    PropertyChanged ("UseForbiddenColor")
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_ExitFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_ExitFocus()

    'Hide dropdown window
    If picDropDown.Visible Then
        picDropDown.Visible = False
        ReleaseCapture
        RaiseEvent DropDownClose
        Flag = True
        isDropped = False
    End If

    Redraw
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Initialize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
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
    Flag = True
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

    If Not Button <> 1 Then
        If X > 0 And X < 23 And Y > 0 And Y < 23 Then
            'Draw Icon
            DrawRectangle UserControl.hDC, 0, 0, 23, 22, m_nBorderColor, m_nDarkFillColor
            DrawSelectedColor
            UserControl.Refresh
        ElseIf X > 23 And X < 35 And Y > 0 And Y < 23 Then
            'Draw dropDown arrow
            DrawRectangle UserControl.hDC, 22, 0, 13, 22, m_nBorderColor, m_nDarkFillColor
            DrawSelectedColor
            UserControl.Refresh
        End If
    End If

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

    If Not Button <> 0 Then
        If X < 0 Or Y < 0 Or X > UserControl.ScaleWidth Or Y > UserControl.ScaleHeight Then
            'MouseLeave
            ReleaseCapture
            Redraw
            RaiseEvent MouseOut
            mouseIsIn = False
        Else

            'MouseOver
            With UserControl
                SetCapture .hWnd
                DrawRectangle .hDC, 0, 0, 23, 22, m_nBorderColor, m_nFillColor
                DrawRectangle .hDC, 22, 0, 13, 22, m_nBorderColor, m_nFillColor
            End With

            DrawSelectedColor
            UserControl.Refresh

            If Not mouseIsIn Then
                RaiseEvent MouseIn
                mouseIsIn = True
            End If
        End If
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

    If Not Button <> 1 Then
        If X > 0 And X < 23 And Y > 0 And Y < 23 Then
            'Click Icon
            RaiseEvent Click
        ElseIf X > 23 And X < 35 And Y > 0 And Y < 23 Then
            'Click dropdown arrow
            ReleaseCapture

            If Not isDropped Then
                If Flag Then
                    DropDown
                Else
                    Flag = True
                End If

            Else
                ReleaseCapture
                Redraw
                RaiseEvent DropDownClose
                Flag = True
                isDropped = False
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_ReadProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropBag (PropertyBag)
'!--------------------------------------------------------------------------------
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

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Resize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Resize()

    On Error Resume Next

    UserControl.Width = 35 * Screen.TwipsPerPixelX
    UserControl.Height = 22 * Screen.TwipsPerPixelY
    Redraw
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_WriteProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PropBag (PropertyBag)
'!--------------------------------------------------------------------------------
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

End Sub

'Value property
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Value
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Value() As OLE_COLOR
    Value = m_nSelectedColor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Value
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let Value(ByVal NewValue As OLE_COLOR)
    m_nSelectedColor = NewValue
    Redraw
    PropertyChanged ("Value")
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function VBColorToRGB
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   VBColor (Long)
'!--------------------------------------------------------------------------------
Public Function VBColorToRGB(ByVal VBColor As Long) As Long

    If OleTranslateColorByRef(VBColor, 0, VBColorToRGB) Then
        VBColorToRGB = VBColor
    End If

End Function
