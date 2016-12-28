VERSION 5.00
Begin VB.UserControl ctlScrollControl 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   MaskPicture     =   "ctlScrollControl.ctx":0000
   MouseIcon       =   "ctlScrollControl.ctx":0152
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ctlScrollControl.ctx":02A4
End
Attribute VB_Name = "ctlScrollControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within Drivers Installer Assistant.
'This code was originally written by Leandro I. Ascierto
'You may download the original version of this code from the following link (good as of 21 Mar '10):
'http://leandroascierto.com/blog/scrollcontrol/
'---------------------------------------------------------

'*********************************
' Modified by Romeo91 (adia-project.net) Last Edit 2015-11-15
'*********************************
' Change subsclasser to class cSelfSubHookCallback
' Added ScrollPositionH property(Get/Let)
' Added ScrollVChanged Event

Option Explicit

Private Declare Function SetScrollInfo Lib "user32.dll" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32.dll" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function ScrollWindowByNum Lib "user32.dll" Alias "ScrollWindow" (ByVal hWnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, ByVal lpRect As Long, ByVal lpClipRect As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetFocus Lib "user32.dll" () As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRECT As RECT, pClipRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type

Private Type POINTAPI
    X                           As Long
    Y                           As Long
End Type

Private Type SCROLLINFO
    cbSize                      As Long
    fMask                       As Long
    nMin                        As Long
    nMax                        As Long
    nPage                       As Long
    nPos                        As Long
    nTrackPos                   As Long
End Type

Private Const GW_CHILD          As Long = 5
Private Const GW_HWNDNEXT       As Long = 2
Private Const SB_HORZ           As Long = 0
Private Const SB_VERT           As Long = 1
Private Const SB_LINEDOWN       As Long = 1
Private Const SB_LINEUP         As Long = 0
Private Const SB_PAGEDOWN       As Long = 3
Private Const SB_PAGEUP         As Long = 2
Private Const SB_THUMBTRACK     As Long = 5
Private Const SB_ENDSCROLL      As Long = 8
Private Const SB_LEFT           As Long = 6
Private Const SB_RIGHT          As Long = 7
Private Const SIF_ALL           As Long = &H17
Private Const SM_CYBORDER       As Long = 6

Public Enum EnuBorderStyle
    vbBSNone
    vbFixedSingle
End Enum

Private SI                      As SCROLLINFO
Private mBorderSize             As Long
Private OldPosH                 As Long
Private OldPosV                 As Long
Private m_hFocus                As Long
Private m_AutoScrollToFocus     As Boolean
Private m_UseHandsCursor        As Boolean
Private m_HScrollVisible        As Boolean
Private m_VScrollVisible        As Boolean
Private m_ScrollPositionH       As Long

'*************************************************************
'   Windows Messages
'*************************************************************
Private Const WM_THEMECHANGED   As Long = &H31A
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_NCPAINT        As Long = &H85
Private Const WM_EXITSIZEMOVE   As Long = &H232
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_VSCROLL        As Long = &H115
Private Const WM_HSCROLL        As Long = &H114
Private Const WM_DESTROY        As Long = &H2

'*************************************************************
'   TRACK MOUSE
'*************************************************************
Public Event ScrollVChanged()

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
Private mY As Single
Private mX As Single

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property AutoScrollToFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get AutoScrollToFocus() As Boolean
    AutoScrollToFocus = m_AutoScrollToFocus
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property AutoScrollToFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let AutoScrollToFocus(ByVal NewValue As Boolean)
    m_AutoScrollToFocus = NewValue
    PropertyChanged "AutoScrollToFocus"
End Property

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
'! Parameters  (Переменные):   NewValue (OLE_COLOR)
'!--------------------------------------------------------------------------------
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    UserControl.BackColor = NewValue
    PropertyChanged "BackColor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BorderStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get BorderStyle() As EnuBorderStyle
    BorderStyle = UserControl.BorderStyle
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property BorderStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (EnuBorderStyle)
'!--------------------------------------------------------------------------------
Public Property Let BorderStyle(ByVal NewValue As EnuBorderStyle)
    UserControl.BorderStyle = NewValue
    PropertyChanged "BorderStyle"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Enabled
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Enabled
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Value (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let Enabled(ByVal Value As Boolean)
    UserControl.Enabled = Value
    PropertyChanged "Enabled"
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
'! Procedure   (Функция)   :   Property ScrollPositionH
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get ScrollPositionH() As Long
    ScrollPositionH = m_ScrollPositionH
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property ScrollPositionH
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let ScrollPositionH(ByVal NewValue As Long)
    If NewValue <> m_ScrollPositionH Then
        m_ScrollPositionH = NewValue
        SI.nPos = NewValue
        SetScrollInfo UserControl.hWnd, SB_VERT, SI, True
        If NewValue = 0 Then
            ScrollVerticalWindow 0
        End If
        PropertyChanged "ScrollPositionH"
    End If
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseHandsCursor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get UseHandsCursor() As Boolean
    UseHandsCursor = m_UseHandsCursor
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property UseHandsCursor
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewValue (Boolean)
'!--------------------------------------------------------------------------------
Public Property Let UseHandsCursor(ByVal NewValue As Boolean)
    m_UseHandsCursor = NewValue
    PropertyChanged "UseHandsCursor"
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CheckScroll
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub CheckScroll()

    On Error Resume Next

    Dim bWnd As Long
    Dim Rec  As RECT
    Dim mRec As RECT

    bWnd = GetWindow(UserControl.hWnd, GW_CHILD)

    Do While bWnd <> 0
        Call GetChildRectOfMe(bWnd, Rec)

        If Rec.Left < mRec.Left Then mRec.Left = Rec.Left
        If Rec.Top < mRec.Top Then mRec.Top = Rec.Top
        If Rec.Right > mRec.Right Then mRec.Right = Rec.Right
        If Rec.Bottom > mRec.Bottom Then mRec.Bottom = Rec.Bottom
        bWnd = GetWindow(bWnd, GW_HWNDNEXT)
    Loop

    If mRec.Bottom + Abs(mRec.Top) > UserControl.ScaleHeight Or (mRec.Top < 0) Then
        GetScrollInfo UserControl.hWnd, SB_VERT, SI
        SI.nMin = mRec.Top
        SI.nMax = mRec.Bottom
        SI.nPage = UserControl.ScaleHeight
        SetScrollInfo UserControl.hWnd, SB_VERT, SI, True
        GetScrollInfo UserControl.hWnd, SB_VERT, SI
        ScrollVerticalWindow -SI.nPos
        m_VScrollVisible = True
    Else
        SI.nPos = 0
        SI.nPage = 0
        SI.nMax = 0
        SI.nMin = 0
        SetScrollInfo UserControl.hWnd, SB_VERT, SI, True
        ScrollVerticalWindow 0
        m_VScrollVisible = False
    End If

    If mRec.Right + Abs(mRec.Left) > UserControl.ScaleWidth Or (mRec.Left < 0) Then
        GetScrollInfo UserControl.hWnd, SB_HORZ, SI
        SI.nMin = mRec.Left
        SI.nMax = mRec.Right
        SI.nPage = UserControl.ScaleWidth
        SetScrollInfo UserControl.hWnd, SB_HORZ, SI, True
        GetScrollInfo UserControl.hWnd, SB_HORZ, SI
        ScrollHorizontalWindow -SI.nPos
        m_HScrollVisible = True
    Else
        SI.nPos = 0
        SI.nPage = 0
        SI.nMax = 0
        SI.nMin = 0
        SetScrollInfo UserControl.hWnd, SB_HORZ, SI, True
        ScrollHorizontalWindow 0
        m_HScrollVisible = False
    End If

    SetWindowPos UserControl.hWnd, 0, 0, 0, 0, 0, 551
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetChildRectOfMe
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   hWnd (Long)
'                              SrcRect (RECT)
'!--------------------------------------------------------------------------------
Private Function GetChildRectOfMe(hWnd As Long, ByRef srcRect As RECT)

    Dim PT As POINTAPI

    ClientToScreen UserControl.hWnd, PT
    Call GetWindowRect(hWnd, srcRect)

    With srcRect
        .Left = .Left - PT.X - OldPosH
        .Top = .Top - PT.Y - OldPosV
        .Right = .Right - PT.X - OldPosH
        .Bottom = .Bottom - PT.Y - OldPosV
    End With

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsChildOfMe
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   hWnd (Long)
'!--------------------------------------------------------------------------------
Private Function IsChildOfMe(hWnd As Long) As Boolean

    Dim hParent As Long

    hParent = GetParent(hWnd)

    Do While hParent <> 0

        If hParent = UserControl.hWnd Then
            IsChildOfMe = True

            Exit Do

        End If

        hParent = GetParent(hParent)
    Loop

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Refresh
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Refresh()
    m_hFocus = 0
    CheckScroll
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ScrollHorizontalWindow
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewPos (Long)
'!--------------------------------------------------------------------------------
Private Sub ScrollHorizontalWindow(ByVal NewPos As Long)
    If NewPos <> OldPosH Then
        ScrollWindowByNum UserControl.hWnd, NewPos - OldPosH, 0&, 0&, 0&
        OldPosH = NewPos
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ScrollVerticalWindow
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewPos (Long)
'!--------------------------------------------------------------------------------
Private Sub ScrollVerticalWindow(ByVal NewPos As Long)
    If NewPos <> OldPosV Then
        ScrollWindowByNum UserControl.hWnd, 0&, NewPos - OldPosV, 0&, 0&
        OldPosV = NewPos
        If m_VScrollVisible Then
            RaiseEvent ScrollVChanged
        End If
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Initialize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Initialize()

    SI.cbSize = Len(SI)
    SI.fMask = SIF_ALL
    mBorderSize = GetSystemMetrics(SM_CYBORDER)
    
    Set m_cSubclass = New cSelfSubHookCallback
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_InitProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_InitProperties()
    m_AutoScrollToFocus = True
    m_UseHandsCursor = True
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

    If m_UseHandsCursor Then
        If Button = 1 Then
            If m_VScrollVisible Or m_HScrollVisible Then
                SetCursor UserControl.MaskPicture
            End If
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

    If m_UseHandsCursor = False Then

        Exit Sub

    End If

    If Button = 1 Then
        If m_VScrollVisible Then
            GetScrollInfo UserControl.hWnd, SB_VERT, SI
            SI.nPos = -(Y - mY)
            SetScrollInfo UserControl.hWnd, SB_VERT, SI, True
            GetScrollInfo UserControl.hWnd, SB_VERT, SI
            ScrollVerticalWindow -SI.nPos
        End If

        If m_HScrollVisible Then
            GetScrollInfo UserControl.hWnd, SB_HORZ, SI
            SI.nPos = -(X - mX)
            SetScrollInfo UserControl.hWnd, SB_HORZ, SI, True
            GetScrollInfo UserControl.hWnd, SB_HORZ, SI
            ScrollHorizontalWindow -SI.nPos
        End If

    Else

        If m_VScrollVisible Then
            GetScrollInfo UserControl.hWnd, SB_VERT, SI
            mY = Y + SI.nPos
        End If

        If m_HScrollVisible Then
            GetScrollInfo UserControl.hWnd, SB_HORZ, SI
            mX = X + SI.nPos
        End If
    End If

    If m_VScrollVisible Or m_HScrollVisible Then
        If Button = 1 Then
            SetCursor UserControl.MaskPicture
        Else

            If Button = 0 Then SetCursor UserControl.MouseIcon
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

    If m_UseHandsCursor Then
        If Button = 1 Then
            If m_VScrollVisible Or m_HScrollVisible Then
                SetCursor UserControl.MouseIcon
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
        Me.BackColor = .ReadProperty("BackColor", vbButtonFace)
        Me.BorderStyle = .ReadProperty("BorderStyle", vbFixedSingle)
        m_AutoScrollToFocus = .ReadProperty("AutoScrollToFocus", True)
        m_UseHandsCursor = .ReadProperty("UseHandsCursor", True)
        Me.Enabled = .ReadProperty("Enabled", True)
    End With

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
                    .ssc_AddMsg UserControl.hWnd, MSG_AFTER, WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL, WM_NCPAINT, WM_THEMECHANGED, WM_SYSCOLORCHANGE
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

    On Error Resume Next

    CheckScroll
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Show
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Show()
    Me.Refresh
    CheckScroll
End Sub


'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UserControl_Terminate
'! Description (Описание)  :   [The control is terminating - a good place to stop the subclasser]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub UserControl_Terminate()

    On Error Resume Next

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
        .WriteProperty "BackColor", UserControl.BackColor, vbButtonFace
        .WriteProperty "BorderStyle", UserControl.BorderStyle, vbFixedSingle
        .WriteProperty "AutoScrollToFocus", m_AutoScrollToFocus, True
        .WriteProperty "UseHandsCursor", m_UseHandsCursor, True
        .WriteProperty "Enabled", UserControl.Enabled, True
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

        Case WM_VSCROLL, WM_HSCROLL

            Dim xScroll As Long

            xScroll = IIf(uMsg = WM_VSCROLL, SB_VERT, SB_HORZ)
            GetScrollInfo hWnd, xScroll, SI

            Select Case LoWord(wParam)

                Case SB_LINEDOWN
                    SI.nPos = SI.nPos + 10

                Case SB_LINEUP
                    SI.nPos = SI.nPos - 10

                Case SB_PAGEDOWN
                    SI.nPos = SI.nPos + SI.nPage

                Case SB_PAGEUP
                    SI.nPos = SI.nPos - SI.nPage

                Case SB_THUMBTRACK
                    SI.nPos = HiWord(wParam)

                Case SB_ENDSCROLL

                Case SB_LEFT
                    SI.nPos = SI.nMin

                Case SB_RIGHT
                    SI.nPos = SI.nMax
            End Select

            SetScrollInfo hWnd, xScroll, SI, True
            GetScrollInfo hWnd, xScroll, SI

            m_ScrollPositionH = SI.nPos
            
            If uMsg = WM_VSCROLL Then
                ScrollVerticalWindow -SI.nPos
            Else
                ScrollHorizontalWindow -SI.nPos
            End If
            Me.Refresh

        Case WM_MOUSEWHEEL

            If m_VScrollVisible Then
                xScroll = SB_VERT
            Else

                If m_HScrollVisible Then
                    xScroll = SB_HORZ
                Else

                    Exit Sub

                End If
            End If

            GetScrollInfo hWnd, xScroll, SI

            If wParam < 0 Then
                SI.nPos = SI.nPos + 10
            Else
                SI.nPos = SI.nPos - 10
            End If

            SetScrollInfo hWnd, xScroll, SI, True
            GetScrollInfo hWnd, xScroll, SI

            m_ScrollPositionH = SI.nPos
            
            If xScroll = SB_VERT Then
                ScrollVerticalWindow -SI.nPos
            Else
                ScrollHorizontalWindow -SI.nPos
            End If

        Case WM_NCPAINT

            If UserControl.BorderStyle = vbFixedSingle Then

                Dim Rec     As RECT
                Dim ClipRec As RECT
                Dim hTheme  As Long
                Dim DC      As Long

                DC = GetWindowDC(hWnd)
                GetWindowRect UserControl.hWnd, Rec
                Rec.Right = Rec.Right - Rec.Left
                Rec.Bottom = Rec.Bottom - Rec.Top
                Rec.Left = 0
                Rec.Top = 0
                hTheme = OpenThemeData(UserControl.hWnd, StrPtr("Edit"))

                If hTheme Then
                    ExcludeClipRect DC, mBorderSize, mBorderSize, Rec.Right - mBorderSize, Rec.Bottom - mBorderSize

                    Call DrawThemeBackground(hTheme, DC, 0, 0, Rec, Rec)
                    Call CloseThemeData(hTheme)
                    
                End If

                ReleaseDC hWnd, DC
            End If

        Case Else

            On Error Resume Next

            Dim hFocus As Long

            If m_AutoScrollToFocus = False Then

                Exit Sub

            End If

            hFocus = GetFocus

            If Not hFocus = m_hFocus Then
                If IsChildOfMe(hFocus) Then
                    m_hFocus = hFocus
                    Call GetChildRectOfMe(hFocus, Rec)
                    GetScrollInfo UserControl.hWnd, SB_VERT, SI

                    If Rec.Bottom > SI.nPos + SI.nPage Then
                        SI.nPos = Rec.Bottom - SI.nPage
                    Else

                        If Rec.Top < SI.nPos Then
                            SI.nPos = Rec.Top
                        End If
                    End If

                    SetScrollInfo UserControl.hWnd, SB_VERT, SI, True
                    GetScrollInfo UserControl.hWnd, SB_HORZ, SI

                    If Rec.Right > SI.nPos + SI.nPage Then
                        SI.nPos = Rec.Right - SI.nPage
                    Else

                        If Rec.Left < SI.nPos Then
                            SI.nPos = Rec.Left
                        End If
                    End If

                    SetScrollInfo UserControl.hWnd, SB_HORZ, SI, True
                    CheckScroll
                End If
            End If

    End Select

End Sub

