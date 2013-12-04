VERSION 5.00
Begin VB.UserControl ctlProgressBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000040&
   CanGetFocus     =   0   'False
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   ScaleHeight     =   300
   ScaleWidth      =   2970
   ToolboxBitmap   =   "ctlProgressBar.ctx":0000
End
Attribute VB_Name = "ctlProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'progress bar class string
Private Const PROGRESS_CLASSA           As String = "msctls_progress32"

'progress bar messages
Private Const PBM_SETRANGE              As Long = (WM_USER + 1)
Private Const PBM_SETPOS                As Long = (WM_USER + 2)
Private Const PBM_SETRANGE32            As Long = (WM_USER + 6)
Private Const PBM_GETRANGE              As Long = (WM_USER + 7)

'XP
Private Const PBM_SETMARQUEE            As Long = (WM_USER + 10)

'VISTA
Private Const PBM_GETSTEP               As Long = WM_USER + 13
Private Const PBM_GETBKCOLOR            As Long = WM_USER + 14
Private Const PBM_GETBARCOLOR           As Long = WM_USER + 15
Private Const PBM_SETSTATE              As Long = WM_USER + 16
Private Const PBM_GETSTATE              As Long = WM_USER + 17

'progress bar styles
Private Const PBS_SMOOTH                As Long = &H1
Private Const PBS_VERTICAL              As Long = &H4
Private Const PBS_MARQUEE               As Long = &H8
Private Const PBS_SMOOTHREVERSE         As Long = &H10

'progress bar states
'VISTA
Private Const PBST_NORMAL               As Long = &H1
Private Const PBST_ERROR                As Long = &H2
Private Const PBST_PAUSED               As Long = &H3

'progress bar structure
Private Type PPBRANGE
    iLow                                    As Long
    iHigh                               As Long

End Type

'other constants
Private Const RDW_UPDATENOW             As Long = &H100
Private Const RDW_INVALIDATE            As Long = &H1
Private Const RDW_ALLCHILDREN           As Long = &H80
Private Const RDW_ERASE                 As Long = &H4

Public Enum AppearanceConstants
    ccFlat
    cc3D

End Enum

#If False Then

    Private ccFlat, cc3D
#End If

Public Enum BorderStyleConstants
    ccNone
    ccFixedSingle

End Enum

#If False Then

    Private ccNone, ccFixedSingle
#End If

Public Enum ScrollingConstants
    ccScrollingStandard = 0
    ccScrollingSmooth = 1

End Enum

#If False Then

    Private ccScrollingStandard, ccScrollingSmooth
#End If

Public Enum StateConstants
    ccStateNormal = 1
    ccStateError = 2
    ccStatePaused = 3

End Enum

#If False Then

    Private ccStateNormal, ccStateError, ccStatePaused
#End If

'property variables
Private m_Align                         As AlignConstants
Private m_Appearance                    As AppearanceConstants
Private m_BorderStyle                   As BorderStyleConstants
Private m_Scrolling                     As ScrollingConstants
Private m_Value                         As Long
Private m_Max                           As Long
Private m_Min                           As Long
Private m_Step                          As Long
Private m_Marquee                       As Boolean
Private m_State                         As StateConstants

'private vars
Private m_hModShell32                   As Long
Private dwStyle                         As Long
Private pbHwnd                          As Long
Private m_bIsWinXpOrLater               As Boolean

Public Property Get Appearance() As AppearanceConstants
    Appearance = m_Appearance

End Property

Public Property Let Appearance(ByVal New_Value As AppearanceConstants)

    If Not m_Appearance = New_Value Then
        m_Appearance = New_Value
        pvCreate
        PropertyChanged "Appearance"

    End If

End Property

Public Property Get BorderStyle() As BorderStyleConstants
    BorderStyle = m_BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_Value As BorderStyleConstants)

    If Not m_BorderStyle = New_Value Then
        m_BorderStyle = New_Value
        pvCreate
        PropertyChanged "BorderStyle"

    End If

End Property

Public Property Get hWnd() As Long
    hWnd = pbHwnd

End Property

Public Property Get Marquee() As Boolean
    Marquee = m_Marquee

End Property

Public Property Let Marquee(ByVal New_Value As Boolean)

    If Not m_Marquee = New_Value Then
        m_Marquee = New_Value
        PropertyChanged "Marquee"

        If pbHwnd Then
            If Ambient.UserMode Then
                pvCreate
                SendMessageLong pbHwnd, PBM_SETMARQUEE, New_Value, IIf(m_bIsWinXpOrLater, 100, 30)
            End If
        End If

        If Not m_Marquee Then
            Min = m_Min
            Max = m_Max
            Value = m_Value

        End If

    End If

End Property

Public Property Get Max() As Long
    Max = m_Max

End Property

Public Property Let Max(ByVal New_Value As Long)

    If Not m_Max = New_Value Then
        If m_Min > New_Value Then
            If Ambient.UserMode Then
                err.Raise 380, App.EXEName & ".ctlProgressBar"

            End If

        Else
            m_Max = New_Value
            pvSetRange
            PropertyChanged "Max"

        End If

    End If

End Property

Public Property Get Min() As Long
    Min = m_Min

End Property

Public Property Let Min(ByVal New_Value As Long)

    If Not m_Min = New_Value Then
        If New_Value > m_Max Then
            If Ambient.UserMode Then
                err.Raise 380, App.EXEName & ".ctlProgressBar"

            End If

        Else
            m_Min = New_Value
            pvSetRange
            PropertyChanged "Min"

        End If

    End If

End Property

Private Function pvCreate() As Boolean
    pvDestroy

    '    If Not mbInitXPStyle Then
    '        InitXPStyle
    '    End If

    dwStyle = WS_VISIBLE Or WS_CHILD

    If m_Align = vbAlignLeft Or m_Align = vbAlignRight Then
        dwStyle = dwStyle Or PBS_VERTICAL

    End If

    If Scrolling = ccScrollingSmooth Then
        If Ambient.UserMode Then
            dwStyle = dwStyle Or PBS_SMOOTH

        End If

    End If

    If Marquee Then
        If Ambient.UserMode Then
            dwStyle = dwStyle Or PBS_MARQUEE

        End If

    End If

    pbHwnd = CreateWindowEx(0, StrPtr(PROGRESS_CLASSA), StrPtr(vbNullString), dwStyle, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, UserControl.hWnd, 0&, App.hInstance, ByVal 0&)
    UserControl.BackColor = vbButtonFace

    If pbHwnd Then
        pvSetBorder

        If Ambient.UserMode Then
            pvSetRange

            If m_Marquee Then
                SendMessageLong pbHwnd, PBM_SETMARQUEE, True, IIf(m_bIsWinXpOrLater, 100, 30)

            End If

            SendMessage pbHwnd, PBM_SETPOS, m_Value, 0

            If m_State <> ccStateNormal Then
                SendMessageLong pbHwnd, PBM_SETSTATE, m_State, 0

            End If

        Else
            SendMessage pbHwnd, PBM_SETPOS, m_Max, 0

        End If

        Refresh
        pvCreate = True

    End If

End Function

Private Sub pvDestroy()

    If pbHwnd Then
        ShowWindow pbHwnd, SW_HIDE
        SetParent pbHwnd, 0
        DestroyWindow pbHwnd

    End If

End Sub

Private Sub pvSetBorder()

    If m_Appearance = ccFlat Then
        If m_BorderStyle = ccFixedSingle Then
            pvSetStyle WS_BORDER, 0
            pvSetExStyle WS_EX_CLIENTEDGE, 0
        ElseIf m_BorderStyle = ccNone Then
            pvSetExStyle WS_EX_CLIENTEDGE, 0

        End If

    ElseIf m_Appearance = cc3D Then

        If m_BorderStyle = ccFixedSingle Then
            pvSetStyle WS_BORDER, 0
        ElseIf m_BorderStyle = ccNone Then
            pvSetStyle 0, WS_BORDER

        End If

    End If

End Sub

Private Sub pvSetExStyle(ByVal lStyle As Long, ByVal lStyleNot As Long)

Dim lS                                  As Long

    If Not pbHwnd = 0 Then
        lS = GetWindowLong(pbHwnd, GWL_EXSTYLE)
        lS = lS And Not lStyleNot
        lS = lS Or lStyle
        SetWindowLong pbHwnd, GWL_EXSTYLE, lS
        SetWindowPos pbHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

    End If

End Sub

Private Sub pvSetRange()

Dim tPR                                 As PPBRANGE
Dim tPA                                 As PPBRANGE
Dim lR                                  As Long

    If pbHwnd <> 0 Then
        ' try v4.70 PBM_SETRANGE32:
        SendMessageLong pbHwnd, PBM_SETRANGE32, m_Min, m_Max
        ' check whether PBM_SETRANGE32 was supported:
        tPA.iHigh = SendMessage(pbHwnd, PBM_GETRANGE, 0, tPR)
        tPA.iLow = SendMessage(pbHwnd, PBM_GETRANGE, 1, tPR)

        If Not (tPA.iHigh = m_Max) Then
            If Not (tPA.iLow = m_Min) Then
                ' use the original set range message:
                lR = (m_Min And &HFFFF&)
                CopyMemory VarPtr(lR) + 2, (m_Max And &HFFFF&), 2
                SendMessage pbHwnd, PBM_SETRANGE, 0, lR
            End If
        End If

    End If

End Sub

Private Sub pvSetStyle(ByVal lStyle As Long, ByVal lStyleNot As Long)

Dim lS                                  As Long

    lS = GetWindowLong(pbHwnd, GWL_STYLE)
    lS = lS And Not lStyleNot
    lS = lS Or lStyle
    SetWindowLong pbHwnd, GWL_STYLE, lS
    SetWindowPos pbHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

End Sub

Public Sub Refresh()

Dim tR                                  As RECT

    GetWindowRect pbHwnd, tR
    OffsetRect tR, -tR.Left, -tR.Top
    RedrawWindow pbHwnd, tR, 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_ERASE

End Sub

Public Property Get Scrolling() As ScrollingConstants
    Scrolling = m_Scrolling

End Property

Public Property Let Scrolling(ByVal New_Value As ScrollingConstants)

    If Not m_Scrolling = New_Value Then
        m_Scrolling = New_Value

        If pbHwnd Then
            pvCreate

        End If

        PropertyChanged "Scrolling"

    End If

End Property

Public Property Get State() As StateConstants
    State = SendMessageLong(pbHwnd, PBM_GETSTATE, 0, 0)

End Property

Public Property Let State(ByVal New_State As StateConstants)
    SendMessageLong pbHwnd, PBM_SETSTATE, New_State, 0
    m_State = New_State

End Property

Private Sub UserControl_Initialize()

    m_bIsWinXpOrLater = IsWinXPOrLater
    
    m_Appearance = cc3D
    m_Max = 100
    m_Min = 0
    m_Step = 1

End Sub

Private Sub UserControl_InitProperties()
    Appearance = cc3D
    BorderStyle = ccNone
    Scrolling = False
    pvCreate

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    pvCreate

    With PropBag
        Scrolling = .ReadProperty("Scrolling", False)
        Appearance = .ReadProperty("Appearance", cc3D)
        BorderStyle = .ReadProperty("BorderStyle", ccNone)
        Marquee = .ReadProperty("Marquee", False)
        Max = .ReadProperty("Max", 100)
        Min = .ReadProperty("Min", 0)
        Value = .ReadProperty("Value", 0)

    End With

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next

    MoveWindow pbHwnd, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, 1

    If m_Align = UserControl.Extender.Align Then
        Exit Sub

    End If

    m_Align = UserControl.Extender.Align
    pvCreate

End Sub

Private Sub UserControl_Terminate()

    On Error Resume Next

    If pbHwnd Then
        pvDestroy

        'If m_hModShell32 Then FreeLibrary m_hModShell32
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Appearance", m_Appearance, ccFlat
        .WriteProperty "BorderStyle", m_BorderStyle, ccNone
        .WriteProperty "Marquee", m_Marquee, False
        .WriteProperty "Max", m_Max, 100
        .WriteProperty "Min", m_Min, 0
        .WriteProperty "Scrolling", m_Scrolling, False
        .WriteProperty "Value", m_Value, 0

    End With

End Sub

Public Property Get Value() As Long
    Value = m_Value

End Property

Public Property Let Value(ByVal New_Value As Long)

    If New_Value < m_Min Then
        m_Value = m_Min
    ElseIf New_Value > m_Max Then
        'If Ambient.UserMode Then
        'Err.Raise 380, App.EXEName & ".ctlProgressBar"
        'End If
        m_Value = m_Max
    Else
        m_Value = New_Value

    End If

    If Ambient.UserMode Then
        SendMessage pbHwnd, PBM_SETPOS, m_Value, 0

    End If

    PropertyChanged "Value"

End Property

''
