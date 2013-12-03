VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolTipEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'********************************************************************
'*            Написано 10.06.2003 году (Team HomeWork)              *
'*                   e-mail: sne_pro@mail.ru                        *
'********************************************************************
Public Enum gbHWTitleIco
    icoTTNone = &H0
    icoTTInfo = &H1
    icoTTWarning = &H2
    icoTTError = &H3

End Enum

#If False Then

    Private icoTTNone, icoTTInfo, icoTTWarning, icoTTError
#End If

Public Enum TOOLSTYLE
    TTS_ALWAYSTIP = &H1
    TTS_NOPREFIX = &H2
    TTS_NOANIMATE = &H10
    TTS_NOFADE = &H20
    TTS_BALLOON = &H40
    TTS_CLOSE = &H80

End Enum

#If False Then

    Private TTS_ALWAYSTIP, TTS_NOPREFIX, TTS_NOANIMATE, TTS_NOFADE, TTS_BALLOON, TTS_CLOSE
#End If

Public Enum gbHWTTShType
    TTF_IDISHWND = &H1
    TTF_CENTERTIP = &H2
    TTF_SUBCLASS = &H10
    TTF_TRACK = &H20
    TTF_ABSOLUTE = &H80
    TTF_TRANSPARENT = &H100
    TTF_DI_SETITEM = &H8000
    TTF_PARSELINKS = &H1000
    TTF_RTLREADING = &H4

End Enum

#If False Then

    Private TTF_IDISHWND, TTF_CENTERTIP, TTF_SUBCLASS, TTF_TRACK, TTF_ABSOLUTE, TTF_TRANSPARENT, TTF_DI_SETITEM
#End If

Public Enum gbHWTTDelay
    TTDT_AUTOMATIC = &H0
    TTDT_RESHOW = &H1
    TTDT_AUTOPOP = &H2
    TTDT_INITIAL = &H3

End Enum

#If False Then

    Private TTDT_AUTOMATIC, TTDT_RESHOW, TTDT_AUTOPOP, TTDT_INITIAL
#End If

Private Type TOOLINFO
    cbSize                                  As Long
    uFlags                              As gbHWTTShType
    hWnd                                As Long
    uId                                 As Long
    rct                                 As RECT
    hInst                               As Long
    'lpszText                             As String
    lpszText                            As Long
    lParam                              As Long
End Type

Private Const CW_USEDEFAULT             As Long = &H80000000
Private Const WS_POPUP                  As Long = &H80000000
Private Const TTM_SETTITLE              As Long = &H420
Private Const TTM_SETTIPBKCOLOR         As Long = &H413
Private Const TTM_SETTIPTEXTCOLOR       As Long = &H414
Private Const TTM_SETDELAYTIME          As Long = &H403
Private Const TTM_ADDTOOL               As Long = &H404
Private Const TTM_GETMAXTIPWIDTH        As Long = &H419
Private Const TTM_SETMAXTIPWIDTH        As Long = &H418
Private Const TTM_GETTOOLCOUNT          As Long = &H40D
Private Const TTM_UPDATETIPTEXT         As Long = &H40C
Private Const TTM_ENUMTOOLS             As Long = &H40E
Private Const TTM_GETTIPBKCOLOR         As Long = &H416
Private Const TTM_GETDELAYTIME          As Long = &H415
Private Const WM_SETFONT                As Long = &H30
Private Const WM_GETFONT                As Long = &H31

Private hToolTip                        As Long

Public Sub AddToolTip(ByVal CtlhWnd As Long, _
                      ByVal strTTText As String, _
                      Optional ByVal inTTshParam As gbHWTTShType)

Dim ti                                  As TOOLINFO

    With ti
        .uFlags = TTF_SUBCLASS Or TTF_IDISHWND Or inTTshParam Or TTF_ABSOLUTE
        .hWnd = CtlhWnd
        .uId = CtlhWnd
        .hInst = App.hInstance
        '.lpszText = strTTText
        .lpszText = StrPtr(StrConv(strTTText, vbFromUnicode) & vbNullChar)
        Call GetClientRect(CtlhWnd, .rct)
        .cbSize = Len(ti)

    End With

    Call SendMessageW(hToolTip, TTM_ADDTOOL, &H0, ti)

End Sub

' Цвет фона
Public Property Get BackColor() As ColorConstants
    BackColor = SendMessage(hToolTip, TTM_GETTIPBKCOLOR, &H0, &H0)

End Property

Public Property Let BackColor(inColor As ColorConstants)
    Call SendMessage(hToolTip, TTM_SETTIPBKCOLOR, inColor, &H0)

End Property

Private Sub Class_Terminate()

    On Error Resume Next

    Call DestroyToolTip

End Sub

Public Sub CreateToolTip(ByVal FormHandle As Long, _
                         Optional ByVal inTTStyle As TOOLSTYLE = &H0, _
                         Optional ByVal inTTIcon As gbHWTitleIco = icoTTNone, _
                         Optional ByVal strTTTitle As String = vbNullString, _
                         Optional ByVal iTTWidth As Long = &HFFFF, _
                         Optional ByVal inTTDelayTime As Long = &HFFFF, _
                         Optional ByVal inTTVisibleTime As Long = &HFFFF, _
                         Optional ByVal lBackColor As ColorConstants = &HFFFF, _
                         Optional ByVal lForeColor As ColorConstants = &HFFFF)
    hToolTip = CreateWindowEx(WS_EX_TOPMOST, "tooltips_class32", &H0, WS_POPUP Or inTTStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, FormHandle, &H0, App.hInstance, &H0)
    Call SetWindowPos(hToolTip, &HFFFF, &H0, &H0, &H0, &H0, &H2 Or &H1 Or &H10)

    If Not inTTIcon = &H0 Or Len(strTTTitle) Then
        'Call SendMessage(hToolTip, TTM_SETTITLE, inTTIcon, ByVal strTTTile)
        Call SendMessageW(hToolTip, TTM_SETTITLE, inTTIcon, ByVal StrPtr(StrConv(strTTTitle, vbFromUnicode) & vbNullChar))

    End If

    If Not iTTWidth = &HFFFF Then
        Call SendMessage(hToolTip, TTM_SETMAXTIPWIDTH, iTTWidth, &H0)

    End If

    If Not inTTDelayTime = &HFFFF Then
        Call SendMessage(hToolTip, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal inTTDelayTime)

    End If

    If Not inTTVisibleTime = &HFFFF Then
        Call SendMessage(hToolTip, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal inTTVisibleTime)

    End If

    If Not lBackColor = &HFFFF Then
        Call SendMessage(hToolTip, TTM_SETTIPBKCOLOR, lBackColor, ByVal &H0)

    End If

    If Not lForeColor = &HFFFF Then
        Call SendMessage(hToolTip, TTM_SETTIPTEXTCOLOR, lForeColor, ByVal &H0)

    End If

    Call SendMessage(hToolTip, WM_SETFONT, SendMessage(frmMain.hWnd, WM_GETFONT, &H0, &H0), ByVal &H0)

End Sub

' Задержка и ее тип
Public Property Get DelayTime(dwType As gbHWTTDelay) As Long
    DelayTime = SendMessage(hToolTip, TTM_GETDELAYTIME, dwType, &H0)

End Property

Public Property Let DelayTime(dwType As gbHWTTDelay, dwMilliSecs As Long)
    Call SendMessage(hToolTip, TTM_SETDELAYTIME, dwType, ByVal dwMilliSecs)

End Property

Public Sub DestroyToolTip()
    Call DestroyWindow(hToolTip)
    hToolTip = &H0

End Sub

' Цвет обводки и надписей
Public Property Get ForeColor() As ColorConstants
    ForeColor = SendMessage(hToolTip, TTM_SETTIPTEXTCOLOR, &H0, &H0)

End Property

Public Property Let ForeColor(inColor As ColorConstants)
    Call SendMessage(hToolTip, TTM_SETTIPTEXTCOLOR, inColor, &H0)

End Property

Private Function GetToolInfo(ByVal hWndToolTip As Long, ByRef ti As TOOLINFO) As Boolean

Dim li                                  As Long

    ti.cbSize = Len(ti)
    ti.lpszText = String$(81, 0)

    For li = &H0 To ToolTipCount - &H1

        If SendMessage(hToolTip, TTM_ENUMTOOLS, li, ti) Then
            If hWndToolTip = ti.uId Then
                GetToolInfo = True
                Exit Function

            End If

        End If

    Next
    ti.lpszText = TrimNull(ti.lpszText)

End Function

' Максимальная длинна подсказки
Public Property Get MaxToolTipWidth() As Long

    If hToolTip = &H0 Then
        Exit Property

    End If

    MaxToolTipWidth = LoWord(SendMessage(hToolTip, TTM_GETMAXTIPWIDTH, 0, 0))

End Property

Public Property Let MaxToolTipWidth(ByVal CX As Long)

    If hToolTip = &H0 Then
        Exit Property

    End If

    Call SendMessage(hToolTip, TTM_SETMAXTIPWIDTH, &H0, ByVal CLng(CX))

End Property

' Количество зарегистрированных подсказок
Public Property Get ToolTipCount() As Long

    If hToolTip = &H0 Then
        Exit Property

    End If

    ToolTipCount = SendMessage(hToolTip, TTM_GETTOOLCOUNT, 0, 0)

End Property

' Текст подсказки
Public Property Get ToolTipText(CtrlhWnd As Long) As String

Dim ti                                  As TOOLINFO

    If Not GetToolInfo(CtrlhWnd, ti) Then
        Exit Property

    End If

    ToolTipText = ti.lpszText

End Property

Public Property Let ToolTipText(CtrlhWnd As Long, ByVal sText As String)

Dim ti                                  As TOOLINFO

    If Not GetToolInfo(CtrlhWnd, ti) Then
        Exit Property

    End If

    ti.lpszText = sText & vbNullChar
    Call SendMessage(hToolTip, TTM_UPDATETIPTEXT, &H0, ti)

End Property
