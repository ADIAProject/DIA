VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cITaskBarList3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Copyright 2010 Makarov Andrey
'
Option Explicit

Private Declare Function RegisterWindowMessage _
                          Lib "user32.dll" _
                              Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Declare Function DwmSetWindowAttribute _
                          Lib "dwmapi.dll" (ByVal hWnd As Long, _
                                            ByVal dwAttribute As Long, _
                                            ByRef pvAttribute As Long, _
                                            ByVal Length As Long) As Long

Private Const THBN_CLICKED              As Long = &H1800
Private Const WM_DWMSENDICONICTHUMBNAIL As Long = &H323

Private Declare Sub CopyMemory _
                     Lib "kernel32.dll" _
                         Alias "RtlMoveMemory" (lpDest As Any, _
                                                lpSource As Any, _
                                                ByVal cBytes&)

Private Declare Function SetWindowLong _
                          Lib "user32.dll" _
                              Alias "SetWindowLongW" (ByVal hWnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Const AsmMain                   As String = "558BEC83C4FC8D45FC50FF7514FF7510FF750CFF75086800000000B800000000FFD08B45FCC9C21000"

Private ASMArr()                        As Byte
Private oldproc As Long, wnd            As Long
Attribute wnd.VB_VarUserMemId = 1073938433

Private Const CLSID_TaskbarList         As String = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
Private Const IID_ITaskbarList3         As String = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"

Private Enum ITaskbarList3Members
    '/* ITaskbarList methods */
HrInit_ = 3                 'STDMETHOD( HrInit )( THIS ) PURE;
AddTab_ = 4                 'STDMETHOD( AddTab )( THIS_ HWND ) PURE;
DeleteTab_ = 5              'STDMETHOD( DeleteTab )( THIS_ HWND ) PURE;
ActivateTab_ = 6            'STDMETHOD( ActivateTab )( THIS_ HWND ) PURE;
SetActiveAlt_ = 7           'STDMETHOD( SetActiveAlt )( THIS_ HWND ) PURE;
    '/* ITaskbarList2 methods */
MarkFullscreenWindow_ = 8   'STDMETHOD( MarkFullscreenWindow )( THIS_ HWND, BOOL ) PURE;
    '/* ITaskbarList3 methods */
SetProgressValue_ = 9       'STDMETHOD( SetProgressValue )( THIS_ HWND, ULONGLONG, ULONGLONG ) PURE;
SetProgressState_ = 10      'STDMETHOD( SetProgressState )( THIS_ HWND, TBPFLAG ) PURE;
RegisterTab_ = 11           'STDMETHOD( RegisterTab )( THIS_ HWND, HWND ) PURE;
UnregisterTab_ = 12         'STDMETHOD( UnregisterTab )( THIS_ HWND ) PURE;
SetTabOrder_ = 13           'STDMETHOD( SetTabOrder )( THIS_ HWND, HWND ) PURE;
SetTabActive_ = 14          'STDMETHOD( SetTabActive )( THIS_ HWND, HWND, DWORD ) PURE;
ThumbBarAddButtons_ = 15    'STDMETHOD( ThumbBarAddButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
ThumbBarUpdateButtons_ = 16    'STDMETHOD( ThumbBarUpdateButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
ThumbBarSetImageList_ = 17  'STDMETHOD( ThumbBarSetImageList )( THIS_ HWND, HIMAGELIST ) PURE;
SetOverlayIcon_ = 18        'STDMETHOD( SetOverlayIcon )( THIS_ HWND, HICON, LPCWSTR ) PURE;
SetThumbnailTooltip_ = 19   'STDMETHOD( SetThumbnailTooltip )( THIS_ HWND, LPCWSTR ) PURE;
SetThumbnailClip_ = 20      'STDMETHOD( SetThumbnailClip )( THIS_ HWND, RECT * ) PURE;

    '                                '/* ITaskbarList4 methods */
    '    SetTabProperties_ = 21      'STDMETHOD( SetTabProperties )( THIS_ HWND, STPFLAG ) PURE;
End Enum

Private Enum THUMBBUTTONMASK
    THB_BITMAP = 1
    THB_ICON = 2
    THB_TOOLTIP = 4
    THB_FLAGS = 8

End Enum

Private Enum THUMBBUTTONFLAGS
    THBF_ENABLED = 0
    THBF_DISABLED = 1
    THBF_DISMISSONCLICK = 2
    THBF_NOBACKGROUND = 4
    THBF_HIDDEN = 8
    THBF_NONINTERACTIVE = 16

End Enum

Private Type THUMBBUTTON
    dwMask                                  As THUMBBUTTONMASK
    IID                                 As Long
    iBitmap                             As Long
    hIcon                               As Long
    szTip                               As String * 260
    dwFlags                             As THUMBBUTTONFLAGS

End Type

Private obj                             As Long
Attribute obj.VB_VarUserMemId = 1073938435
Private msgTaskbarCreated               As Long
Attribute msgTaskbarCreated.VB_VarUserMemId = 1073938436
Private cButtons_                       As Long
Attribute cButtons_.VB_VarUserMemId = 1073938437
Private icons_()                        As Long
Attribute icons_.VB_VarUserMemId = 1073938438
Private strTips_()                      As String
Attribute strTips_.VB_VarUserMemId = 1073938439
Private buttonsAddedFlag                As Boolean
Attribute buttonsAddedFlag.VB_VarUserMemId = 1073938440

Public Event ButtonPressed(ByVal Index As Integer)

Public Function ButtonsWindowProc(ByVal hWnd As Long, _
                                  ByVal uMsg As Long, _
                                  ByVal wParam As Long, _
                                  ByVal lParam As Long) As Long

    Select Case uMsg

        Case WM_COMMAND:

            'HiWord of wParam
            If wParam \ &H10000 = THBN_CLICKED Then
                RaiseEvent ButtonPressed(wParam And &HFFFF&)

            End If

        Case msgTaskbarCreated:

            If buttonsAddedFlag Then
                ThumbBarAddButtons wnd, cButtons_, icons_, strTips_

            End If

            '    Case WM_DWMSENDICONICTHUMBNAIL:
            '        Dim hbmp As Long, hdib As Long
            '        hdib = FreeImage_CreateFromDC(Form1.cover.hdc)
            '        Debug.Print FreeImage_GetBPP(hdib), FreeImage_GetWidth(hdib), FreeImage_GetHeight(hdib)
            '        hbmp = FreeImage_GetBitmap(hdib, , True)
            '        Taskbar_DwmSetIconicThumbnail(Form1.hwnd) = hbmp
            '        DeleteObject hbmp
            '        Debug.Print "WM_DWMSENDICONICTHUMBNAIL", lParam \ &H10000, lParam And &HFFFF&, hbmp, Form1.cover.Image.Handle
    End Select

    '    If Not (uMsg = 174 Or uMsg = 127 Or uMsg = 32) Then Debug.Print uMsg
    ButtonsWindowProc = CallWindowProc(oldproc, hWnd, uMsg, wParam, lParam)

End Function

Private Sub Class_Initialize()

Dim CLSID As Guid, InterfaceGuid        As Guid

    Call CLSIDFromString(StrConv(CLSID_TaskbarList, vbUnicode), CLSID)
    Call IIDFromString(StrConv(IID_ITaskbarList3, vbUnicode), InterfaceGuid)
    Call CoCreateInstance(CLSID, 0, 1, InterfaceGuid, obj)
    msgTaskbarCreated = RegisterWindowMessage("TaskbarButtonCreated")

End Sub

Private Sub Class_Terminate()

    On Error Resume Next

    If wnd Then StopSubclass wnd, oldproc
    If obj Then Call CallInterface(obj, unk_Release, 0)

End Sub

Public Function SetProgressState(ByVal hWnd As Long, ByVal tbpFlags As Long) As Long
    SetProgressState = CallInterface(obj, SetProgressState_, 2, hWnd, tbpFlags)

End Function

Public Function SetProgressValue(ByVal hWnd As Long, _
                                 ByVal ullCompleted As Long, _
                                 ByVal ullTotal As Long) As Long
    SetProgressValue = CallInterface(obj, SetProgressValue_, 5, hWnd, ullCompleted, 0, ullTotal, 0)

End Function

Public Function SetOverlayIcon(ByVal hWnd As Long, _
                               ByVal hIcon As Long, _
                               Optional ByRef pszDescription As String) As Long
    SetOverlayIcon = CallInterface(obj, SetOverlayIcon_, 3, hWnd, hIcon, StrPtr(pszDescription))

End Function

Public Function SetThumbnailClip(ByVal hWnd As Long, _
                                 ByVal Left As Long, _
                                 ByVal Top As Long, _
                                 ByVal Right As Long, _
                                 ByVal Bottom As Long) As Long

Dim RC                                  As RECT

    RC.Left = Left
    RC.Top = Top
    RC.Right = Right
    RC.Bottom = Bottom
    SetThumbnailClip = CallInterface(obj, SetThumbnailClip_, 2, hWnd, VarPtr(RC))

End Function

Public Function SetThumbnailTooltip(ByVal hWnd As Long, ByRef pszTip As String) As Long
    SetThumbnailTooltip = CallInterface(obj, SetThumbnailTooltip_, 2, hWnd, StrPtr(pszTip))

End Function

Public Function ThumbBarAddButtons(ByVal hWnd As Long, _
                                   ByVal cButtons As Long, _
                                   ByRef Icons() As Long, _
                                   ByRef strTips() As String) As Long

Dim i As Integer, icn(6) As Long, iTips(6) As String

    If cButtons < 1 Or cButtons > 7 Then
        Exit Function

    End If

    ReDim Buttons(cButtons - 1) As THUMBBUTTON

    For i = 0 To UUBound(Icons)
        icn(i) = Icons(i)
    Next

    For i = 0 To UUBound(strTips)
        iTips(i) = strTips(i)
    Next

    For i = 0 To cButtons - 1
        Buttons(i).dwMask = THB_ICON Or THB_TOOLTIP
        'Or THB_FLAGS
        Buttons(i).IID = i
        Buttons(i).hIcon = icn(i)
        Buttons(i).szTip = iTips(i) & vbNullChar
        '        buttons(i).dwFlags = THBF_NOBACKGROUND
    Next
    ThumbBarAddButtons = CallInterface(obj, ThumbBarAddButtons_, 3, hWnd, cButtons, VarPtr(Buttons(0)))

    'If it was too early to add buttons, this function will be recalled after TaskbarButtonCreated message
    If Not buttonsAddedFlag Then
        cButtons_ = cButtons
        icons_ = Icons
        strTips_ = strTips
        buttonsAddedFlag = True

        If wnd = 0 Then
            wnd = hWnd
            StartSubclass ASMArr, wnd, oldproc

        End If

    End If

End Function

Private Sub StartSubclass(ByRef ASM() As Byte, _
                          ByVal hWnd As Long, _
                          ByRef OldWndProc As Long, _
                          Optional ByVal ProcNumber As Long)

' Сабклассинг с пом. ASM (автора не знаю...)
Dim Lng As Long, tPtr                   As Long

    Lng = Len(AsmMain) \ 2&
    ReDim ASM(0 To Lng - 1)

    For Lng = 0 To Lng - 1
        'ASM(lng) = Val("&H" & Mid$(AsmMain, (lng) * 2& + 1, 2&))
        ASM(Lng) = CLng("&H" & Mid$(AsmMain, (Lng) * 2& + 1, 2&))
    Next
    Call CopyMemory(tPtr, ByVal ObjPtr(Me), 4&)
    Call CopyMemory(Lng, ByVal tPtr + &H1C + (4& * ProcNumber), 4&)
    Call CopyMemory(ASM(23), ObjPtr(Me), 4&)
    Call CopyMemory(ASM(28), Lng, 4&)
    OldWndProc = SetWindowLong(hWnd, &HFFFC, VarPtr(ASM(0)))

End Sub

Private Sub StopSubclass(ByVal hWnd As Long, ByVal OldWndProc As Long)
    Call SetWindowLong(hWnd, &HFFFC, OldWndProc)

End Sub

Private Function UUBound(Arr) As Long

    On Error GoTo 1:

    UUBound = UBound(Arr)
    Exit Function
1
    UUBound = -1

End Function

Public Property Get isAccessible() As Boolean
    isAccessible = obj

End Property

Public Property Let Taskbar_DwmSetWindowAttribute(ByVal hWnd As Long, ByVal property As Long, ByVal Value As Long)
    DwmSetWindowAttribute hWnd, property, Value, 4

End Property

'
'Public Property Let Taskbar_DwmSetIconicThumbnail(ByVal hwnd As Long, ByVal hbmp As Long)
'    Debug.Print Hex$(DwmSetIconicThumbnail(hwnd, hbmp, 0&)), "Taskbar_DwmSetIconicThumbnail"
'End Property
