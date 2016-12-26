Attribute VB_Name = "mApiWindow"
Option Explicit

Public Const ALL_MESSAGES       As Long = -1    'All messages added or deleted

' --Color Constant
Public Const COLOR_BTNFACE      As Long = 15
Public Const COLOR_BTNHIGHLIGHT As Long = 20
Public Const COLOR_BTNSHADOW    As Long = 16
Public Const COLOR_HIGHLIGHT    As Long = 13
Public Const COLOR_GRAYTEXT     As Long = 17
Public Const CLR_INVALID        As Long = &HFFFF
Public Const DIB_RGB_COLORS     As Long = 0

' --Windows Messages
Public Const WM_USER            As Long = &H400
Public Const WM_EXITSIZEMOVE    As Long = &H232
Public Const WM_LBUTTONDOWN     As Long = &H201
Public Const WM_LBUTTONUP       As Long = &H202
Public Const WM_MOUSELEAVE      As Long = &H2A3
Public Const WM_MOUSEMOVE       As Long = &H200
Public Const WM_MOVING          As Long = &H216
Public Const WM_RBUTTONDBLCLK   As Long = &H206
Public Const WM_RBUTTONDOWN     As Long = &H204
Public Const WM_SETFOCUS        As Long = &H7
Public Const WM_SHOWWINDOW      As Long = &H18
Public Const WM_SIZING          As Long = &H214
Public Const WM_SYSCOLORCHANGE  As Long = &H15
Public Const WM_THEMECHANGED    As Long = &H31A
Public Const WM_PAINT           As Long = &HF
Public Const WM_NCPAINT         As Long = &H85
Public Const WM_CLOSE           As Long = &H10
Public Const WM_DESTROY         As Long = &H2
Public Const WM_NCDESTROY       As Long = &H82
Public Const WM_QUIT            As Long = &H12
Public Const WM_COMMAND         As Long = &H111
Public Const WM_NOTIFY          As Long = &H4E
Public Const WM_NCACTIVATE      As Long = &H86
Public Const WM_ACTIVATE        As Long = &H6
Public Const WM_SETTEXT         As Long = &HC
Public Const WM_GETTEXT         As Long = &HD
Public Const WM_GETTEXTLENGTH   As Long = &HE
Public Const WM_KILLFOCUS       As Long = &H8
Public Const EM_SETREADONLY     As Long = &HCF
Public Const EM_NOSETFOCUS      As Long = (&H1500 + 7)
Public Const WS_CAPTION         As Long = &HC00000
Public Const WS_THICKFRAME      As Long = &H40000
Public Const WS_MINIMIZEBOX     As Long = &H20000
Public Const WS_BORDER          As Long = &H800000
Public Const WS_VISIBLE         As Long = &H10000000
Public Const WS_CHILD           As Long = &H40000000
Public Const WS_EX_TOPMOST      As Long = &H8&
Public Const WS_EX_TOOLWINDOW   As Long = &H80
Public Const WS_EX_LAYOUTRTL    As Long = &H400000
Public Const WS_EX_STATICEDGE   As Long = &H20000
Public Const WS_EX_CLIENTEDGE   As Long = &H200&
Public Const SWP_REFRESH        As Long = (&H1 Or &H2 Or &H4 Or &H20)
Public Const SWP_NOACTIVATE     As Long = &H10
Public Const SWP_NOMOVE         As Long = &H2
Public Const SWP_NOSIZE         As Long = &H1
Public Const SWP_SHOWWINDOW     As Long = &H40
Public Const SWP_NOOWNERZORDER  As Long = &H200
Public Const SWP_NOZORDER       As Long = &H4
Public Const SWP_FRAMECHANGED   As Long = &H20
Public Const HWND_TOPMOST       As Long = -&H1
Public Const CW_USEDEFAULT      As Long = &H80000000
Public Const GWL_EXSTYLE        As Long = (-20)
Public Const GWL_STYLE          As Long = -16
Public Const GWL_WNDPROC        As Long = -4               'Get/SetWindow offset to the WndProc procedure address
Public Const SW_HIDE            As Long = 0
Public Const SW_SHOWNORMAL      As Long = 1
Public Const GW_HWNDPREV        As Long = 3
Public Const GW_OWNER = 4
Public Const WM_SETFONT         As Long = &H30
Public Const WM_GETFONT         As Long = &H31

'Tooltip Window Constants
Public Const TTS_NOPREFIX        As Long = &H2
Public Const TTF_TRANSPARENT     As Long = &H100
Public Const TTF_IDISHWND        As Long = &H1
Public Const TTF_CENTERTIP       As Long = &H2
Public Const TTM_ADDTOOLA        As Long = (WM_USER + 4)
Public Const TTM_ADDTOOLW        As Long = (WM_USER + 50)
Public Const TTM_ACTIVATE        As Long = WM_USER + 1
Public Const TTM_UPDATETIPTEXTA  As Long = (WM_USER + 12)
Public Const TTM_SETMAXTIPWIDTH  As Long = (WM_USER + 24)
Public Const TTM_SETTIPBKCOLOR   As Long = (WM_USER + 19)
Public Const TTM_SETTIPTEXTCOLOR As Long = (WM_USER + 20)
Public Const TTM_SETTITLE        As Long = (WM_USER + 32)
Public Const TTM_SETTITLEW       As Long = (WM_USER + 33)
Public Const TTS_BALLOON         As Long = &H40
Public Const TTS_ALWAYSTIP       As Long = &H1
Public Const TTF_SUBCLASS        As Long = &H10
Public Const TOOLTIPS_CLASSA     As String = "tooltips_class32"

' </-- Установка минимальных размеров окна
Public Const WM_GETMINMAXINFO    As Long = &H24

Public Type MINMAXINFO
    ptReserved                   As POINTAPI
    ptMaxSize                    As POINTAPI
    ptMaxPosition                As POINTAPI
    ptMinTrackSize               As POINTAPI
    ptMaxTrackSize               As POINTAPI
End Type

Public Type Resize
    xMin                         As Single
    yMin                         As Single
    xMax                         As Single
    yMax                         As Single
End Type

Public Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, ByVal cbCopy As Long)
Public Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function EnableWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function RedrawWindow Lib "user32.dll" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageW Lib "user32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowLongA Lib "user32.dll" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function PostMessageLong Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hWndLock As Long) As Long
Public Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function GetFocus Lib "user32.dll" () As Long
Public Declare Function UpdateWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function EnumThreadWindows Lib "user32.dll" (ByVal dwThreadID As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

'Public Declare Function GetActiveWindow Lib "user32" () As Long
'Public Declare Function GetForegroundWindow Lib "user32" () As Long
'Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'Public Declare Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long

Public Function IsChildOfControl(ByVal hWndControl As Long, ByVal hWndParentControl As Long) As Boolean

    Dim hParent As Long

    hParent = GetParent(hWndControl)

    Do While hParent <> 0

        If hParent = hWndParentControl Then
            IsChildOfControl = True
            Exit Do
        End If

        hParent = GetParent(hParent)
    Loop

End Function

