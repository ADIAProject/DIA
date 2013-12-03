Attribute VB_Name = "mApiWindow"
Option Explicit

Public Const ALL_MESSAGES               As Long = -1    'All messages added or deleted

' --Color Constant
Public Const COLOR_BTNFACE              As Long = 15
Public Const COLOR_BTNHIGHLIGHT         As Long = 20
Public Const COLOR_BTNSHADOW            As Long = 16
Public Const COLOR_HIGHLIGHT            As Long = 13
Public Const COLOR_GRAYTEXT             As Long = 17
Public Const CLR_INVALID                As Long = &HFFFF
Public Const DIB_RGB_COLORS             As Long = 0

' --Windows Messages
Public Const WM_USER                    As Long = &H400
Public Const WM_EXITSIZEMOVE            As Long = &H232
Public Const WM_LBUTTONDOWN             As Long = &H201
Public Const WM_LBUTTONUP               As Long = &H202
Public Const WM_MOUSELEAVE              As Long = &H2A3
Public Const WM_MOUSEMOVE               As Long = &H200
Public Const WM_MOVING                  As Long = &H216
Public Const WM_RBUTTONDBLCLK           As Long = &H206
Public Const WM_RBUTTONDOWN             As Long = &H204
Public Const WM_SETFOCUS                As Long = &H7
Public Const WM_SHOWWINDOW              As Long = &H18
Public Const WM_SIZING                  As Long = &H214
Public Const WM_SYSCOLORCHANGE          As Long = &H15
Public Const WM_THEMECHANGED            As Long = &H31A
Public Const WM_PAINT                   As Long = &HF
Public Const WM_NCPAINT                 As Long = &H85
Public Const WM_CLOSE                   As Long = &H10
Public Const WM_COMMAND                 As Long = &H111
Public Const WM_NOTIFY                  As Long = &H4E
Public Const WM_NCACTIVATE              As Long = &H86
Public Const WM_ACTIVATE                As Long = &H6
Public Const WM_SETTEXT                 As Long = &HC
Public Const WM_KILLFOCUS               As Long = &H8
Public Const EM_SETREADONLY             As Long = &HCF
Public Const EM_NOSETFOCUS              As Long = (&H1500 + 7)
Public Const WS_CAPTION                 As Long = &HC00000
Public Const WS_THICKFRAME              As Long = &H40000
Public Const WS_MINIMIZEBOX             As Long = &H20000
Public Const WS_BORDER                  As Long = &H800000
Public Const WS_VISIBLE                 As Long = &H10000000
Public Const WS_CHILD                   As Long = &H40000000
Public Const WS_EX_TOPMOST              As Long = &H8&
Public Const WS_EX_TOOLWINDOW           As Long = &H80
Public Const WS_EX_LAYOUTRTL            As Long = &H400000
Public Const SWP_REFRESH                As Long = (&H1 Or &H2 Or &H4 Or &H20)
Public Const SWP_NOACTIVATE             As Long = &H10
Public Const SWP_NOMOVE                 As Long = &H2
Public Const SWP_NOSIZE                 As Long = &H1
Public Const SWP_SHOWWINDOW             As Long = &H40
Public Const SWP_NOOWNERZORDER          As Long = &H200
Public Const SWP_NOZORDER               As Long = &H4
Public Const SWP_FRAMECHANGED           As Long = &H20
Public Const HWND_TOPMOST               As Long = -&H1
Public Const CW_USEDEFAULT              As Long = &H80000000
Public Const GWL_EXSTYLE                As Long = (-20)
Public Const GWL_STYLE                  As Long = -16
Public Const GWL_WNDPROC                As Long = -4               'Get/SetWindow offset to the WndProc procedure address
Public Const SW_HIDE                    As Long = 0
Public Const GW_HWNDPREV                As Long = 3
Public Const GW_OWNER = 4

'Tooltip Window Constants
'Public Const TTS_NOPREFIX        As Long = &H2
Public Const TTF_TRANSPARENT            As Long = &H100

'Public Const TTF_IDISHWND        As Long = &H1
'Public Const TTF_CENTERTIP       As Long = &H2
Public Const TTM_ADDTOOLA               As Long = (WM_USER + 4)
Public Const TTM_ADDTOOLW               As Long = (WM_USER + 50)
Public Const TTM_ACTIVATE               As Long = WM_USER + 1
Public Const TTM_UPDATETIPTEXTA         As Long = (WM_USER + 12)
Public Const TTM_SETMAXTIPWIDTH         As Long = (WM_USER + 24)
Public Const TTM_SETTIPBKCOLOR          As Long = (WM_USER + 19)
Public Const TTM_SETTIPTEXTCOLOR        As Long = (WM_USER + 20)
Public Const TTM_SETTITLE               As Long = (WM_USER + 32)
Public Const TTM_SETTITLEW              As Long = (WM_USER + 33)

'Public Const TTS_BALLOON         As Long = &H40
'Public Const TTS_ALWAYSTIP       As Long = &H1
'Public Const TTF_SUBCLASS        As Long = &H10
Public Const TOOLTIPS_CLASSA            As String = "tooltips_class32"


' </-- ”становка минимальных размеров окна
Public Const WM_GETMINMAXINFO           As Long = &H24

Public Type MINMAXINFO
    ptReserved                              As POINT
    ptMaxSize                           As POINT
    ptMaxPosition                       As POINT
    ptMinTrackSize                      As POINT
    ptMaxTrackSize                      As POINT
End Type

Public Type Resize
    xMin                                    As Single
    yMin                                As Single
    xMax                                As Single
    yMax                                As Single
End Type

Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Sub CopyMemoryToMinMaxInfo _
                    Lib "kernel32.dll" _
                        Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, _
                                               ByVal hpvSource As Long, _
                                               ByVal cbCopy As Long)

Public Declare Sub CopyMemoryFromMinMaxInfo _
                    Lib "kernel32.dll" _
                        Alias "RtlMoveMemory" (ByVal hpvDest As Long, _
                                               hpvSource As MINMAXINFO, _
                                               ByVal cbCopy As Long)

' ƒезјктиваци€ окна
Public Declare Function EnableWindow _
                         Lib "user32.dll" (ByVal hWnd As Long, _
                                           ByVal fEnable As Long) As Long

Public Declare Function GetWindow _
                         Lib "user32.dll" (ByVal hWnd As Long, _
                                           ByVal wCmd As Long) As Long

Public Declare Function SendMessage _
                         Lib "user32.dll" _
                             Alias "SendMessageA" (ByVal hWnd As Long, _
                                                   ByVal wMsg As Long, _
                                                   ByVal wParam As Long, _
                                                   lParam As Any) As Long

Public Declare Function SendMessageW _
                         Lib "user32.dll" (ByVal hWnd As Long, _
                                           ByVal wMsg As Long, _
                                           ByVal wParam As Long, _
                                           lParam As Any) As Long

Public Declare Function SendMessageLong _
                         Lib "user32.dll" _
                             Alias "SendMessageA" (ByVal hWnd As Long, _
                                                   ByVal wMsg As Long, _
                                                   ByVal wParam As Long, _
                                                   ByVal lParam As Long) As Long

Public Declare Function SendMessageLongW _
                         Lib "user32.dll" _
                             Alias "SendMessageW" (ByVal hWnd As Long, _
                                                   ByVal wMsg As Long, _
                                                   ByVal wParam As Long, _
                                                   ByVal lParam As Long) As Long

Public Declare Function GetWindowLong _
                         Lib "user32.dll" _
                             Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                                     ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong _
                         Lib "user32.dll" _
                             Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                     ByVal nIndex As Long, _
                                                     ByVal dwNewLong As Long) As Long

Public Declare Function SetWindowLongA _
                         Lib "user32.dll" (ByVal hWnd As Long, _
                                           ByVal nIndex As Long, _
                                           ByVal dwNewLong As Long) As Long

Public Declare Function PostMessageLong _
                         Lib "user32.dll" _
                             Alias "PostMessageA" (ByVal hWnd As Long, _
                                                   ByVal Msg As Long, _
                                                   ByVal wParam As Long, _
                                                   ByVal lParam As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long

Public Declare Function SetFocusApi _
                         Lib "user32.dll" _
                             Alias "SetFocus" (ByVal hWnd As Long) As Long

Public Declare Function GetFocus Lib "user32.dll" () As Long

Public Declare Function UpdateWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function FindWindowEx _
                         Lib "user32.dll" _
                             Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                                    ByVal hWnd2 As Long, _
                                                    ByVal lpsz1 As String, _
                                                    ByVal lpsz2 As String) As Long

Public Declare Function EnumThreadWindows _
                         Lib "user32.dll" (ByVal dwThreadId As Long, _
                                           ByVal lpfn As Long, _
                                           ByVal lParam As Long) As Long

Public Declare Function GetWindowThreadProcessId _
                         Lib "user32.dll" (ByVal hWnd As Long, _
                                           lpdwProcessId As Long) As Long

Public Declare Function SetWindowText _
                         Lib "user32.dll" _
                             Alias "SetWindowTextA" (ByVal hWnd As Long, _
                                                     ByVal lpString As String) As Long

Public Declare Function GetWindowText _
                         Lib "user32.dll" _
                             Alias "GetWindowTextA" (ByVal hWnd As Long, _
                                                     ByVal lpString As String, _
                                                     ByVal cch As Long) As Long

Public Declare Function EnumChildWindows _
                         Lib "user32.dll" (ByVal hWndParent As Long, _
                                           ByVal lpEnumFunc As Long, _
                                           ByVal lParam As Long) As Long

Public Declare Function GetClassLong _
                         Lib "user32.dll" _
                             Alias "GetClassLongA" (ByVal hWnd As Long, _
                                                    ByVal nIndex As Long) As Long

Public Declare Function SetClassLong _
                         Lib "user32.dll" _
                             Alias "SetClassLongA" (ByVal hWnd As Long, _
                                                    ByVal nIndex As Long, _
                                                    ByVal dwNewLong As Long) As Long

Public Declare Function ReleaseDC _
                         Lib "user32.dll" (ByVal hWnd As Long, _
                                           ByVal hDC As Long) As Long


