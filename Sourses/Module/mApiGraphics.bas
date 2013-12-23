Attribute VB_Name = "mApiGraphics"
Option Explicit

' --Formatting Text Consts
Public Const DT_LEFT       As Long = &H0
Public Const DT_CENTER     As Long = &H1
Public Const DT_RIGHT      As Long = &H2
Public Const DT_NOCLIP     As Long = &H100
Public Const DT_WORDBREAK  As Long = &H10
Public Const DT_CALCRECT   As Long = &H400
Public Const DT_RTLREADING As Long = &H20000
Public Const DT_DRAWFLAG   As Long = DT_CENTER Or DT_WORDBREAK
Public Const DT_TOP        As Long = &H0
Public Const DT_BOTTOM     As Long = &H8
Public Const DT_VCENTER    As Long = &H4
Public Const DT_SINGLELINE As Long = &H20
Public Const DT_WORD_ELLIPSIS = &H40000
Public Const TransColor      As Long = &H8000000F

'   DrawEdge Message Constants
Public Const BDR_RAISEDOUTER As Long = &H1
Public Const BDR_SUNKENOUTER As Long = &H2
Public Const BDR_RAISEDINNER As Long = &H4
Public Const BDR_SUNKENINNER As Long = &H8
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const BF_LEFT      As Long = &H1
Public Const BF_TOP       As Long = &H2
Public Const BF_RIGHT     As Long = &H4
Public Const BF_BOTTOM    As Long = &H8
Public Const BF_RECT      As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_SUNKEN95 As Long = &HA
Public Const BDR_RAISED95 As Long = &H5

' §§§§§§§§§§§§§§§§§§§§§§§§§§ ImageList §§§§§§§§§§§§§§§§§§§§§§§§§§
Public Const SM_CXICON    As Long = 11
Public Const SM_CYICON    As Long = 12
Public Const SM_CYSMICON  As Long = 50
Public Const SM_CXSMICON  As Long = 49

' --System Hand Pointer
Public Const IDC_HAND     As Long = 32649

' --drawing Icon Constants
Public Const DI_NORMAL    As Long = &H3

Public Type Size
    CX                                  As Long
    CY                                  As Long
End Type

Public Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Public Type POINT
    X                                   As Long
    Y                                   As Long
End Type

Public Type RGB
    Red                                 As Byte
    Green                               As Byte
    Blue                                As Byte
End Type

Public Type RGBQUAD
    Blue                                As Byte
    Green                               As Byte
    Red                                 As Byte
    Alpha                               As Byte
End Type

Public Type ICONINFO
    fIcon                               As Long
    XHotspot                            As Long
    YHotspot                            As Long
    hBMMask                             As Long
    hBMColor                            As Long
End Type

'  for gradient painting and bitmap tiling
Public Type BITMAPINFOHEADER
    biSize                              As Long
    biWidth                             As Long
    biHeight                            As Long
    biPlanes                            As Integer
    biBitCount                          As Integer
    biCompression                       As Long
    biSizeImage                         As Long
    biXPelsPerMeter                     As Long
    biYPelsPerMeter                     As Long
    biClrUsed                           As Long
    biClrImportant                      As Long
End Type

'flicker free drawing
Public Type BITMAP
    BMType                              As Long
    BMWidth                             As Long
    BMHeight                            As Long
    BMWidthBytes                        As Long
    BMPlanes                            As Integer
    BMBitsPixel                         As Integer
    BMBits                              As Long
End Type

Public Type BITMAPINFO
    bmiHeader                           As BITMAPINFOHEADER
    bmiColors                           As RGB
End Type

''ooltip Window Types
Public Type TOOLINFO
    lSize                               As Long
    lFlags                              As Long
    lhWnd                               As Long
    lID                                 As Long
    lpRect                              As RECT
    hInstance                           As Long
    lpStr                               As String
    lParam                              As Long
End Type

'Tooltip Window Types [for UNICODE support]
Public Type TOOLINFOW
    lSize                               As Long
    lFlags                              As Long
    lhWnd                               As Long
    lID                                 As Long
    lpRect                              As RECT
    hInstance                           As Long
    lpStrW                              As Long
    lParam                              As Long
End Type

Public Type BITMAPINFO8
    bmiHeader                           As BITMAPINFOHEADER
    bmiColors(255)                      As RGBQUAD
End Type

Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function ReleaseCapture Lib "user32.dll" () As Long
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT) As Long
Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetCapture Lib "user32.dll" () As Long
Public Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FrameRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function OffsetRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CopyRect Lib "user32.dll" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINT) As Long
Public Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, lpPoint As POINT) As Long
Public Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal XLeft As Long, ByVal YTop As Long, ByVal hIcon As Long, ByVal CXWidth As Long, ByVal CYWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDcDest As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, qRC As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function OleTranslateColor Lib "OlePro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Public Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Public Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Public Declare Function GetThemeBackgroundRegion Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pRegion As Long) As Long
Public Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetLayout Lib "gdi32.dll" (ByVal hDC As Long, ByVal dwLayout As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function CreateDIBSection8 Lib "gdi32.dll" Alias "CreateDIBSection" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO8, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Public Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function InflateRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function OleTranslateColorByRef Lib "oleaut32.dll" Alias "OleTranslateColor" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Public Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetPixelV Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function RoundRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32.dll" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function GetNearestColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function DwmIsCompositionEnabled Lib "dwmapi" (ByRef pfEnabled As Long) As Long
Public Declare Function GradientFill Lib "msimg32.dll" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Public Type TRIVERTEX
    X                                   As Long
    Y                                   As Long
    Red                                 As Integer    'ushort value
    Green                               As Integer    'ushort value
    Blue                                As Integer    'ushort value
    Alpha                               As Integer    'ushort value
End Type

Public Type GRADIENT_RECT
    UpperLeft                           As Long
    LowerRight                          As Long
End Type

Public Const GRADIENT_FILL_RECT_V As Long = &H1

'!--------------------------------------------------------------------------------
'! Procedure   (Ôóíêöèÿ)   :   Sub FlatBorderButton
'! Description (Îïèñàíèå)  :   [Äåëàåò êíîïêó íàæàòîé/íîðìàëüíîé]
'! Parameters  (Ïåðåìåííûå):   lngHWnd (Long)
'                              mbFlat (Boolean = True)
'!--------------------------------------------------------------------------------
Public Sub FlatBorderButton(ByVal lngHWnd As Long, Optional mbFlat As Boolean = True)

    Dim TFlat As Long

    TFlat = GetWindowLong(lngHWnd, GWL_EXSTYLE)

    If mbFlat Then
        TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    Else
        TFlat = TFlat And WS_EX_CLIENTEDGE
    End If

    SetWindowLong lngHWnd, GWL_EXSTYLE, TFlat
    SetWindowPos lngHWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub
