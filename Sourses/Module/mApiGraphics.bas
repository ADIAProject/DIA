Attribute VB_Name = "mApiGraphics"
Option Explicit

' --Formatting Text Consts
Public Const DT_LEFT                    As Long = &H0
Public Const DT_CENTER                  As Long = &H1
Public Const DT_RIGHT                   As Long = &H2
Public Const DT_NOCLIP                  As Long = &H100
Public Const DT_WORDBREAK               As Long = &H10
Public Const DT_CALCRECT                As Long = &H400
Public Const DT_RTLREADING              As Long = &H20000    ' Right to left
Public Const DT_DRAWFLAG                As Long = DT_CENTER Or DT_WORDBREAK
Public Const DT_TOP                     As Long = &H0
Public Const DT_BOTTOM                  As Long = &H8
Public Const DT_VCENTER                 As Long = &H4
Public Const DT_SINGLELINE              As Long = &H20
Public Const DT_WORD_ELLIPSIS = &H40000
Public Const TransColor                 As Long = &H8000000F

'   DrawEdge Message Constants
Public Const BDR_RAISEDOUTER            As Long = &H1
Public Const BDR_SUNKENOUTER            As Long = &H2
Public Const BDR_RAISEDINNER            As Long = &H4
Public Const BDR_SUNKENINNER            As Long = &H8
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const BF_LEFT                    As Long = &H1
Public Const BF_TOP                     As Long = &H2
Public Const BF_RIGHT                   As Long = &H4
Public Const BF_BOTTOM                  As Long = &H8
Public Const BF_RECT                    As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_SUNKEN95               As Long = &HA
Public Const BDR_RAISED95               As Long = &H5

' ßßßßßßßßßßßßßßßßßßßßßßßßßß ImageList ßßßßßßßßßßßßßßßßßßßßßßßßßß
Public Const SM_CXICON                  As Long = 11
Public Const SM_CYICON                  As Long = 12
Public Const SM_CYSMICON                As Long = 50
Public Const SM_CXSMICON                As Long = 49

' --System Hand Pointer
Public Const IDC_HAND                   As Long = 32649

' --drawing Icon Constants
Public Const DI_NORMAL                  As Long = &H3

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

Public Type RGBTRIPLE
    rgbBlue                             As Byte
    rgbGreen                            As Byte
    rgbRed                              As Byte
End Type

'  RGB Colors structure
Public Type RGBColor
    R                                   As Single
    G                                   As Single
    B                                   As Single
End Type

Public Type RGBQUAD
    rgbBlue                             As Byte
    rgbGreen                            As Byte
    rgbRed                              As Byte
    rgbAlpha                            As Byte
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
    bmiColors                           As RGBTRIPLE
End Type

''Tooltip Window Types
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
Public Declare Function WindowFromPoint _
                         Lib "user32.dll" (ByVal xPoint As Long, _
                                           ByVal yPoint As Long) As Long

Public Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetCapture Lib "user32.dll" () As Long
Public Declare Function Rectangle _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal X1 As Long, _
                                          ByVal Y1 As Long, _
                                          ByVal X2 As Long, _
                                          ByVal Y2 As Long) As Long

Public Declare Function DrawTextW _
                         Lib "user32.dll" (ByVal hDC As Long, _
                                           ByVal lpStr As Long, _
                                           ByVal nCount As Long, _
                                           lpRect As RECT, _
                                           ByVal wFormat As Long) As Long

Public Declare Function DrawFocusRect _
                         Lib "user32.dll" (ByVal hDC As Long, _
                                           lpRect As RECT) As Long

Public Declare Function GetTextExtentPoint32 _
                         Lib "gdi32.dll" _
                             Alias "GetTextExtentPoint32W" (ByVal hDC As Long, _
                                                            ByVal lpsz As Long, _
                                                            ByVal cbString As Long, _
                                                            lpSize As Size) As Long

Public Declare Function FillRect _
                         Lib "user32.dll" (ByVal hDC As Long, _
                                           lpRect As RECT, _
                                           ByVal hBrush As Long) As Long

Public Declare Function FrameRect _
                         Lib "user32.dll" (ByVal hDC As Long, _
                                           lpRect As RECT, _
                                           ByVal hBrush As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function SetTextColor _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal crColor As Long) As Long

Public Declare Function GetTextColor Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function SelectObject _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal hObject As Long) As Long

Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function GetPixel _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal X As Long, _
                                          ByVal Y As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function GetObject _
                         Lib "gdi32.dll" _
                             Alias "GetObjectA" (ByVal hObject As Long, _
                                                 ByVal nCount As Long, _
                                                 lpObject As Any) As Long

Public Declare Function OffsetRect _
                         Lib "user32.dll" (lpRect As RECT, _
                                           ByVal X As Long, _
                                           ByVal Y As Long) As Long

Public Declare Function CopyRect _
                         Lib "user32.dll" (lpDestRect As RECT, _
                                           lpSourceRect As RECT) As Long

Public Declare Function GetClientRect _
                         Lib "user32.dll" (ByVal hWnd As Long, _
                                           lpRect As RECT) As Long

Public Declare Function SetRect _
                         Lib "user32.dll" (lpRect As RECT, _
                                           ByVal X1 As Long, _
                                           ByVal Y1 As Long, _
                                           ByVal X2 As Long, _
                                           ByVal Y2 As Long) As Long

Public Declare Function SetWindowRgn _
                         Lib "user32.dll" (ByVal hWnd As Long, _
                                           ByVal hRgn As Long, _
                                           ByVal bRedraw As Boolean) As Long

Public Declare Function LoadCursor _
                         Lib "user32.dll" _
                             Alias "LoadCursorA" (ByVal hInstance As Long, _
                                                  ByVal lpCursorName As Long) As Long

Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Public Declare Function MoveToEx _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal X As Long, _
                                          ByVal Y As Long, _
                                          lpPoint As POINT) As Long

Public Declare Function LineTo _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal X As Long, _
                                          ByVal Y As Long) As Long

Public Declare Function CreatePen _
                         Lib "gdi32.dll" (ByVal nPenStyle As Long, _
                                          ByVal nWidth As Long, _
                                          ByVal crColor As Long) As Long


Public Declare Function GetWindowRect _
                         Lib "user32.dll" (ByVal hWnd As Long, _
                                           lpRect As RECT) As Long
                                                      
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long

Public Declare Function SetParent _
                         Lib "user32.dll" (ByVal hWndChild As Long, _
                                           ByVal hWndNewParent As Long) As Long


Public Declare Function ScreenToClient _
                         Lib "user32.dll" (ByVal hWnd As Long, _
                                           lpPoint As POINT) As Long

Public Declare Function DrawText _
                         Lib "user32.dll" _
                             Alias "DrawTextA" (ByVal hDC As Long, _
                                                ByVal lpStr As String, _
                                                ByVal nCount As Long, _
                                                lpRect As RECT, _
                                                ByVal wFormat As Long) As Long

Public Declare Function DrawIconEx _
                         Lib "user32.dll" (ByVal hDC As Long, _
                                           ByVal XLeft As Long, _
                                           ByVal YTop As Long, _
                                           ByVal hIcon As Long, _
                                           ByVal CXWidth As Long, _
                                           ByVal CYWidth As Long, _
                                           ByVal istepIfAniCur As Long, _
                                           ByVal hbrFlickerFreeDraw As Long, _
                                           ByVal diFlags As Long) As Long

Public Declare Function BitBlt _
                         Lib "gdi32.dll" (ByVal hDcDest As Long, _
                                          ByVal X As Long, _
                                          ByVal Y As Long, _
                                          ByVal nWidth As Long, _
                                          ByVal nHeight As Long, _
                                          ByVal hDCSrc As Long, _
                                          ByVal XSrc As Long, _
                                          ByVal YSrc As Long, _
                                          ByVal dwRop As Long) As Long

Public Declare Function FillRgn _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal hRgn As Long, _
                                          ByVal hBrush As Long) As Long

Public Declare Function SetPixel _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal X As Long, _
                                          ByVal Y As Long, _
                                          ByVal crColor As Long) As Long

Public Declare Function CreateCompatibleBitmap _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal nWidth As Long, _
                                          ByVal nHeight As Long) As Long

Public Declare Function CreateBitmap _
                         Lib "gdi32.dll" (ByVal nWidth As Long, _
                                          ByVal nHeight As Long, _
                                          ByVal nPlanes As Long, _
                                          ByVal nBitCount As Long, _
                                          lpBits As Any) As Long

Public Declare Function DrawEdge _
                         Lib "user32.dll" (ByVal hDC As Long, _
                                           qRC As RECT, _
                                           ByVal Edge As Long, _
                                           ByVal grfFlags As Long) As Long

Public Declare Function OleTranslateColor _
                         Lib "OlePro32.dll" (ByVal OLE_COLOR As Long, _
                                             ByVal HPALETTE As Long, _
                                             pccolorref As Long) As Long

Public Declare Function GetDeviceCaps _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal nIndex As Long) As Long

Public Declare Function OpenThemeData _
                         Lib "uxtheme.dll" (ByVal hWnd As Long, _
                                            ByVal pszClassList As Long) As Long
Public Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

Public Declare Function DrawThemeBackground _
                         Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                            ByVal lhDC As Long, _
                                            ByVal iPartId As Long, _
                                            ByVal iStateId As Long, _
                                            pRect As RECT, _
                                            pClipRect As RECT) As Long

Public Declare Function GetThemeBackgroundRegion _
                         Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                            ByVal hDC As Long, _
                                            ByVal iPartId As Long, _
                                            ByVal iStateId As Long, _
                                            pRect As RECT, _
                                            pRegion As Long) As Long

Public Declare Function GetCurrentThemeName Lib "uxtheme.dll" ( _
                                            ByVal pszThemeFileName As Long, _
                                            ByVal dwMaxNameChars As Long, _
                                            ByVal pszColorBuff As Long, _
                                            ByVal cchMaxColorChars As Long, _
                                            ByVal pszSizeBuff As Long, _
                                            ByVal cchMaxSizeChars As Long) As Long
   
Public Declare Function StretchBlt _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal X As Long, _
                                          ByVal Y As Long, _
                                          ByVal nWidth As Long, _
                                          ByVal nHeight As Long, _
                                          ByVal hSrcDC As Long, _
                                          ByVal XSrc As Long, _
                                          ByVal YSrc As Long, _
                                          ByVal nSrcWidth As Long, _
                                          ByVal nSrcHeight As Long, _
                                          ByVal dwRop As Long) As Long

Public Declare Function SetLayout _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal dwLayout As Long) As Long

Public Declare Function TransparentBlt _
                         Lib "msimg32.dll" (ByVal hDC As Long, _
                                            ByVal X As Long, _
                                            ByVal Y As Long, _
                                            ByVal nWidth As Long, _
                                            ByVal nHeight As Long, _
                                            ByVal hSrcDC As Long, _
                                            ByVal XSrc As Long, _
                                            ByVal YSrc As Long, _
                                            ByVal nSrcWidth As Long, _
                                            ByVal nSrcHeight As Long, _
                                            ByVal crTransparent As Long) As Boolean

Public Declare Function CreateDIBSection8 _
                         Lib "gdi32.dll" _
                             Alias "CreateDIBSection" (ByVal hDC As Long, _
                                                       pBitmapInfo As BITMAPINFO8, _
                                                       ByVal un As Long, _
                                                       ByVal lplpVoid As Long, _
                                                       ByVal Handle As Long, _
                                                       ByVal dw As Long) As Long

Public Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function InflateRect _
                         Lib "user32.dll" (lpRect As RECT, _
                                           ByVal X As Long, _
                                           ByVal Y As Long) As Long

Public Declare Function OleTranslateColorByRef _
                         Lib "oleaut32.dll" _
                             Alias "OleTranslateColor" (ByVal lOleColor As Long, _
                                                        ByVal lHPalette As Long, _
                                                        ByVal lColorRef As Long) As Long

Public Declare Function SetBkColor _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal crColor As Long) As Long

Public Declare Function SetBkMode _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal nBkMode As Long) As Long

Public Declare Function SetPixelV _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal X As Long, _
                                          ByVal Y As Long, _
                                          ByVal crColor As Long) As Long

Public Declare Function CreateRectRgn _
                         Lib "gdi32.dll" (ByVal X1 As Long, _
                                          ByVal Y1 As Long, _
                                          ByVal X2 As Long, _
                                          ByVal Y2 As Long) As Long

Public Declare Function CombineRgn _
                         Lib "gdi32.dll" (ByVal hDestRgn As Long, _
                                          ByVal hSrcRgn1 As Long, _
                                          ByVal hSrcRgn2 As Long, _
                                          ByVal nCombineMode As Long) As Long

Public Declare Function RoundRect _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal Left As Long, _
                                          ByVal Top As Long, _
                                          ByVal Right As Long, _
                                          ByVal Bottom As Long, _
                                          ByVal EllipseWidth As Long, _
                                          ByVal EllipseHeight As Long) As Long

Public Declare Function CreatePolygonRgn _
                         Lib "gdi32.dll" (lpPoint As Any, _
                                          ByVal nCount As Long, _
                                          ByVal nPolyFillMode As Long) As Long

Public Declare Function CreateRoundRectRgn _
                         Lib "gdi32.dll" (ByVal X1 As Long, _
                                          ByVal Y1 As Long, _
                                          ByVal X2 As Long, _
                                          ByVal Y2 As Long, _
                                          ByVal X3 As Long, _
                                          ByVal Y3 As Long) As Long

Public Declare Function GetDIBits _
                         Lib "gdi32.dll" (ByVal aHDC As Long, _
                                          ByVal hBitmap As Long, _
                                          ByVal nStartScan As Long, _
                                          ByVal nNumScans As Long, _
                                          lpBits As Any, _
                                          lpBI As BITMAPINFO, _
                                          ByVal wUsage As Long) As Long

Public Declare Function SetDIBitsToDevice _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal X As Long, _
                                          ByVal Y As Long, _
                                          ByVal dx As Long, _
                                          ByVal dy As Long, _
                                          ByVal srcX As Long, _
                                          ByVal srcY As Long, _
                                          ByVal Scan As Long, _
                                          ByVal NumScans As Long, _
                                          Bits As Any, _
                                          BitsInfo As BITMAPINFO, _
                                          ByVal wUsage As Long) As Long

Public Declare Function StretchDIBits _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal X As Long, _
                                          ByVal Y As Long, _
                                          ByVal dx As Long, _
                                          ByVal dy As Long, _
                                          ByVal srcX As Long, _
                                          ByVal srcY As Long, _
                                          ByVal wSrcWidth As Long, _
                                          ByVal wSrcHeight As Long, _
                                          lpBits As Any, _
                                          lpBitsInfo As Any, _
                                          ByVal wUsage As Long, _
                                          ByVal dwRop As Long) As Long

Public Declare Function GetNearestColor _
                         Lib "gdi32.dll" (ByVal hDC As Long, _
                                          ByVal crColor As Long) As Long

Public Declare Function DwmIsCompositionEnabled Lib "dwmapi" (ByRef pfEnabled As Long) As Long

Public Declare Function GradientFill _
                         Lib "msimg32" (ByVal hDC As Long, _
                                        pVertex As TRIVERTEX, _
                                        ByVal dwNumVertex As Long, _
                                        pMesh As GRADIENT_RECT, _
                                        ByVal dwNumMesh As Long, _
                                        ByVal dwMode As Long) As Long

Public Type TRIVERTEX
    X                                       As Long
    Y                                   As Long
    Red                                 As Integer    'ushort value
    Green                               As Integer    'ushort value
    Blue                                As Integer    'ushort value
    Alpha                               As Integer    'ushort value

End Type

Public Type GRADIENT_RECT
    UpperLeft                               As Long
    LowerRight                          As Long
End Type

Public Const GRADIENT_FILL_RECT_V       As Long = &H1


'! -----------------------------------------------------------
'!  ‘ÛÌÍˆËˇ     :  FlatBorder
'!  œÂÂÏÂÌÌ˚Â  :  ByVal hwnd as long
'!  ŒÔËÒ‡ÌËÂ    :  ƒÂÎ‡ÂÚ ÍÌÓÔÍÛ ‚‰‡‚ÎÂÌÌÓÈ
'! -----------------------------------------------------------
Public Sub FlatBorderButton(ByVal lngHWnd As Long)

Dim TFlat                               As Long

    TFlat = GetWindowLong(lngHWnd, GWL_EXSTYLE)
    TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    SetWindowLong lngHWnd, GWL_EXSTYLE, TFlat
    SetWindowPos lngHWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE

End Sub

'! -----------------------------------------------------------
'!  ‘ÛÌÍˆËˇ     :  UnFlatBorderButton
'!  œÂÂÏÂÌÌ˚Â  :  ByVal hwnd as long
'!  ŒÔËÒ‡ÌËÂ    :  ƒÂÎ‡ÂÚ ÍÌÓÔÍÛ ÓÚ‰‡‚ÎÂÌÌÓÈ
'! -----------------------------------------------------------
Public Sub UnFlatBorderButton(ByVal lngHWnd As Long)

Dim TFlat                               As Long

    TFlat = GetWindowLong(lngHWnd, GWL_EXSTYLE)
    TFlat = TFlat And WS_EX_CLIENTEDGE
    SetWindowLong lngHWnd, GWL_EXSTYLE, TFlat
    SetWindowPos lngHWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE

End Sub

