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
Option Explicit

' ------------------------------------------------------
' Autor:    Leandro I. Ascierto
' Web:      www.leandroascierto.com.ar
' Fecha:    21 de Marzo de 2010
' ------------------------------------------------------
Private Declare Sub CopyMemory _
                     Lib "kernel32.dll" _
                         Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                                ByRef Source As Any, _
                                                ByVal Length As Long)

'Private Declare Function SetWindowLong _
                          Lib "user32.dll" _
                              Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function VirtualAlloc _
                          Lib "kernel32.dll" (ByRef lpAddress As Long, _
                                              ByVal dwSize As Long, _
                                              ByVal flAllocationType As Long, _
                                              ByVal flProtect As Long) As Long

Private Declare Function VirtualFree _
                          Lib "kernel32.dll" (ByRef lpAddress As Long, _
                                              ByVal dwSize As Long, _
                                              ByVal dwFreeType As Long) As Long

Private Declare Function SetScrollInfo _
                          Lib "user32.dll" (ByVal hWnd As Long, _
                                            ByVal n As Long, _
                                            lpcScrollInfo As SCROLLINFO, _
                                            ByVal bool As Boolean) As Long

Private Declare Function GetScrollInfo _
                          Lib "user32.dll" (ByVal hWnd As Long, _
                                            ByVal n As Long, _
                                            lpScrollInfo As SCROLLINFO) As Long

Private Declare Function ScrollWindowByNum& _
                          Lib "user32.dll" _
                              Alias "ScrollWindow" (ByVal hWnd As Long, _
                                                    ByVal XAmount As Long, _
                                                    ByVal YAmount As Long, _
                                                    ByVal lpRect As Long, _
                                                    ByVal lpClipRect As Long)

Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long

Private Declare Function ExcludeClipRect _
                          Lib "gdi32.dll" (ByVal hDC As Long, _
                                           ByVal X1 As Long, _
                                           ByVal Y1 As Long, _
                                           ByVal X2 As Long, _
                                           ByVal Y2 As Long) As Long

Private Declare Function GetFocus Lib "user32.dll" () As Long
Private Declare Function IsChild Lib "user32.dll" (ByVal hWndParent As Long, ByVal hWnd As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINT) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Type POINT
    X                                   As Long
    Y                                   As Long
End Type

Private Type SCROLLINFO
    cbSize                              As Long
    fMask                               As Long
    nMin                                As Long
    nMax                                As Long
    nPage                               As Long
    nPos                                As Long
    nTrackPos                           As Long
End Type

Private Const MEM_COMMIT                As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40
Private Const MEM_RELEASE               As Long = &H8000&
Private Const WM_MOUSEWHEEL             As Long = &H20A
Private Const WM_VSCROLL                As Long = &H115
Private Const WM_HSCROLL                As Long = &H114
Private Const WM_NCPAINT                As Long = &H85
Private Const WM_DESTROY                As Long = &H2
Private Const GWL_WNDPROC               As Long = -4
Private Const GWL_STYLE                 As Long = (-16)
Private Const WS_VSCROLL                As Long = &H200000
Private Const WS_HSCROLL                As Long = &H100000
Private Const GW_CHILD                  As Long = 5
Private Const GW_HWNDNEXT               As Long = 2
Private Const SB_HORZ                   As Long = 0
Private Const SB_VERT                   As Long = 1
Private Const SB_BOTH                   As Long = 3
Private Const SB_LINEDOWN               As Long = 1
Private Const SB_LINEUP                 As Long = 0
Private Const SB_PAGEDOWN               As Long = 3
Private Const SB_PAGEUP                 As Long = 2
Private Const SB_THUMBTRACK             As Long = 5
Private Const SB_ENDSCROLL              As Long = 8
Private Const SB_LEFT                   As Long = 6
Private Const SB_RIGHT                  As Long = 7
Private Const SIF_ALL                   As Long = &H17
Private Const SM_CYBORDER               As Long = 6

Public Enum EnuBorderStyle
    vbBSNone
    vbFixedSingle

End Enum

Private SI                              As SCROLLINFO
Private pASMWrapper                     As Long
Private PrevWndProc                     As Long
Private hSubclassedWnd                  As Long
Private mBorderSize                     As Long
Private OldPosH                         As Long
Private OldPosV                         As Long
Private m_hFocus                        As Long
Private m_AutoScrollToFocus             As Boolean
Private m_UseHandsCursor                As Boolean
Private m_HScrollVisible                As Boolean
Private m_VScrollVisible                As Boolean

Public Function WindowProc(ByVal hWnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
    WindowProc = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)

    Select Case uMsg

        Case WM_DESTROY
            Call StopSubclassing

        Case WM_VSCROLL, WM_HSCROLL

            Dim xScroll                 As Long

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

                    '
                Case SB_LEFT
                    SI.nPos = SI.nMin

                Case SB_RIGHT
                    SI.nPos = SI.nMax

            End Select

            SetScrollInfo hWnd, xScroll, SI, True
            GetScrollInfo hWnd, xScroll, SI

            If uMsg = WM_VSCROLL Then
                ScrollVerticalWindow -SI.nPos
            Else
                ScrollHorizontalWindow -SI.nPos

            End If

        Case WM_MOUSEWHEEL

            If m_VScrollVisible Then
                xScroll = SB_VERT
            Else

                If m_HScrollVisible Then
                    xScroll = SB_HORZ
                Else
                    Exit Function

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

            If xScroll = SB_VERT Then
                ScrollVerticalWindow -SI.nPos
            Else
                ScrollHorizontalWindow -SI.nPos

            End If

        Case WM_NCPAINT

            If UserControl.BorderStyle = vbFixedSingle Then

                Dim Rec                 As RECT
                Dim ClipRec             As RECT
                Dim hTheme              As Long
                Dim DC                  As Long

                DC = GetWindowDC(hWnd)
                GetWindowRect UserControl.hWnd, Rec
                Rec.Right = Rec.Right - Rec.Left
                Rec.Bottom = Rec.Bottom - Rec.Top
                Rec.Left = 0
                Rec.Top = 0
                hTheme = OpenThemeData(UserControl.hWnd, StrPtr("Edit"))

                If hTheme Then
                    ExcludeClipRect DC, mBorderSize, mBorderSize, Rec.Right - mBorderSize, Rec.Bottom - mBorderSize

                    If DrawThemeBackground(hTheme, DC, 0, 0, Rec, Rec) = 0 Then

                    End If

                    Call CloseThemeData(hTheme)

                End If

                ReleaseDC hWnd, DC

            End If

        Case Else

            On Error Resume Next

            Dim hFocus                  As Long

            If m_AutoScrollToFocus = False Then
                Exit Function

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
                    '----------
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

End Function

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    UserControl.BackColor = NewValue
    PropertyChanged "BackColor"

End Property

Public Property Get BorderStyle() As EnuBorderStyle
    BorderStyle = UserControl.BorderStyle

End Property

Public Property Let BorderStyle(ByVal NewValue As EnuBorderStyle)
    UserControl.BorderStyle = NewValue
    PropertyChanged "BorderStyle"

End Property

Public Property Get AutoScrollToFocus() As Boolean
    AutoScrollToFocus = m_AutoScrollToFocus

End Property

Public Property Let AutoScrollToFocus(ByVal NewValue As Boolean)
    m_AutoScrollToFocus = NewValue
    PropertyChanged "AutoScrollToFocus"

End Property

Public Property Get UseHandsCursor() As Boolean
    UseHandsCursor = m_UseHandsCursor

End Property

Public Property Let UseHandsCursor(ByVal NewValue As Boolean)
    m_UseHandsCursor = NewValue
    PropertyChanged "UseHandsCursor"

End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal Value As Boolean)
    UserControl.Enabled = Value
    PropertyChanged "Enabled"

End Property

Private Sub UserControl_InitProperties()
    m_AutoScrollToFocus = True
    m_UseHandsCursor = True

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    If m_UseHandsCursor Then
        If Button = 1 Then
            If m_VScrollVisible Or m_HScrollVisible Then
                SetCursor UserControl.MaskPicture
            End If
        End If
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    If m_UseHandsCursor Then
        If Button = 1 Then
            If m_VScrollVisible Or m_HScrollVisible Then
                SetCursor UserControl.MouseIcon
            End If
        End If
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

Static mY                               As Single
Static mX                               As Single

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

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Me.BackColor = .ReadProperty("BackColor", vbButtonFace)
        Me.BorderStyle = .ReadProperty("BorderStyle", vbFixedSingle)
        m_AutoScrollToFocus = .ReadProperty("AutoScrollToFocus", True)
        m_UseHandsCursor = .ReadProperty("UseHandsCursor", True)
        Me.Enabled = .ReadProperty("Enabled", True)

    End With

    If Ambient.UserMode Then
        SetSubclassing UserControl.hWnd

    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "BackColor", UserControl.BackColor, vbButtonFace
        .WriteProperty "BorderStyle", UserControl.BorderStyle, vbFixedSingle
        .WriteProperty "AutoScrollToFocus", m_AutoScrollToFocus, True
        .WriteProperty "UseHandsCursor", m_UseHandsCursor, True
        .WriteProperty "Enabled", UserControl.Enabled, True

    End With

End Sub

Private Function GetChildRectOfMe(hWnd As Long, ByRef SrcRect As RECT)

Dim PT                                  As POINT

    ClientToScreen UserControl.hWnd, PT
    Call GetWindowRect(hWnd, SrcRect)

    With SrcRect
        .Left = .Left - PT.X - OldPosH
        .Top = .Top - PT.Y - OldPosV
        .Right = .Right - PT.X - OldPosH
        .Bottom = .Bottom - PT.Y - OldPosV

    End With

End Function

Private Function IsChildOfMe(hWnd As Long) As Boolean

Dim hParent                             As Long

    hParent = GetParent(hWnd)

    Do While hParent <> 0

        If hParent = UserControl.hWnd Then
            IsChildOfMe = True
            Exit Do

        End If

        hParent = GetParent(hParent)
    Loop

End Function

' ActiveVB
Private Function SetSubclassing(ByVal hWnd As Long) As Boolean

'Setzt Subclassing, sofern nicht schon gesetzt
    If PrevWndProc = 0 Then
        If pASMWrapper <> 0 Then
            PrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, pASMWrapper)

            If PrevWndProc <> 0 Then
                hSubclassedWnd = hWnd
                SetSubclassing = True

            End If

        End If

    End If

End Function

' ActiveVB
Private Function StopSubclassing() As Boolean

'Stopt Subclassing, sofern gesetzt
    If hSubclassedWnd <> 0 Then
        If PrevWndProc <> 0 Then
            Call SetWindowLong(hSubclassedWnd, GWL_WNDPROC, PrevWndProc)
            hSubclassedWnd = 0
            PrevWndProc = 0
            StopSubclassing = True

        End If

    End If

End Function

Private Sub UserControl_Initialize()

Dim ASM(0 To 104)                       As Byte
Dim pVar                                As Long
Dim ThisClass                           As Long
Dim CallbackFunction                    As Long
Dim pVirtualFree

    SI.cbSize = Len(SI)
    SI.fMask = SIF_ALL
    mBorderSize = GetSystemMetrics(SM_CYBORDER)
    'www.ActiveVB.net
    'Virtuellen Speicher anfordern
    pASMWrapper = VirtualAlloc(ByVal 0&, 104, MEM_COMMIT, PAGE_EXECUTE_READWRITE)

    If pASMWrapper <> 0 Then
        'Instanzzeiger der Klasse auslesen
        ThisClass = ObjPtr(Me)
        'Zeiger auf die Callback-Funktion auslesen
        Call CopyMemory(pVar, ByVal ThisClass, 4)
        Call CopyMemory(CallbackFunction, ByVal (pVar + 1956), 4)
        'Zeiger auf die VirtualFree-Funktion ermitteln
        pVirtualFree = GetProcAddress(GetModuleHandle(StrPtr("kernel32.dll")), "VirtualFree")
        'ASM-Wrapper mit Maschinencode befüllen
        ASM(0) = &H90
        '&Hcc int 3 (Software Interrupt zum debuggen), &H90=nop (No Operation Point)
        ASM(1) = &HFF
        'inc (Zähler)
        ASM(2) = &H5
        ASM(7) = &H6A
        'push 0
        ASM(8) = &H0
        ASM(9) = &H54
        'push esp
        ASM(10) = &HFF
        'push (esp+18h) (laram)
        ASM(11) = &H74
        ASM(12) = &H24
        ASM(13) = &H18
        ASM(14) = &HFF
        'push (esp+18h) (wParam)
        ASM(15) = &H74
        ASM(16) = &H24
        ASM(17) = &H18
        ASM(18) = &HFF
        'push (esp+18h) (msg)
        ASM(19) = &H74
        ASM(20) = &H24
        ASM(21) = &H18
        ASM(22) = &HFF
        'push (esp+18h) (hwnd)
        ASM(23) = &H74
        ASM(24) = &H24
        ASM(25) = &H18
        ASM(26) = &H68
        'push Instanzzeiger
        ASM(31) = &HB8
        'mov eax, Adresse WindowProc
        ASM(36) = &HFF
        'call eax
        ASM(37) = &HD0
        ASM(38) = &HFF
        'dec (Zähler)
        ASM(39) = &HD
        ASM(44) = &HA1
        'mov eax, (Signal)
        ASM(49) = &H85
        'test eax, eax
        ASM(50) = &HC0
        ASM(51) = &H75
        'jne
        ASM(52) = &H4
        ASM(53) = &H58
        'pop eax (Rückgabewert)
        ASM(54) = &HC2
        'ret &H10
        ASM(55) = &H10
        ASM(56) = &H0
        ASM(57) = &HA1
        'mov eax, (Zähler)
        ASM(62) = &H85
        'test eax, eax
        ASM(63) = &HC0
        ASM(64) = &H74
        'je
        ASM(65) = &H4
        ASM(66) = &H58
        'pop eax (Rückgabewert)
        ASM(67) = &HC2
        'ret &H10
        ASM(68) = &H10
        ASM(69) = &H0
        ASM(70) = &H58
        'pop eax retval
        ASM(71) = &H59
        'pop ecx (Rücksprungzeiger)
        ASM(72) = &H58
        'pop eax hwnd
        ASM(73) = &H58
        'pop eax msg
        ASM(74) = &H58
        'pop eax wparam
        ASM(75) = &H58
        'pop eax lparam
        ASM(76) = &H68
        'push MEM_RELEASE
        ASM(77) = &H0
        ASM(78) = &H80
        ASM(79) = &H0
        ASM(80) = &H0
        ASM(81) = &H6A
        'push 0
        ASM(82) = &H0
        ASM(83) = &H68
        'push Zeiger auf den Wrapper
        ASM(88) = &H51
        'push ecx (Rücksprungzeiger)
        ASM(89) = &HB8
        'mov eax, VirtualFree Adresse
        ASM(94) = &HFF
        'jmp eax
        ASM(95) = &HE0
        ASM(96) = &H0
        'Speicher für Zähler
        ASM(97) = &H0
        ASM(98) = &H0
        ASM(99) = &H0
        ASM(100) = &H0
        'Speicher für Signal
        ASM(101) = &H0
        ASM(102) = &H0
        ASM(103) = &H0
        'Zähler Variable setzen
        pVar = pASMWrapper + 96
        Call CopyMemory(ASM(3), pVar, 4)
        Call CopyMemory(ASM(40), pVar, 4)
        Call CopyMemory(ASM(58), pVar, 4)
        'Flag Variable setzen
        pVar = pASMWrapper + 100
        Call CopyMemory(ASM(45), pVar, 4)
        'Wrapper Adresse setzen
        pVar = pASMWrapper
        Call CopyMemory(ASM(84), pVar, 4)
        'Instanzzeiger setzen
        pVar = ThisClass
        Call CopyMemory(ASM(27), pVar, 4)
        'Funktionszeiger setzen
        pVar = CallbackFunction
        Call CopyMemory(ASM(32), pVar, 4)
        'VirtualFree Adresse setzen
        pVar = pVirtualFree
        Call CopyMemory(ASM(90), pVar, 4)
        'fertigen Wrapper in DEP-kompatiblen Speicher kopieren
        Call CopyMemory(ByVal pASMWrapper, ASM(0), 104)

    End If

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next

    CheckScroll

End Sub

Public Sub Refresh()
    m_hFocus = 0
    CheckScroll

End Sub

Private Sub CheckScroll()

    On Error Resume Next

    Dim bWnd                            As Long
    Dim Rec                             As RECT
    Dim mRec                            As RECT

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

Private Sub ScrollVerticalWindow(ByVal NewPos As Long)
    ScrollWindowByNum UserControl.hWnd, 0&, NewPos - OldPosV, 0&, 0&
    OldPosV = NewPos

End Sub

Private Sub ScrollHorizontalWindow(ByVal NewPos As Long)
    ScrollWindowByNum UserControl.hWnd, NewPos - OldPosH, 0&, 0&, 0&
    OldPosH = NewPos

End Sub

Private Function GetLoWord(dw As Long) As Long

    If dw And &H8000& Then
        GetLoWord = &H8000 Or (dw And &H7FFF&)
    Else
        GetLoWord = dw And &HFFFF&

    End If

End Function

Private Function GetHiWord(dw As Long) As Long

    If dw And &H80000000 Then
        GetHiWord = (dw \ 65535) - 1
    Else
        GetHiWord = dw \ 65535

    End If

End Function

Private Sub UserControl_Show()
    Me.Refresh
    CheckScroll

End Sub

Private Sub UserControl_Terminate()

'Veranlasst das Freigeben des virtuellen Speichers
Dim Counter                             As Long
Dim Flag                                As Long

    On Error Resume Next

    If pASMWrapper <> 0 Then
        Call StopSubclassing
        'Zähler auslesen
        Call CopyMemory(Counter, ByVal (pASMWrapper + 104), 4)

        If Counter = 0 Then
            'Wrapper kann von VB aus gelöscht werden
            Call VirtualFree(ByVal pASMWrapper, 0, MEM_RELEASE)
        Else
            'Wrapper befindet sich noch innerhalb einer Rekursion und muss sich selbst löschen; Flag setzen
            Flag = 1
            Call CopyMemory(ByVal (pASMWrapper + 108), Flag, 4)

        End If

    End If

End Sub
