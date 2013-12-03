VERSION 5.00
Begin VB.UserControl ctlFontCombo 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ForeColor       =   &H80000008&
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   ToolboxBitmap   =   "ctlFontCombo.ctx":0000
   Begin VB.Timer TmrAutoText 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   90
   End
   Begin VB.Timer TmrFocus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3330
      Top             =   90
   End
   Begin VB.PictureBox PicPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   3330
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   810
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox PicList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   3285
      Begin VB.VScrollBar VScroll1 
         CausesValidation=   0   'False
         Height          =   2595
         Left            =   2760
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   240
      End
      Begin VB.Shape SelBox 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         DrawMode        =   14  'Copy Pen
         FillColor       =   &H8000000D&
         Height          =   285
         Left            =   0
         Top             =   600
         Width           =   2565
      End
   End
   Begin VB.Timer TmrOver 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2820
      Top             =   90
   End
   Begin VB.Shape FocusBox 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H8000000D&
      Height          =   285
      Left            =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   2565
   End
End
Attribute VB_Name = "ctlFontCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mEnabled                            As Boolean
Dim mBorderStyle                        As CfBdrStyle
Dim mSorted                             As Boolean
Dim inRct                               As Boolean
Dim tPos                                As Integer
Dim mButtonBackColor                    As Long
Dim mButtonForeColor                    As Long
Dim mButtonOverColor                    As Long
Dim mButtonBorderStyle                  As CfBdrStyle
Dim mShowFocus                          As Boolean

Private mListFont()                     As String
Private mListCount                      As Integer
Private mListPos                        As Integer
Private mUsedList()                     As String
Private mUsedCount                      As Integer
Private mUsedBackColor                  As Long
Private mUsedForeColor                  As Long
Private mRecent()                       As tpRecents
Private mRecentCount                    As Integer
Private mRecentMax                      As Integer
Private mRecentBackColor                As Long
Private mRecentForeColor                As Long
Private mPreviewText                    As String
Private mShowPreview                    As Boolean
Private mShowFontName                   As Boolean
Private mPreviewSize                    As Integer
Private mShowFontInCombo                As Boolean
Private mComboFontCount                 As Integer
Private mComboFontSize                  As Integer
Private mComboFontBold                  As Boolean
Private mComboFontItalic                As Boolean
Private mComboWidth                     As Single
Private mForeColor                      As Long
Private mBackColor                      As Long
Private mComboForeColor                 As Long
Private mComboBackColor                 As Long
Private mComboSelectColor               As Long
Private mUseMouseWheel                  As Boolean
Private mAutoText                       As String
Private CloseMe                         As Boolean

Dim doNothing                           As Boolean
Dim fList()                             As tpRecents
Dim fPos                                As Integer

Private bCancel                         As Boolean

Dim Resultat                            As Long
Dim Ident                               As Long
Dim Donnee                              As String
Dim TailleBuffer                        As Long
Dim Btn                                 As RECT
Dim uRct                                As RECT

Private MouseCoords                     As POINT

Dim mXPStyle                            As Boolean

Private Declare Function PtInRect _
                          Lib "user32.dll" (ByRef lpRect As RECT, _
                                            ByVal X As Long, _
                                            ByVal Y As Long) As Long

Private Declare Function GetMessage _
                          Lib "user32.dll" _
                              Alias "GetMessageA" (lpMsg As tMSG, _
                                                   ByVal hWnd As Long, _
                                                   ByVal wMsgFilterMin As Long, _
                                                   ByVal wMsgFilterMax As Long) As Long

Private Declare Function TranslateMessage Lib "user32.dll" (lpMsg As tMSG) As Long
Private Declare Function DispatchMessage _
                          Lib "user32.dll" _
                              Alias "DispatchMessageA" (lpMsg As tMSG) As Long

Private Declare Function RegCreateKey _
                          Lib "advapi32.dll" _
                              Alias "RegCreateKeyA" (ByVal hkey As Long, _
                                                     ByVal lpSubKey As String, _
                                                     phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
                          Lib "advapi32.dll" _
                              Alias "RegQueryValueExA" (ByVal hkey As Long, _
                                                        ByVal lpValueName As String, _
                                                        ByVal lpReserved As Long, _
                                                        lpType As Long, _
                                                        lpData As Any, _
                                                        lpcbData As Long) As Long

Private Declare Function RegSetValueEx _
                          Lib "advapi32.dll" _
                              Alias "RegSetValueExA" (ByVal hkey As Long, _
                                                      ByVal lpValueName As String, _
                                                      ByVal Reserved As Long, _
                                                      ByVal dwType As Long, _
                                                      lpData As Any, _
                                                      ByVal cbData As Long) As Long

Private Declare Function SetWindowLong _
                          Lib "user32.dll" _
                              Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Const HWND_TOP                  As Long = 0
Private Const HWND_BOTTOM               As Long = 1
Private Const HWND_NOTOPMOST = -2
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_NOPREFIX = &H800
Private Const DT_TABSTOP = &H80
Private Const WM_MOUSEWHEEL = 522
Private Const PM_REMOVE = &H1

Public Enum CfBdrStyle
    sNone = 0
    sRaised = &H1 Or &H4
    sSunken = &H2 Or &H8
    sBump = &H1 Or &H8
    sEtched = &H2 Or &H4
    sSmoothRaised = &H4
    sSmoothSunken = &H2

End Enum

Public Enum CfEdgeStyle
    edgeAll = &HF
    edgeLeft = &H2
    edgeTop = &H4
    edgeRight = &H1
    edgeBottom = &H8

End Enum

Public Enum HkeyLoc2
    'HKEY_CLASSES_ROOT = &H80000000
    'HKEY_CURRENT_USER = &H80000001
    'HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_DYN_DATA = &H80000004

End Enum

Private Enum eBtnState
    bUp = 0
    bOver = 1
    bDown = 2

End Enum

Private Enum sTxtPosition
    TopLeft = 0
    TopCenter = 1
    TopRight = 2
    MiddleLeft = 3
    MiddleCenter = 4
    MiddleRight = 5
    BottomLeft = 6
    BottomCenter = 7
    BottomRight = 8

End Enum

Private Enum HkeyLoc
    'HKEY_CLASSES_ROOT = &H80000000
    'HKEY_CURRENT_USER = &H80000001
    'HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_DYN_DATA = &H80000004

End Enum

Private Type tpRecents
    fName                                   As String
    fIndex                              As String
    fRecent                             As Boolean

End Type

Private Type tMSG
    hWnd                                    As Long
    nMsg                                As Long
    wParam                              As Long
    lParam                              As Long
    time                                As Long
    PT                                  As POINT

End Type

Private Msg                             As tMSG

Public Event SelectedFontChanged(NewFontName As String)
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event FontNotFound(FontName As String)

Public Function FontExist(Font2Find As String, _
                          Optional StartPos As Integer = 0) As Integer

Dim i                                   As Integer

    FontExist = -1

    For i = StartPos To mListCount

        If LCase$(mListFont(i)) Like LCase$(Font2Find) Then
            FontExist = i
            Exit For

        End If

    Next

End Function

Private Function DrawTheme(sClass As String, _
                           ByVal iPart As Long, _
                           ByVal iState As Long, _
                           rtRect As RECT) As Boolean

Dim hTheme                              As Long
Dim lResult                             As Long

    On Error GoTo NoXP

    hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))

    If (hTheme) Then
        lResult = DrawThemeBackground(hTheme, UserControl.hDC, iPart, iState, rtRect, rtRect)
        DrawTheme = IIf(lResult, False, True)
    Else
        DrawTheme = False

    End If

    Call CloseThemeData(hTheme)
    Exit Function
NoXP:
    DrawTheme = False

End Function

Public Property Get ButtonForeColor() As OLE_COLOR
    ButtonForeColor = mButtonForeColor

End Property

Public Property Get ListIndex() As Integer
    ListIndex = mListPos

End Property

Public Property Let ListIndex(ByVal vNewValue As Integer)
Attribute ListIndex.VB_MemberFlags = "400"
    mListPos = vNewValue

End Property

Public Property Get ListCount() As Integer
Attribute ListCount.VB_MemberFlags = "400"
    ListCount = mListCount

End Property

Private Sub PicList_KeyDown(KeyCode As Integer, Shift As Integer)
    UserControl_KeyDown KeyCode, Shift

End Sub

Private Sub PicList_KeyPress(KeyAscii As Integer)
    UserControl_KeyPress KeyAscii

End Sub

Private Sub PicList_KeyUp(KeyCode As Integer, Shift As Integer)
    UserControl_KeyUp KeyCode, Shift

End Sub

Private Sub PicList_LostFocus()
    PicList.Visible = False
    PicPreview.Visible = False
    TmrFocus.Enabled = False
    CloseMe = True

End Sub

Private Sub PicPreview_KeyDown(KeyCode As Integer, Shift As Integer)
    UserControl_KeyDown KeyCode, Shift

End Sub

Private Sub PicPreview_KeyPress(KeyAscii As Integer)
    UserControl_KeyPress KeyAscii

End Sub

Private Sub PicPreview_KeyUp(KeyCode As Integer, Shift As Integer)
    UserControl_KeyUp KeyCode, Shift

End Sub

Private Sub TmrAutoText_Timer()
    mAutoText = ""
    TmrAutoText.Enabled = False

End Sub

Private Sub TmrFocus_Timer()

Dim Focus                               As Long

    Focus = GetFocus

    'If (Focus <> PicList.hWnd And Focus <> UserControl.hWnd And _
     'Focus <> PicPreview.hWnd And Focus <> VScroll1.hWnd) Or CloseMe = True Then
    If (Focus <> UserControl.hWnd) Or CloseMe = True Then
        bCancel = True
        PicPreview.Visible = False
        PicList.Visible = False
        TmrFocus.Enabled = False
        CloseMe = True

    End If

End Sub

Private Sub PicList_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

Dim ti                                  As Integer

    ti = Int(Y \ (mComboFontSize * 2))

    If ti < mRecentCount Then
        mListPos = mRecent(ti).fIndex
    Else
        mListPos = fList(ti - mRecentCount).fIndex

    End If

End Sub

Private Sub PicList_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    On Local Error Resume Next

    Dim tFont                           As String
    Dim ti                              As Integer

    ti = Int(Y \ (mComboFontSize * 2))
    fPos = ti

    If TmrAutoText.Enabled = False Then
        SelBox.Move 0, CLng(Y \ (mComboFontSize * 2)) * (mComboFontSize * 2), PicList.ScaleWidth + 2, (mComboFontSize * 2) + 2

        If ti < mRecentCount Then
            tFont = mRecent(ti).fName
        Else
            tFont = fList(ti - mRecentCount).fName

        End If

        ShowFont tFont
        DoEvents

    End If

    If TmrAutoText.Enabled = True Then
        Exit Sub

    End If

    Do
        GetCursorPos MouseCoords

        If WindowFromPoint(MouseCoords.X, MouseCoords.Y) = PicList.hWnd Then
            If mUseMouseWheel = True Then
                GetMessage Msg, Parent.hWnd, 0, 0
                DispatchMessage Msg
                TranslateMessage Msg
                DoEvents

                With Msg

                    If .nMsg = WM_MOUSEWHEEL Then
                        If VScroll1.Value < VScroll1.Max And Sgn(.wParam) < 0 Then
                            If VScroll1.Value + 3 > VScroll1.Max Then
                                VScroll1.Value = VScroll1.Max
                            Else
                                VScroll1.Value = VScroll1.Value + 3

                            End If

                        Else

                            If VScroll1.Value - 3 < 0 Then
                                VScroll1.Value = 0
                            Else
                                VScroll1.Value = VScroll1.Value - 3

                            End If

                        End If

                    End If

                End With

            End If

        ElseIf CloseMe = False Then

            If WindowFromPoint(MouseCoords.X, MouseCoords.Y) = UserControl.hWnd Then Exit Do
            GetMessage Msg, Parent.hWnd, 0, 0
            DispatchMessage Msg
            TranslateMessage Msg
            DoEvents

            If Msg.nMsg = 513 Then
                CloseMe = True
                Exit Do

            End If

        Else
            Exit Do

        End If

        DoEvents
    Loop

End Sub

Private Sub PicList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetRecents mListFont(mListPos), mListPos
    PicList.Visible = False
    PicPreview.Visible = False
    TmrFocus.Enabled = False
    DrawControl , True
    CloseMe = True
    RaiseEvent SelectedFontChanged(mListFont(mListPos))

End Sub

Private Sub TmrOver_Timer()

Dim Pos                                 As POINT
Dim WFP                                 As Long

    GetCursorPos Pos
    WFP = WindowFromPoint(Pos.X, Pos.Y)

    If WFP <> Me.hWnd Then
        DrawControl bUp
        TmrOver.Enabled = False

    End If

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click

End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick

End Sub

Private Sub UserControl_GotFocus()

    If mShowFocus = True Then
        FocusBox.Visible = True
    Else
        FocusBox.Visible = False

    End If

End Sub

Private Sub UserControl_Initialize()
    CloseMe = False
    SetWindowLong PicList.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
    SetParent PicList.hWnd, 0
    SetWindowLong PicPreview.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
    SetParent PicPreview.hWnd, 0

End Sub

Private Sub UserControl_InitProperties()
    mEnabled = True
    mPreviewText = Ambient.DisplayName
    mBorderStyle = sSunken
    mButtonBorderStyle = sRaised
    mShowPreview = True
    mShowFontName = True
    mPreviewSize = 36
    mSorted = True
    mShowFontInCombo = True
    mComboFontCount = 20
    mComboFontSize = 8
    mComboFontBold = False
    mComboFontItalic = False
    mComboWidth = 250
    mRecentMax = 4
    mRecentBackColor = vbWindowBackground
    mRecentForeColor = vbWindowText
    mForeColor = vbWindowText
    mBackColor = vbWindowBackground
    mComboForeColor = vbWindowText
    mComboBackColor = vbWindowBackground
    mComboSelectColor = vbHighlight
    mButtonBackColor = vbButtonFace
    mButtonForeColor = vbButtonText
    mUseMouseWheel = False
    Set UserControl.Font = Ambient.Font
    mUsedBackColor = vbInfoBackground
    mUsedForeColor = vbInfoText
    mXPStyle = True
    mShowFocus = True

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

Dim kCode                               As String
Dim fI                                  As Integer
Dim kC                                  As Boolean

    If PicList.Visible = True Then

        Select Case KeyCode

            Case vbKeyUp

                If VScroll1.Value > 0 Then
                    VScroll1.Value = VScroll1.Value - 1

                End If

            Case vbKeyDown

                If VScroll1.Value < VScroll1.Max Then
                    VScroll1.Value = VScroll1.Value + 1

                End If

            Case vbKeyPageUp

                If VScroll1.Value - VScroll1.LargeChange > 0 Then
                    VScroll1.Value = VScroll1.Value - VScroll1.LargeChange
                Else
                    VScroll1.Value = VScroll1.Min

                End If

            Case vbKeyPageDown

                If VScroll1.Value + VScroll1.LargeChange < VScroll1.Max Then
                    VScroll1.Value = VScroll1.Value + VScroll1.LargeChange
                Else
                    VScroll1.Value = VScroll1.Max

                End If

            Case vbKeyHome
                VScroll1.Value = 0

            Case vbKeyEnd
                VScroll1.Value = VScroll1.Max

        End Select

        If mSorted = False Then
            Exit Sub

        End If

        kCode = LCase$(Chr$(KeyCode))

        If Asc(kCode) >= 97 Then
            If Asc(kCode) <= 122 Then
                kC = mAutoText = kCode

                If kC = False Then
                    mAutoText = mAutoText & kCode
                End If

                fI = FontExist(mAutoText & "*", mListPos + IIf(kC = True, 1, 0))

                ' check from current position
                If fI >= 0 Then
                    TmrAutoText.Enabled = False
                    mListPos = fI

                    If fI <= VScroll1.Max Then
                        VScroll1.Value = fI
                    Else
                        VScroll1.Value = VScroll1.Max

                    End If

                    SelBox.Move 0, (fI - VScroll1.Value + mRecentCount) * (mComboFontSize * 2), PicList.ScaleWidth + 2, (mComboFontSize * 2) + 2

                    If kC = False Then
                        TmrAutoText.Interval = 1500
                    Else
                        TmrAutoText.Interval = 800

                    End If

                    TmrAutoText.Enabled = True
                Else
                    fI = FontExist(mAutoText & "*")

                    'check from position 0
                    If fI >= 0 Then
                        TmrAutoText.Enabled = False
                        mListPos = fI
                        VScroll1.Value = fI
                        TmrAutoText.Interval = 1500
                        TmrAutoText.Enabled = True
                    Else
                        mAutoText = vbNullString
                    End If

                End If
            End If
        End If

    End If

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_LostFocus()
    FocusBox.Visible = False

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    On Local Error Resume Next

    If Button = 1 Then
        inRct = PtInRect(uRct, X, Y)

        If inRct = True Then
            DrawControl bDown, True
            DoEvents

            If PicList.Visible = False Then
                ShowList
            Else
                PicList.Visible = False
                PicPreview.Visible = False
                TmrFocus.Enabled = False
                CloseMe = True

            End If

        End If

    End If

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    If Button = 0 Then
        DrawControl bOver, True
        TmrOver.Enabled = True

    End If

    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    If Button = 1 Then
        If inRct = True Then DrawControl bUp
        inRct = False

    End If

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        mEnabled = .ReadProperty("Enabled", True)
        mPreviewText = .ReadProperty("PreviewText", Ambient.DisplayName)
        mBorderStyle = .ReadProperty("BorderStyle", sSunken)
        mButtonBorderStyle = .ReadProperty("ButtonBorderStyle", sRaised)
        mShowPreview = .ReadProperty("ShowPreview", True)
        mShowFontName = .ReadProperty("ShowFontName", True)
        mPreviewSize = .ReadProperty("PreviewSize", 36)
        mSorted = .ReadProperty("Sorted", True)
        mShowFontInCombo = .ReadProperty("ShowFontInCombo", True)
        mComboFontCount = .ReadProperty("ComboFontCount", 20)
        mComboFontSize = .ReadProperty("ComboFontSize", 8)
        mComboFontBold = .ReadProperty("ComboFontBold", False)
        mComboFontItalic = .ReadProperty("ComboFontItalic", False)
        mComboWidth = .ReadProperty("ComboWidth", 250)
        mRecentMax = .ReadProperty("RecentMax", 4)
        mRecentBackColor = .ReadProperty("RecentBackColor", vbWindowBackground)
        mRecentForeColor = .ReadProperty("RecentForeColor", vbWindowText)
        mForeColor = .ReadProperty("ForeColor", vbWindowText)
        mBackColor = .ReadProperty("BackColor", vbWindowBackground)
        mComboForeColor = .ReadProperty("ComboForeColor", vbWindowText)
        mComboBackColor = .ReadProperty("ComboBackColor", vbWindowBackground)
        mComboSelectColor = .ReadProperty("ComboSelectColor", vbHighlight)
        mButtonBackColor = .ReadProperty("ButtonBackColor", vbButtonFace)
        mButtonForeColor = .ReadProperty("ButtonForeColor", vbButtonText)
        mButtonOverColor = .ReadProperty("ButtonOverColor", vbBlue)
        mUseMouseWheel = .ReadProperty("UseMouseWheel", False)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        mUsedBackColor = .ReadProperty("UsedBackColor", vbInfoBackground)
        mUsedForeColor = .ReadProperty("UsedForeColor", vbInfoText)
        mXPStyle = .ReadProperty("XPStyle", True)
        mShowFocus = .ReadProperty("ShowFocus", True)

    End With

    UserControl.ForeColor = mForeColor
    UserControl.BackColor = mBackColor
    FocusBox.BackColor = mComboSelectColor
    ReDim mRecent(mRecentMax)

    If Ambient.UserMode = True Then
        FillList

        If mSorted = True Then SortList

    End If

    DrawControl , True

End Sub

Private Sub UserControl_Resize()

Dim tBdr                                As Single
Dim V                                   As Integer

    On Error Resume Next

    If mXPStyle = False Then
        V = 0

        Select Case mBorderStyle

            Case sNone
                tBdr = 0

            Case sSmoothRaised, sSmoothSunken
                tBdr = 1

            Case Else
                tBdr = 2

        End Select

    Else
        V = 2
        tBdr = 1

    End If

    UserControl.Height = ScaleY(TextHeight("X") + (tBdr * 2) + 4 + V, vbPixels, vbTwips)

    If UserControl.Width < 600 Then UserControl.Width = 600
    FocusBox.Move tBdr + 1, tBdr + 1, UserControl.ScaleWidth - tBdr - 20 + V, UserControl.ScaleHeight - (tBdr * 2) - 1
    SetRect uRct, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    SetRect Btn, UserControl.ScaleWidth - tBdr - 17, tBdr, UserControl.ScaleWidth - tBdr, UserControl.ScaleHeight - tBdr
    DrawControl bUp, True

End Sub

Public Property Get Enabled() As Boolean
    Enabled = mEnabled

End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    mEnabled = vNewValue
    DrawControl , True
    PropertyChanged "Enabled"

End Property

Private Sub DrawControl(Optional eDraw As eBtnState = bUp, _
                        Optional DrawAll As Boolean = False)

Dim Br                                  As Long
Dim tC                                  As Long

Static OldDr                            As eBtnState

Dim tCol                                As Long

    UserControl.Enabled = mEnabled
    mXPStyle = mXPStyle And DrawTheme("Button", 1, 1, Btn)

    If mXPStyle = False Then
        OleTranslateColor mButtonBackColor, 0, tC
        Br = CreateSolidBrush(tC)

        If mEnabled = False Then
            Cls
            FillRect UserControl.hDC, Btn, Br
            DrawEdge UserControl.hDC, uRct, mBorderStyle, edgeAll
            DrawEdge UserControl.hDC, Btn, mButtonBorderStyle, edgeAll
            tCol = UserControl.ForeColor
            UserControl.ForeColor = &H80000011
            UserControl.CurrentY = ((ScaleHeight - TextHeight("X")) / 2) - 1
            UserControl.CurrentX = 4

            If Ambient.UserMode = True And mListCount > 0 Then
                UserControl.Print mListFont(mListPos)
            Else
                UserControl.Print Ambient.DisplayName

            End If

            UserControl.ForeColor = tCol
            DrawArw
            DeleteObject Br
            Exit Sub

        End If

        If OldDr = eDraw Then
            If DrawAll = False Then
                Exit Sub
            End If
        End If

        UserControl.Cls
        UserControl.CurrentY = ((ScaleHeight - TextHeight("X")) / 2) - 1
        UserControl.CurrentX = 4

        If Ambient.UserMode = True And mListCount > 0 Then
            UserControl.Print mListFont(mListPos)
        Else
            UserControl.Print Ambient.DisplayName

        End If

        Select Case eDraw

            Case bUp
                DrawEdge UserControl.hDC, uRct, mBorderStyle, edgeAll
                FillRect UserControl.hDC, Btn, Br
                DrawEdge UserControl.hDC, Btn, ButtonBorderStyle, edgeAll

            Case bOver
                DrawEdge UserControl.hDC, uRct, mBorderStyle, edgeAll
                FillRect UserControl.hDC, Btn, Br
                DrawEdge UserControl.hDC, Btn, ButtonBorderStyle, edgeAll

            Case bDown
                DrawEdge UserControl.hDC, uRct, mBorderStyle, edgeAll
                FillRect UserControl.hDC, Btn, Br
                DrawEdge UserControl.hDC, Btn, InvBdr(ButtonBorderStyle), edgeAll

        End Select

        DeleteObject Br

        If eDraw = bOver Then
            DrawArw mButtonOverColor
        Else
            DrawArw

        End If

    Else
        UserControl.Cls

        If mEnabled = True Then

            Select Case eDraw

                Case bUp
                    DrawTheme "ComboBox", 2, 1, uRct
                    DrawTheme "ComboBox", 1, 1, Btn

                Case bOver
                    DrawTheme "ComboBox", 2, 2, uRct
                    DrawTheme "ComboBox", 1, 2, Btn

                Case bDown
                    DrawTheme "ComboBox", 2, 3, uRct
                    DrawTheme "ComboBox", 1, 3, Btn

            End Select

        Else
            DrawTheme "ComboBox", 2, 4, uRct
            DrawTheme "ComboBox", 1, 4, Btn
            tCol = UserControl.ForeColor
            UserControl.ForeColor = &H80000011
            UserControl.CurrentY = ((ScaleHeight - TextHeight("X")) / 2) - 1
            UserControl.CurrentX = 4
            UserControl.Print mListFont(mListPos)
            UserControl.ForeColor = tCol
            Exit Sub

        End If

        UserControl.CurrentY = ((ScaleHeight - TextHeight("X")) / 2) - 1
        UserControl.CurrentX = 4

        If Ambient.UserMode = True And mListCount > 0 Then
            UserControl.Print mListFont(mListPos)
        Else
            UserControl.Print Ambient.DisplayName

        End If

    End If

    OldDr = eDraw
    Refresh

End Sub

Private Sub DrawArw(Optional ArrowColor As Long = -1)

Dim ColUp                               As Long
Dim tCol                                As Long

    If ArrowColor = -1 Then
        tCol = mButtonForeColor
    Else
        tCol = ArrowColor

    End If

    If mEnabled = False Then
        OleTranslateColor vbGrayText, 0, ColUp
    Else
        OleTranslateColor tCol, 0, ColUp

    End If

    SetPixel UserControl.hDC, Btn.Left - 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left - 2 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 2 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left - 3 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 3 + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + (Btn.Right - Btn.Left) \ 2, Btn.Top - 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left - 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left - 2 + (Btn.Right - Btn.Left) \ 2, Btn.Top + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 2 + (Btn.Right - Btn.Left) \ 2, Btn.Top + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + (Btn.Right - Btn.Left) \ 2, Btn.Top + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left - 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top + 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + 1 + (Btn.Right - Btn.Left) \ 2, Btn.Top + 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + (Btn.Right - Btn.Left) \ 2, Btn.Top + 1 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    SetPixel UserControl.hDC, Btn.Left + (Btn.Right - Btn.Left) \ 2, Btn.Top + 2 + (Btn.Bottom - Btn.Top) \ 2, ColUp
    Refresh

End Sub

Private Sub ShowList()

Dim cb                                  As RECT

    CloseMe = False
    GetWindowRect UserControl.hWnd, cb
    tPos = mListPos
    PicList.Width = ScaleX(mComboWidth, vbPixels, vbTwips)
    PicList.Height = ScaleY(((mComboFontSize * 2) * (mComboFontCount + mRecentCount) + 2), vbPixels, vbTwips)
    VScroll1.Move PicList.ScaleWidth - 18, (mComboFontSize * 2) * mRecentCount, 18, PicList.ScaleHeight - ((mComboFontSize * 2) * mRecentCount)

    If cb.Bottom + (PicList.Height / Screen.TwipsPerPixelY) < Screen.Height / Screen.TwipsPerPixelY Then
        SetWindowPos PicList.hWnd, HWND_TOP, cb.Left, cb.Bottom, PicList.Width / Screen.TwipsPerPixelX, PicList.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Else
        SetWindowPos PicList.hWnd, HWND_TOP, cb.Left, cb.Top - (PicList.Height / Screen.TwipsPerPixelY), PicList.Width / Screen.TwipsPerPixelX, PicList.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW

    End If

    SetWindowPos PicPreview.hWnd, HWND_TOP, (PicList.Left + PicList.Width) / Screen.TwipsPerPixelX, (PicList.Top / Screen.TwipsPerPixelY), PicPreview.Width / Screen.TwipsPerPixelX, PicPreview.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE
    fPos = mListPos
    DrawList
    UserControl.SetFocus
    TmrFocus.Enabled = True

End Sub

Private Sub ShowFont(fName As String)

Dim tRc                                 As RECT
Dim tStr                                As String

Static OldFont                          As String

Dim Br                                  As Long
Dim tC                                  As Long

    If fName = "" Or mShowPreview = False Then
        Exit Sub

    End If

    If Trim$(mPreviewText) = "" Then
        tStr = fName
    Else
        tStr = mPreviewText

    End If

    If fName <> OldFont Then
        OldFont = fName
    Else
        Exit Sub

    End If

    PicPreview.FontName = fName
    PicPreview.FontSize = mPreviewSize
    PicPreview.FontBold = False
    PicPreview.FontItalic = False
    PicPreview.Cls
    PicPreview.Height = (PicPreview.TextHeight(tStr) * Screen.TwipsPerPixelY) + 200
    PicPreview.Width = (PicPreview.TextWidth(tStr) * Screen.TwipsPerPixelX) + 200

    If PicPreview.Width > Screen.Width / 2 Then PicPreview.Width = Screen.Width / 2
    If Screen.Width - (PicList.Left + PicList.Width) < PicPreview.Width Then
        PicPreview.Left = PicList.Left - PicPreview.Width
    Else
        PicPreview.Left = PicList.Left + PicList.Width

    End If

    SetRect tRc, 0, 0, PicPreview.ScaleWidth, PicPreview.ScaleHeight
    DrawTxt PicPreview.hDC, tStr, tRc, MiddleCenter, False, True, True

    If mShowFontName = True Then
        OleTranslateColor mComboForeColor, 0, tC
        Br = CreateSolidBrush(vbBlack)
        PicPreview.FontName = "MS Sans Serif"
        PicPreview.FontSize = 8
        PicPreview.FontBold = False
        PicPreview.FontItalic = False
        PicPreview.Height = PicPreview.Height + (PicPreview.TextHeight("X") * Screen.TwipsPerPixelY)
        SetRect tRc, -1, PicPreview.ScaleHeight - PicPreview.TextHeight("X") - 2, PicPreview.ScaleWidth + 1, PicPreview.ScaleHeight + 1
        DrawTxt PicPreview.hDC, fName, tRc, MiddleCenter
        FrameRect PicPreview.hDC, tRc, Br
        DeleteObject Br

    End If

    If PicPreview.Visible = False Then PicPreview.Visible = True
    PicPreview.Refresh

End Sub

Private Sub DrawTxt(ObjhDC As Long, _
                    oText As String, _
                    TxtRect As RECT, _
                    mPosition As sTxtPosition, _
                    Optional MultiLine As Boolean = False, _
                    Optional WordWrap As Boolean = False, _
                    Optional WordEllipsis As Boolean = False)

Dim tFormat                             As Long

    Select Case mPosition

        Case TopLeft
            tFormat = DT_TOP + DT_LEFT

        Case TopCenter
            tFormat = DT_TOP + DT_CENTER

        Case TopRight
            tFormat = DT_TOP + DT_RIGHT

        Case MiddleLeft
            tFormat = DT_VCENTER + DT_LEFT

        Case MiddleCenter
            tFormat = DT_VCENTER + DT_CENTER

        Case MiddleRight
            tFormat = DT_VCENTER + DT_RIGHT

        Case BottomLeft
            tFormat = DT_BOTTOM + DT_LEFT

        Case BottomCenter
            tFormat = DT_BOTTOM + DT_CENTER

        Case BottomRight
            tFormat = DT_BOTTOM + DT_RIGHT

    End Select

    If MultiLine = False Then
        tFormat = tFormat + DT_SINGLELINE
    End If
    If WordWrap = True And MultiLine = True Then
        tFormat = tFormat + DT_WORDBREAK
    End If
    If WordEllipsis = True Then
        tFormat = tFormat + DT_WORD_ELLIPSIS
    End If
    tFormat = tFormat + DT_NOCLIP
    DrawText ObjhDC, oText, Len(oText), TxtRect, tFormat

End Sub

Private Sub mgSort(ByVal pStart As Long, ByVal pEnd As Long)

Dim m                                   As Long
Dim n                                   As Long
Dim tStr1                               As String

    m = pStart
    n = pEnd
    tStr1 = LCase$(mListFont((pStart + pEnd) \ 2))
    Do
        Do While LCase$(mListFont(m)) < tStr1
            m = m + 1
        Loop

        Do While LCase$(mListFont(n)) > tStr1
            n = n - 1
        Loop

        If m <= n Then
            SwapStrings mListFont(m), mListFont(n)
            m = m + 1
            n = n - 1

        End If

    Loop Until m > n

    If pStart < n Then Call mgSort(pStart, n)
    If m < pEnd Then Call mgSort(m, pEnd)

End Sub

Private Function ReadValue(MyHkey As HkeyLoc, _
                           myKey As String, _
                           MyValue As String, _
                           Optional ByVal MyDefaultData As String = "") As String

    On Error GoTo ReadValue_Error

    Resultat = 0
    Ident = 0
    TailleBuffer = 0
    Resultat = RegCreateKey(MyHkey, myKey, Ident)

    If Resultat <> 0 Then
        Exit Function

    End If

    Resultat = RegQueryValueEx(Ident, MyValue, 0&, 1, 0&, TailleBuffer)

    If TailleBuffer < 2 Then
        ReadValue = MyDefaultData
        Exit Function

    End If

    Donnee = String$(TailleBuffer + 1, " ")
    Resultat = RegQueryValueEx(Ident, MyValue, 0&, 1, ByVal Donnee, TailleBuffer)
    Donnee = Left$(Donnee, TailleBuffer - 1)
    ReadValue = Donnee

    On Error GoTo 0

ReadValue_Error:
    Exit Function

End Function

Private Sub SetRecents(Optional CurRecent As String, Optional CurIndex As Integer)

Dim m                                   As Integer
Dim n                                   As Integer
Dim TmpLast()                           As tpRecents
Dim a%, B%
Dim myLast                              As tpRecents

    For n = 0 To mRecentMax - 1

        If mRecent(0).fName = CurRecent Then
            If n <> 0 Then
                myLast = mRecent(0)
                mRecent(0) = mRecent(n)
                mRecent(n) = myLast

            End If

            Exit For

        End If

    Next
    ReDim TmpLast(mRecentMax)

    If CurRecent = "" Then
        TmpLast = mRecent
    Else

        For n = 1 To mRecentMax
            myLast = mRecent(n - 1)

            If LenB(Trim$(myLast.fName)) > 0 Then
                TmpLast(n) = myLast

            End If

        Next
        TmpLast(0).fName = CurRecent
        TmpLast(0).fIndex = CurIndex

    End If

    For a% = 0 To mRecentMax
        For B% = 0 To mRecentMax

            If B% <> a% Then
                If LenB(TmpLast(a%).fName) > 0 Then
                    If TmpLast(a%).fName = TmpLast(B%).fName Then
                        TmpLast(B%).fName = ""
                        B% = B% - 1
                    End If
                End If
            End If

        Next
    Next
    m = 0
    ReDim mRecent(mRecentMax)

    For n = 0 To mRecentMax - 1

        If LenB(Trim$(TmpLast(n).fName)) > 0 Then
            mRecent(m).fName = TmpLast(n).fName
            mRecent(m).fIndex = TmpLast(n).fIndex
            mRecent(m).fRecent = True
            m = m + 1

        End If

    Next
    mRecentCount = m

End Sub

Private Function SetValue(MyHkey As HkeyLoc, _
                          myKey As String, _
                          MyValue As String, _
                          ByVal MyData As String)

    On Error GoTo SetValue_Error

    Resultat = 0
    Ident = 0
    TailleBuffer = 0
    Resultat = RegCreateKey(MyHkey, myKey, Ident)

    If Resultat = 0 Then
        Resultat = RegSetValueEx(Ident, MyValue, 0&, 1, ByVal MyData, Len(MyData) + 1)

    End If

    On Error GoTo 0

    Exit Function
SetValue_Error:

End Function

Private Sub SortList()

Dim n                                   As Long
Dim tStart                              As Long
Dim tEnd                                As Long
Dim bStr1                               As String
Dim bStr2                               As String
Dim qRec                                As Long

    mgSort 0, mListCount
    tStart = 0
    Do
        bStr1 = mListFont(tStart)
        qRec = 0

        For n = tStart To mListCount
            bStr2 = mListFont(n)

            If LCase$(bStr1) = LCase$(bStr2) Then
                qRec = qRec + 1
            Else
                Exit For

            End If

        Next
        tEnd = tStart + qRec
        mgSort tStart, tEnd - 1
        tStart = tEnd
    Loop While tEnd < mListCount

End Sub

Private Sub SwapStrings(String1 As String, String2 As String)

Dim tHold                               As Long

    CopyMemory tHold, ByVal VarPtr(String1), 4
    CopyMemory ByVal VarPtr(String1), ByVal VarPtr(String2), 4
    CopyMemory ByVal VarPtr(String2), tHold, 4

End Sub

Private Sub DrawList()
    On Local Error Resume Next

    Dim i                               As Integer
    Dim Br                              As Long
    Dim tC                              As Long
    Dim rct                             As RECT

    OleTranslateColor mRecentBackColor, 0, tC
    Br = CreateSolidBrush(tC)
    PicList.Cls
    doNothing = True
    VScroll1.Max = mListCount - mComboFontCount + mRecentCount
    VScroll1.LargeChange = ((mListCount + mRecentCount) \ mComboFontCount) + 1
    SetList
    SetRect rct, 0, 0, PicList.ScaleWidth, mRecentCount * (mComboFontSize * 2)
    FillRect PicList.hDC, rct, Br
    DeleteObject Br
    OleTranslateColor mUsedBackColor, 0, tC
    Br = CreateSolidBrush(tC)
    PicList.Line (0, mRecentCount * (mComboFontSize * 2))-(PicList.ScaleWidth, mRecentCount * (mComboFontSize * 2))

    For i = 0 To mRecentCount - 1
        PicList.CurrentX = 2
        PicList.CurrentY = (i * (mComboFontSize * 2)) + 2

        If mShowFontInCombo = True Then PicList.FontName = mRecent(i).fName
        PicList.FontSize = mComboFontSize
        PicList.FontItalic = mComboFontItalic
        PicList.FontBold = mComboFontBold

        If IsUsed(mRecent(i).fName) = False Then
            PicList.ForeColor = mRecentForeColor
        Else
            SetRect rct, 0, i * (mComboFontSize * 2), PicList.ScaleWidth, (i + 1) * (mComboFontSize * 2)
            FillRect PicList.hDC, rct, Br
            PicList.ForeColor = mUsedForeColor

        End If

        PicList.Print mRecent(i).fName
    Next

    For i = 0 To mComboFontCount - 1

        If IsUsed(fList(i).fName) = False Then
            PicList.ForeColor = mComboForeColor
        Else
            SetRect rct, 0, (i * (mComboFontSize * 2)) + ((mComboFontSize * 2) * mRecentCount) + 2, PicList.ScaleWidth, ((i + 1) * (mComboFontSize * 2)) + ((mComboFontSize * 2) * mRecentCount)
            FillRect PicList.hDC, rct, Br
            PicList.ForeColor = mUsedForeColor

        End If

        PicList.CurrentX = 2
        PicList.CurrentY = (i * (mComboFontSize * 2)) + 2 + ((mComboFontSize * 2) * mRecentCount)

        If mShowFontInCombo = True Then PicList.FontName = fList(i).fName
        PicList.FontSize = mComboFontSize
        PicList.FontItalic = mComboFontItalic
        PicList.FontBold = mComboFontBold
        PicList.Print fList(i).fName
    Next
    DeleteObject Br
    SelBox.Move 0, (fPos - VScroll1.Value + mRecentCount) * (mComboFontSize * 2), PicList.ScaleWidth + 2, (mComboFontSize * 2) + 2
    doNothing = False

End Sub

Private Function IsUsed(FontName As String) As Boolean

Dim i                                   As Integer
Dim F                                   As Boolean

    For i = 0 To mUsedCount - 1

        If LCase$(mUsedList(i)) = LCase$(FontName) Then
            F = True
            Exit For

        End If

    Next
    IsUsed = F

End Function

Private Sub SetList()

Dim i                                   As Integer
Dim RecQ                                As Integer
Dim Start                               As Integer

    ReDim fList(mComboFontCount)
    Start = fPos

    If Start + mComboFontCount > mListCount Then
        Start = mListCount - mComboFontCount

    End If

    VScroll1.Value = Start

    For i = Start To Start + mComboFontCount - RecQ
        fList(RecQ).fName = mListFont(i)
        fList(RecQ).fIndex = i
        fList(RecQ).fRecent = False
        RecQ = RecQ + 1
    Next

End Sub

Public Property Get PreviewText() As String
    PreviewText = mPreviewText

End Property

Public Property Let PreviewText(ByVal vNewValue As String)
    mPreviewText = vNewValue
    PropertyChanged "PreviewText"

End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Enabled", mEnabled, True
        .WriteProperty "PreviewText", mPreviewText, Ambient.DisplayName
        .WriteProperty "BorderStyle", mBorderStyle, sSunken
        .WriteProperty "ButtonBorderStyle", mButtonBorderStyle, sRaised
        .WriteProperty "ShowPreview", mShowPreview, True
        .WriteProperty "ShowFontName", mShowFontName, True
        .WriteProperty "PreviewSize", mPreviewSize, 36
        .WriteProperty "Sorted", mSorted, True
        .WriteProperty "ShowFontInCombo", mShowFontInCombo, True
        .WriteProperty "ComboFontCount", mComboFontCount, 20
        .WriteProperty "ComboFontSize", mComboFontSize, 8
        .WriteProperty "ComboFontBold", mComboFontBold, False
        .WriteProperty "ComboFontItalic", mComboFontItalic, False
        .WriteProperty "ComboWidth", mComboWidth, 250
        .WriteProperty "RecentMax", mRecentMax, 4
        .WriteProperty "RecentBackColor", mRecentBackColor, vbWindowBackground
        .WriteProperty "RecentForeColor", mRecentForeColor, vbWindowText
        .WriteProperty "ForeColor", mForeColor, vbWindowText
        .WriteProperty "BackColor", mBackColor, vbWindowBackground
        .WriteProperty "ComboForeColor", mComboForeColor, vbWindowText
        .WriteProperty "ComboBackColor", mComboBackColor, vbWindowBackground
        .WriteProperty "ComboSelectColor", mComboSelectColor, vbHighlight
        .WriteProperty "ButtonBackColor", mButtonBackColor, vbButtonFace
        .WriteProperty "ButtonForeColor", mButtonForeColor, vbButtonText
        .WriteProperty "ButtonOverColor", mButtonOverColor, vbBlue
        .WriteProperty "UseMouseWheel", mUseMouseWheel, False
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "UsedBackColor", mUsedBackColor, vbInfoBackground
        .WriteProperty "UsedForeColor", mUsedForeColor, vbInfoText
        .WriteProperty "XPStyle", mXPStyle, True
        .WriteProperty "ShowFocus", mShowFocus, True

    End With

End Sub

Public Property Get BorderStyle() As CfBdrStyle
    BorderStyle = mBorderStyle

End Property

Public Property Let BorderStyle(ByVal vNewValue As CfBdrStyle)
    mBorderStyle = vNewValue
    UserControl_Resize
    PropertyChanged "BorderStyle"

End Property

Private Function InvBdr(Bdr As CfBdrStyle) As CfBdrStyle

    Select Case Bdr

        Case sNone
            InvBdr = sNone

        Case sRaised
            InvBdr = sSunken

        Case sSunken
            InvBdr = sRaised

        Case sBump
            InvBdr = sEtched

        Case sEtched
            InvBdr = sBump

        Case sSmoothRaised
            InvBdr = sSmoothSunken

        Case sSmoothSunken
            InvBdr = sSmoothRaised

    End Select

End Function

Public Property Get ShowPreview() As Boolean
    ShowPreview = mShowPreview

End Property

Public Property Let ShowPreview(ByVal vNewValue As Boolean)
    mShowPreview = vNewValue
    PropertyChanged "ShowPreview"

End Property

Public Property Get PreviewSize() As Integer
    PreviewSize = mPreviewSize

End Property

Public Property Let PreviewSize(ByVal vNewValue As Integer)

    If vNewValue > 10 Then
        If vNewValue < 200 Then
            mPreviewSize = vNewValue
            PropertyChanged "PreviewSize"
        End If
    End If

End Property

Public Property Get Sorted() As Boolean
    Sorted = mSorted

End Property

Public Property Let Sorted(ByVal vNewValue As Boolean)

Dim i                                   As Integer
Dim fI                                  As Integer

    mSorted = vNewValue

    If Ambient.UserMode = True Then
        FillList

        If mSorted = True Then SortList

        For i = 0 To mRecentCount - 1
            fI = FontExist(mRecent(i).fName)
            mRecent(i).fIndex = fI
        Next

    End If

    DrawControl , True
    PropertyChanged "Sorted"

End Property

Public Property Get ListFont(Index As Integer) As String
    ListFont = mListFont(Index)

End Property

Public Property Get SelectedFont() As String
    SelectedFont = mListFont(mListPos)

End Property

Public Property Let SelectedFont(ByVal vNewValue As String)
Attribute SelectedFont.VB_MemberFlags = "400"

Dim i                                   As Integer

    i = FontExist(vNewValue)

    If i > -1 Then
        mListPos = i
        RaiseEvent SelectedFontChanged(mListFont(mListPos))
        DrawControl , True
    Else
        RaiseEvent FontNotFound(vNewValue)

    End If

End Property

Public Property Get ShowFontInCombo() As Boolean
    ShowFontInCombo = mShowFontInCombo

End Property

Public Property Let ShowFontInCombo(ByVal vNewValue As Boolean)
    mShowFontInCombo = vNewValue
    PropertyChanged "ShowFontInCombo"

End Property

Public Property Get ComboFontCount() As Integer
    ComboFontCount = mComboFontCount

End Property

Public Property Let ComboFontCount(ByVal vNewValue As Integer)

    If vNewValue > 50 Or vNewValue < 5 Then vNewValue = 20
    mComboFontCount = vNewValue
    PropertyChanged "ComboFontCount"

End Property

Public Property Get ComboFontSize() As Integer
    ComboFontSize = mComboFontSize

End Property

Public Property Let ComboFontSize(ByVal vNewValue As Integer)

    If vNewValue > 50 Or vNewValue < 6 Then vNewValue = 8
    mComboFontSize = vNewValue
    PropertyChanged "ComboFontSize"

End Property

Public Property Get ComboWidth() As Single
    ComboWidth = mComboWidth

End Property

Public Property Let ComboWidth(ByVal vNewValue As Single)
    mComboWidth = vNewValue
    PropertyChanged "ComboWidth"

End Property

Public Property Get RecentMax() As Integer
    RecentMax = mRecentMax

End Property

Public Property Let RecentMax(ByVal vNewValue As Integer)
Attribute RecentMax.VB_Description = "If you don't want to use Recents feature enter 0"
    mRecentMax = vNewValue
    PropertyChanged "RecentMax"

End Property

Public Property Get RecentBackColor() As OLE_COLOR
    RecentBackColor = mRecentBackColor

End Property

Public Property Let RecentBackColor(ByVal vNewValue As OLE_COLOR)
    mRecentBackColor = vNewValue
    PropertyChanged "RecentBackColor"

End Property

Public Property Get RecentForeColor() As OLE_COLOR
    RecentForeColor = mRecentForeColor

End Property

Public Property Let RecentForeColor(ByVal vNewValue As OLE_COLOR)
    mRecentForeColor = vNewValue
    PropertyChanged "RecentForeColor"

End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor

End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    mForeColor = vNewValue
    UserControl.ForeColor = mForeColor
    DrawControl , True
    PropertyChanged "ForeColor"

End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor

End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    mBackColor = vNewValue
    UserControl.BackColor = mBackColor
    DrawControl , True
    PropertyChanged "BackColor"

End Property

Public Property Get ComboForeColor() As OLE_COLOR
    ComboForeColor = mComboForeColor

End Property

Public Property Let ComboForeColor(ByVal vNewValue As OLE_COLOR)
    mComboForeColor = vNewValue
    PropertyChanged "ComboForeColor"

End Property

Public Property Get ComboBackColor() As OLE_COLOR
    ComboBackColor = mComboBackColor

End Property

Public Property Let ComboBackColor(ByVal vNewValue As OLE_COLOR)
    mComboBackColor = vNewValue
    PropertyChanged "ComboBackColor"

End Property

Public Property Get ComboSelectColor() As OLE_COLOR
    ComboSelectColor = mComboSelectColor

End Property

Public Property Let ComboSelectColor(ByVal vNewValue As OLE_COLOR)
    mComboSelectColor = vNewValue
    PropertyChanged "ComboSelectColor"

End Property

Public Sub LoadRecentFonts(MyHkey As HkeyLoc2, _
                           MyGroup As String, _
                           MySection As String, _
                           myKey As String)

Dim i                                   As Integer
Dim fN                                  As String
Dim fI                                  As Integer

    ReDim mRecent(mRecentMax)

    For i = 0 To mRecentMax - 1
        fN = ReadValue(MyHkey, MyGroup & "\" & MySection & "\" & myKey, "RecentFontName" & i + 1, "")
        fI = FontExist(fN)

        If fI > -1 Then
            mRecent(i).fName = fN
            mRecent(i).fIndex = fI

        End If

    Next
    SetRecents

End Sub

Public Sub SaveRecentFonts(MyHkey As HkeyLoc2, _
                           MyGroup As String, _
                           MySection As String, _
                           myKey As String)

Dim i                                   As Integer

    For i = 0 To mRecentCount - 1
        SetValue MyHkey, MyGroup & "\" & MySection & "\" & myKey, "RecentFontName" & i + 1, mRecent(i).fName
    Next

End Sub

Public Property Get UseMouseWheel() As Boolean
    UseMouseWheel = mUseMouseWheel

End Property

Public Property Let UseMouseWheel(ByVal vNewValue As Boolean)
    mUseMouseWheel = vNewValue
    PropertyChanged "UseMouseWheel"

End Property

Public Sub ClearRecent()
    mRecentCount = 0
    ReDim mRecent(0)

End Sub

Public Property Get Font() As StdFont
    Set Font = UserControl.Font

End Property

Public Property Set Font(ByVal vNewValue As StdFont)
    Set UserControl.Font = vNewValue
    UserControl_Resize
    PropertyChanged "Font"

End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd

End Property

Public Property Get ShowFontName() As Boolean
    ShowFontName = mShowFontName

End Property

Public Property Let ShowFontName(ByVal vNewValue As Boolean)
    mShowFontName = vNewValue
    PropertyChanged "ShowFontName"

End Property

Private Sub FillList()

Dim i                                   As Integer

    mListCount = Screen.FontCount - 1
    ReDim mListFont(mListCount)

    For i = 0 To Screen.FontCount - 1
        mListFont(i) = Screen.Fonts(i)
    Next

End Sub

Public Property Get ComboFontBold() As Boolean
    ComboFontBold = mComboFontBold

End Property

Public Property Let ComboFontBold(ByVal vNewValue As Boolean)
Attribute ComboFontBold.VB_MemberFlags = "400"
    mComboFontBold = vNewValue
    PropertyChanged "ComboFontBold"

End Property

Public Property Get ComboFontItalic() As Boolean
    ComboFontItalic = mComboFontItalic

End Property

Public Property Let ComboFontItalic(ByVal vNewValue As Boolean)
Attribute ComboFontItalic.VB_MemberFlags = "400"
    mComboFontItalic = vNewValue
    PropertyChanged "ComboFontItalic"

End Property

Public Property Get ButtonBackColor() As OLE_COLOR
    ButtonBackColor = mButtonBackColor

End Property

Public Property Let ButtonBackColor(ByVal vNewValue As OLE_COLOR)
    mButtonBackColor = vNewValue
    DrawControl , True
    PropertyChanged "ButtonBackColor"

End Property

Public Property Get ButtonOverColor() As OLE_COLOR
    ButtonOverColor = mButtonOverColor

End Property

Public Property Let ButtonOverColor(ByVal vNewValue As OLE_COLOR)
    mButtonOverColor = vNewValue
    PropertyChanged "ButtonOverColor"

End Property

Public Property Let ButtonForeColor(ByVal vNewValue As OLE_COLOR)
    mButtonForeColor = vNewValue
    DrawControl , True
    PropertyChanged "ButtonForeColor"

End Property

Public Property Get ButtonBorderStyle() As CfBdrStyle
    ButtonBorderStyle = mButtonBorderStyle

End Property

Public Property Let ButtonBorderStyle(ByVal vNewValue As CfBdrStyle)
    mButtonBorderStyle = vNewValue
    DrawControl , True
    PropertyChanged "ButtonBorderStyle"

End Property

Public Function AddToUsedList(FontName As String) As Integer

Dim i                                   As Integer
Dim F                                   As Boolean

    For i = 0 To mUsedCount - 1

        If LCase$(mUsedList(i)) = LCase$(FontName) Then
            F = True
            Exit For

        End If

    Next

    If F = False Then
        mUsedCount = mUsedCount + 1
        ReDim Preserve mUsedList(mUsedCount)
        mUsedList(mUsedCount - 1) = FontName
        AddToUsedList = mUsedCount - 1
    Else
        AddToUsedList = -1

    End If

End Function

Public Sub RemoveFromUsedList(FontName As String)

Dim i                                   As Integer
Dim tUL()                               As String
Dim fQ                                  As Integer

    ReDim tUL(mUsedCount)
    fQ = 1

    For i = 0 To mUsedCount - 1

        If LCase$(mUsedList(i)) <> LCase$(FontName) Then
            tUL(fQ - 1) = mUsedList(i)
            fQ = fQ + 1

        End If

    Next
    mUsedList = tUL
    mUsedCount = fQ
    ReDim Preserve mUsedList(mUsedCount)

End Sub

Public Property Get UsedCount() As Integer
    UsedCount = mUsedCount

End Property

Public Sub ClearUsedList()
    mUsedCount = 0
    ReDim mUsedList(0)

End Sub

Public Property Get UsedBackColor() As OLE_COLOR
    UsedBackColor = mUsedBackColor

End Property

Public Property Let UsedBackColor(ByVal vNewValue As OLE_COLOR)
    mUsedBackColor = vNewValue
    PropertyChanged "UsedBackColor"

End Property

Public Property Get UsedForeColor() As OLE_COLOR
    UsedForeColor = mUsedForeColor

End Property

Public Property Let UsedForeColor(ByVal vNewValue As OLE_COLOR)
    mUsedForeColor = vNewValue
    PropertyChanged "UsedForeColor"

End Property

Public Property Get XpStyle() As Boolean
    XpStyle = mXPStyle

End Property

Public Property Let XpStyle(ByVal vNewValue As Boolean)
    mXPStyle = vNewValue
    UserControl_Resize
    PropertyChanged "XPStyle"

End Property

Public Property Get ShowFocus() As Boolean
    ShowFocus = mShowFocus

End Property

Public Property Let ShowFocus(ByVal vNewValue As Boolean)
    mShowFocus = vNewValue
    PropertyChanged "ShowFocus"

End Property

Private Sub VScroll1_Change()

Dim tFont                               As String

    If doNothing = True Then
        Exit Sub

    End If

    fPos = VScroll1.Value
    DrawList
    tFont = fList(fPos - VScroll1.Value).fName
    ShowFont tFont

End Sub

Private Sub VScroll1_GotFocus()
    PicList.SetFocus

End Sub

Private Sub VScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
    UserControl_KeyDown KeyCode, Shift

End Sub
