VERSION 5.00
Begin VB.UserControl TextBoxW 
   BackColor       =   &H80000005&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   1  'vbSimpleBound
   ForeColor       =   &H80000008&
   HasDC           =   0   'False
   PropertyPages   =   "TextBoxW.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "TextBoxW.ctx":0046
End
Attribute VB_Name = "TextBoxW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If False Then
Private TxtCharacterCasingNormal, TxtCharacterCasingUpper, TxtCharacterCasingLower
Private TxtIconNone, TxtIconInfo, TxtIconWarning, TxtIconError
#End If
Public Enum TxtCharacterCasingConstants
TxtCharacterCasingNormal = 0
TxtCharacterCasingUpper = 1
TxtCharacterCasingLower = 2
End Enum
Private Const TTI_NONE As Long = 0
Private Const TTI_INFO As Long = 1
Private Const TTI_WARNING As Long = 2
Private Const TTI_ERROR As Long = 3
Public Enum TxtIconConstants
TxtIconNone = TTI_NONE
TxtIconInfo = TTI_INFO
TxtIconWarning = TTI_WARNING
TxtIconError = TTI_ERROR
End Enum
Private Type TagInitCommonControlsEx
dwSize As Long
dwICC As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type SIZEAPI
CX As Long
CY As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Const LF_FACESIZE As Long = 32
Private Const FW_NORMAL As Long = 400
Private Const FW_BOLD As Long = 700
Private Const DEFAULT_QUALITY As Long = 0
Private Type LOGFONT
LFHeight As Long
LFWidth As Long
LFEscapement As Long
LFOrientation As Long
LFWeight As Long
LFItalic As Byte
LFUnderline As Byte
LFStrikeOut As Byte
LFCharset As Byte
LFOutPrecision As Byte
LFClipPrecision As Byte
LFQuality As Byte
LFPitchAndFamily As Byte
LFFaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
End Type
Private Type EDITBALLOONTIP
cbStruct As Long
pszTitle As Long
pszText As Long
iIcon As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Public Event Scroll()
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."
Public Event ContextMenu(ByRef Handled As Boolean, ByVal X As Single, ByVal Y As Single)
Attribute ContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
Public Event PreviewKeyDown(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyDown.VB_Description = "Occurs before the KeyDown event."
Public Event PreviewKeyUp(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyUp.VB_Description = "Occurs before the KeyUp event."
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Public Event KeyPress(KeyChar As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an character key."
Attribute KeyPress.VB_UserMemId = -603
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Public Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TagInitCommonControlsEx) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Private Declare Function DragDetect Lib "user32" (ByVal hWnd As Long, ByVal PX As Integer, ByVal PY As Integer) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const RDW_UPDATENOW As Long = &H100
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4
Private Const RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const CF_UNICODETEXT As Long = 13
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_BORDER As Long = &H800000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_EX_CLIENTEDGE As Long = &H200
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_WINDOWEDGE As Long = &H100
Private Const WS_EX_RTLREADING As Long = &H2000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const SB_LINELEFT As Long = 0, SB_LINERIGHT As Long = 1
Private Const SB_LINEUP As Long = 0, SB_LINEDOWN As Long = 1
Private Const SB_THUMBPOSITION = 4, SB_THUMBTRACK As Long = 5
Private Const SB_HORZ As Long = 0, SB_VERT As Long = 1
Private Const WM_MOUSEACTIVATE As Long = &H21, MA_NOACTIVATE As Long = &H3, MA_NOACTIVATEANDEAT As Long = &H4, HTBORDER As Long = 18
Private Const SW_HIDE As Long = &H0
Private Const SW_SHOW As Long = &H5
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_COMMAND As Long = &H111
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_IME_CHAR As Long = &H286
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_HSCROLL As Long = &H114
Private Const WM_VSCROLL As Long = &H115
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const WM_COPY As Long = &H301
Private Const WM_CUT As Long = &H300
Private Const WM_PASTE As Long = &H302
Private Const WM_CLEAR As Long = &H303
Private Const EM_SETREADONLY As Long = &HCF, ES_READONLY As Long = &H800
Private Const EM_GETSEL As Long = &HB0
Private Const EM_SETSEL As Long = &HB1
Private Const EM_SCROLL As Long = &HB5
Private Const EM_LINESCROLL As Long = &HB6
Private Const EM_SCROLLCARET As Long = &HB7
Private Const EM_REPLACESEL As Long = &HC2
Private Const EM_GETPASSWORDCHAR As Long = &HD2
Private Const EM_SETPASSWORDCHAR As Long = &HCC
Private Const EM_GETLIMITTEXT As Long = &HD5
Private Const EM_LIMITTEXT As Long = &HC5
Private Const EM_SETLIMITTEXT As Long = EM_LIMITTEXT
Private Const EM_GETMODIFY As Long = &HB8
Private Const EM_SETMODIFY As Long = &HB9
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_GETTHUMB As Long = &HBE
Private Const EM_LINELENGTH As Long = &HC1
Private Const EM_GETLINE As Long = &HC4
Private Const EM_UNDO As Long = &HC7
Private Const EM_CANUNDO As Long = &HC6
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_EMPTYUNDOBUFFER As Long = &HCD
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_GETMARGINS As Long = &HD4
Private Const EM_SETMARGINS As Long = &HD3
Private Const EM_POSFROMCHAR As Long = &HD6
Private Const EM_CHARFROMPOS As Long = &HD7
Private Const ECM_FIRST As Long = &H1500
Private Const EM_SETCUEBANNER As Long = (ECM_FIRST + 1)
Private Const EM_GETCUEBANNER As Long = (ECM_FIRST + 2)
Private Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)
Private Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)
Private Const EN_UPDATE As Long = &H400
Private Const EN_CHANGE As Long = &H300
Private Const EN_HSCROLL As Long = &H601
Private Const EN_VSCROLL As Long = &H602
Private Const ES_AUTOHSCROLL As Long = &H80
Private Const ES_AUTOVSCROLL As Long = &H40
Private Const ES_NUMBER As Long = &H2000
Private Const ES_NOHIDESEL As Long = &H100
Private Const ES_LEFT As Long = &H0
Private Const ES_CENTER As Long = &H1
Private Const ES_RIGHT As Long = &H2
Private Const ES_MULTILINE As Long = &H4
Private Const ES_UPPERCASE As Long = &H8
Private Const ES_LOWERCASE As Long = &H10
Private Const ES_PASSWORD As Long = &H20
Private Const ES_WANTRETURN As Long = &H1000
Private Const EC_LEFTMARGIN As Long = &H1
Private Const EC_RIGHTMARGIN As Long = &H2
Private Const EC_USEFONTINFO As Long = &HFFFF&
Implements ISubclass
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IOleControlVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private TextBoxHandle As Long
Private TextBoxFontHandle As Long
Private TextBoxLogFont As LOGFONT
Private TextBoxAutoDragInSel As Boolean, TextBoxAutoDragIsActive As Boolean
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropOLEDragMode As VBRUN.OLEDragConstants
Private PropOLEDragDropScroll As Boolean
Private PropOLEDropMode As VBRUN.OLEDropConstants
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropBorderStyle As CCBorderStyleConstants
Private PropText As String
Private PropAlignment As VBRUN.AlignmentConstants
Private PropAllowOnlyNumbers As Boolean
Private PropLocked As Boolean
Private PropHideSelection As Boolean
Private PropPasswordChar As Integer
Private PropUseSystemPasswordChar As Boolean
Private PropMultiLine As Boolean
Private PropMaxLength As Long
Private PropScrollBars As VBRUN.ScrollBarConstants
Private PropCueBanner As String
Private PropCharacterCasing As TxtCharacterCasingConstants
Private PropWantReturn As Boolean

Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
    Dim KeyCode As Integer, IsInputKey As Boolean
    KeyCode = wParam And &HFF&
    If wMsg = WM_KEYDOWN Then
        RaiseEvent PreviewKeyDown(KeyCode, IsInputKey)
    ElseIf wMsg = WM_KEYUP Then
        RaiseEvent PreviewKeyUp(KeyCode, IsInputKey)
    End If
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd
            If TextBoxHandle <> 0 Then
                SendMessage TextBoxHandle, wMsg, wParam, ByVal lParam
                Handled = True
            End If
        Case vbKeyTab, vbKeyReturn, vbKeyEscape
            If IsInputKey = True Then
                If TextBoxHandle <> 0 Then
                    SendMessage TextBoxHandle, wMsg, wParam, ByVal lParam
                    Handled = True
                End If
            End If
    End Select
End If
End Sub

Private Sub IOleControlVB_GetControlInfo(ByRef Handled As Boolean, ByRef AccelCount As Integer, ByRef AccelTable As Long, ByRef Flags As Long)
If PropWantReturn = True And PropMultiLine = True Then
    Flags = CTRLINFO_EATS_RETURN
    Handled = True
End If
End Sub

Private Sub IOleControlVB_OnMnemonic(ByRef Handled As Boolean, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
End Sub

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispID As Long, ByRef DisplayName As String)
If DispID = DispIDMousePointer Then
    Select Case PropMousePointer
        Case 0: DisplayName = "0 - Default"
        Case 1: DisplayName = "1 - Arrow"
        Case 2: DisplayName = "2 - Cross"
        Case 3: DisplayName = "3 - I-Beam"
        Case 4: DisplayName = "4 - Hand"
        Case 5: DisplayName = "5 - Size"
        Case 6: DisplayName = "6 - Size NE SW"
        Case 7: DisplayName = "7 - Size N S"
        Case 8: DisplayName = "8 - Size NW SE"
        Case 9: DisplayName = "9 - Size W E"
        Case 10: DisplayName = "10 - Up Arrow"
        Case 11: DisplayName = "11 - Hourglass"
        Case 12: DisplayName = "12 - No Drop"
        Case 13: DisplayName = "13 - Arrow and Hourglass"
        Case 14: DisplayName = "14 - Arrow and Question"
        Case 15: DisplayName = "15 - Size All"
        Case 16: DisplayName = "16 - Arrow and CD"
        Case 99: DisplayName = "99 - Custom"
    End Select
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDMousePointer Then
    ReDim StringsOut(0 To (17 + 1)) As String
    ReDim CookiesOut(0 To (17 + 1)) As Long
    StringsOut(0) = "0 - Default": CookiesOut(0) = 0
    StringsOut(1) = "1 - Arrow": CookiesOut(1) = 1
    StringsOut(2) = "2 - Cross": CookiesOut(2) = 2
    StringsOut(3) = "3 - I-Beam": CookiesOut(3) = 3
    StringsOut(4) = "4 - Hand": CookiesOut(4) = 4
    StringsOut(5) = "5 - Size": CookiesOut(5) = 5
    StringsOut(6) = "6 - Size NE SW": CookiesOut(6) = 6
    StringsOut(7) = "7 - Size N S": CookiesOut(7) = 7
    StringsOut(8) = "8 - Size NW SE": CookiesOut(8) = 8
    StringsOut(9) = "9 - Size W E": CookiesOut(9) = 9
    StringsOut(10) = "10 - Up Arrow": CookiesOut(10) = 10
    StringsOut(11) = "11 - Hourglass": CookiesOut(11) = 11
    StringsOut(12) = "12 - No Drop": CookiesOut(12) = 12
    StringsOut(13) = "13 - Arrow and Hourglass": CookiesOut(13) = 13
    StringsOut(14) = "14 - Arrow and Question": CookiesOut(14) = 14
    StringsOut(15) = "15 - Size All": CookiesOut(15) = 15
    StringsOut(16) = "16 - Arrow and CD": CookiesOut(16) = 16
    StringsOut(17) = "99 - Custom": CookiesOut(17) = 99
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispID As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispID = DispIDMousePointer Then
    Value = Cookie
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Dim ICCEX As TagInitCommonControlsEx
With ICCEX
.dwSize = LenB(ICCEX)
.dwICC = ICC_STANDARD_CLASSES
End With
InitCommonControlsEx ICCEX
Call SetVTableSubclass(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableSubclass(Me, VTableInterfaceControl)
Call SetVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
DispIDMousePointer = GetDispID(Me, "MousePointer")
End Sub

Private Sub UserControl_InitProperties()
Set PropFont = Ambient.Font
PropVisualStyles = True
PropOLEDragMode = vbOLEDragManual
PropOLEDragDropScroll = True
PropOLEDropMode = vbOLEDropNone
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropBorderStyle = CCBorderStyleSunken
PropText = Ambient.DisplayName
PropAlignment = vbLeftJustify
PropAllowOnlyNumbers = False
PropLocked = False
PropHideSelection = True
PropPasswordChar = 0
PropUseSystemPasswordChar = False
PropMultiLine = False
PropMaxLength = 0
PropScrollBars = vbSBNone
PropCueBanner = vbNullString
PropCharacterCasing = TxtCharacterCasingNormal
PropWantReturn = False
Call CreateTextBox
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
Set PropFont = .ReadProperty("Font", Ambient.Font)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.BackColor = .ReadProperty("BackColor", vbWindowBackground)
Me.ForeColor = .ReadProperty("ForeColor", vbWindowText)
Me.Enabled = .ReadProperty("Enabled", True)
PropOLEDragMode = .ReadProperty("OLEDragMode", vbOLEDragManual)
PropOLEDragDropScroll = .ReadProperty("OLEDragDropScroll", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropBorderStyle = .ReadProperty("BorderStyle", CCBorderStyleSunken)
PropText = VarToStr(.ReadProperty("Text", vbNullString))
PropAlignment = .ReadProperty("Alignment", vbLeftJustify)
PropAllowOnlyNumbers = .ReadProperty("AllowOnlyNumbers", False)
PropLocked = .ReadProperty("Locked", False)
PropHideSelection = .ReadProperty("HideSelection", True)
PropPasswordChar = .ReadProperty("PasswordChar", 0)
PropUseSystemPasswordChar = .ReadProperty("UseSystemPasswordChar", False)
PropMultiLine = .ReadProperty("MultiLine", False)
PropMaxLength = .ReadProperty("MaxLength", 0)
PropScrollBars = .ReadProperty("ScrollBars", vbSBNone)
PropCueBanner = VarToStr(.ReadProperty("CueBanner", vbNullString))
PropCharacterCasing = .ReadProperty("CharacterCasing", TxtCharacterCasingNormal)
PropWantReturn = .ReadProperty("WantReturn", False)
End With
Call CreateTextBox
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", PropFont, Ambient.Font
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", Me.BackColor, vbWindowBackground
.WriteProperty "ForeColor", Me.ForeColor, vbWindowText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDragMode", PropOLEDragMode, vbOLEDragManual
.WriteProperty "OLEDragDropScroll", PropOLEDragDropScroll, True
.WriteProperty "OLEDropMode", PropOLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "BorderStyle", PropBorderStyle, CCBorderStyleSunken
.WriteProperty "Text", StrToVar(PropText), vbNullString
.WriteProperty "Alignment", PropAlignment, vbLeftJustify
.WriteProperty "AllowOnlyNumbers", PropAllowOnlyNumbers, False
.WriteProperty "Locked", PropLocked, False
.WriteProperty "HideSelection", PropHideSelection, True
.WriteProperty "PasswordChar", PropPasswordChar, 0
.WriteProperty "UseSystemPasswordChar", PropUseSystemPasswordChar, False
.WriteProperty "MultiLine", PropMultiLine, False
.WriteProperty "MaxLength", PropMaxLength, 0
.WriteProperty "ScrollBars", PropScrollBars, vbSBNone
.WriteProperty "CueBanner", StrToVar(PropCueBanner), vbNullString
.WriteProperty "CharacterCasing", PropCharacterCasing, TxtCharacterCasingNormal
.WriteProperty "WantReturn", PropWantReturn, False
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
If PropOLEDragMode = vbOLEDragAutomatic And TextBoxAutoDragIsActive = True And Effect = vbDropEffectMove Then
    If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_CLEAR, 0, ByVal 0&
End If
RaiseEvent OLECompleteDrag(Effect)
TextBoxAutoDragIsActive = False
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
If PropOLEDropMode = vbOLEDropAutomatic And TextBoxHandle <> 0 Then
    If Not Effect = vbDropEffectNone Then
        Me.Refresh
        Dim Text As String
        If Data.GetFormat(CF_UNICODETEXT) = True Then
            Text = Data.GetData(CF_UNICODETEXT)
            Text = Left$(Text, InStr(Text, vbNullChar) - 1)
        ElseIf Data.GetFormat(vbCFText) = True Then
            Text = Data.GetData(vbCFText)
        End If
        If Not Text = vbNullString Then
            Dim CharPos As Long
            CharPos = CIntToUInt(LoWord(SendMessage(TextBoxHandle, EM_CHARFROMPOS, 0, ByVal MakeDWord(X, Y))))
            If TextBoxAutoDragIsActive = True Then
                TextBoxAutoDragIsActive = False
                Dim SelStart As Long, SelEnd As Long
                SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
                If CharPos >= SelStart And CharPos <= SelEnd Then
                    Effect = vbDropEffectNone
                    Exit Sub
                End If
                If SelStart < CharPos Then CharPos = CharPos - (SelEnd - SelStart)
                If Effect = vbDropEffectMove Then SendMessage TextBoxHandle, WM_CLEAR, 0, ByVal 0&
            Else
                If GetFocus() <> TextBoxHandle Then SetFocusAPI UserControl.hWnd
            End If
            SendMessage TextBoxHandle, EM_SETSEL, CharPos, ByVal CharPos
            SendMessage TextBoxHandle, EM_REPLACESEL, 0, ByVal StrPtr(Text)
            SendMessage TextBoxHandle, EM_SETSEL, CharPos, ByVal (CharPos + Len(Text))
        End If
    End If
End If
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
If TextBoxHandle <> 0 Then
    If State = vbOver And Not Effect = vbDropEffectNone Then
        If PropOLEDragDropScroll = True Then
            Dim RC As RECT
            GetWindowRect TextBoxHandle, RC
            Dim dwStyle As Long
            dwStyle = GetWindowLong(TextBoxHandle, GWL_STYLE)
            If (dwStyle And WS_HSCROLL) = WS_HSCROLL Then
                If Abs(X) < 16 Then
                    SendMessage TextBoxHandle, WM_HSCROLL, SB_LINELEFT, ByVal 0&
                ElseIf Abs(X - (RC.Right - RC.Left)) < 16 Then
                    SendMessage TextBoxHandle, WM_HSCROLL, SB_LINERIGHT, ByVal 0&
                End If
            End If
            If (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
                If Abs(Y) < 16 Then
                    SendMessage TextBoxHandle, WM_VSCROLL, SB_LINEUP, ByVal 0&
                ElseIf Abs(Y - (RC.Bottom - RC.Top)) < 16 Then
                    SendMessage TextBoxHandle, WM_VSCROLL, SB_LINEDOWN, ByVal 0&
                End If
            End If
        End If
    End If
    If PropOLEDropMode = vbOLEDropAutomatic Then
        If State = vbOver And Not Effect = vbDropEffectNone Then
            Dim CharPos As Long, CaretPos As Long
            CharPos = CIntToUInt(LoWord(SendMessage(TextBoxHandle, EM_CHARFROMPOS, 0, ByVal MakeDWord(X, Y))))
            CaretPos = SendMessage(TextBoxHandle, EM_POSFROMCHAR, CharPos, ByVal 0&)
            If CaretPos > -1 Then
                Dim hDC As Long, Size As SIZEAPI
                hDC = GetDC(TextBoxHandle)
                SelectObject hDC, TextBoxFontHandle
                GetTextExtentPoint32 hDC, StrPtr("|"), 1, Size
                ReleaseDC TextBoxHandle, hDC
                CreateCaret TextBoxHandle, 0, 0, Size.CY
                SetCaretPos LoWord(CaretPos), HiWord(CaretPos)
                ShowCaret TextBoxHandle
            Else
                If GetFocus() <> TextBoxHandle Then
                    DestroyCaret
                Else
                    Me.Refresh
                End If
            End If
        ElseIf State = vbLeave Then
            If GetFocus() <> TextBoxHandle Then
                DestroyCaret
            Else
                Me.Refresh
            End If
        End If
    End If
End If
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
If PropOLEDragMode = vbOLEDragAutomatic Then
    Dim Text As String
    Text = Me.SelText
    Data.SetData StrToVar(Text & vbNullChar), CF_UNICODETEXT
    Data.SetData StrToVar(Text), vbCFText
    AllowedEffects = vbDropEffectMove
    TextBoxAutoDragIsActive = True
End If
RaiseEvent OLEStartDrag(Data, AllowedEffects)
If AllowedEffects = vbDropEffectNone Then TextBoxAutoDragIsActive = False
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
UserControl.OLEDrag
End Sub

Private Sub UserControl_Resize()
If TextBoxHandle = 0 Then Exit Sub
With UserControl
MoveWindow TextBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableSubclass(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableSubclass(Me, VTableInterfaceControl)
Call RemoveVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyTextBox
Call ComCtlsReleaseShellMod
End Sub

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name used in code to identify an object."
Name = Ambient.DisplayName
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal Value As String)
Extender.Tag = Value
End Property

Public Property Get Parent() As Object
Attribute Parent.VB_Description = "Returns the object on which this object is located."
Set Parent = UserControl.Parent
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of an object."
Width = Extender.Width
End Property

Public Property Let Width(ByVal Value As Single)
Extender.Width = Value
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns/sets the height of an object."
Height = Extender.Height
End Property

Public Property Let Height(ByVal Value As Single)
Extender.Height = Value
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
hWnd = TextBoxHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
Set Font = PropFont
End Property

Public Property Let Font(ByVal NewFont As StdFont)
Set Me.Font = NewFont
End Property

Public Property Set Font(ByVal NewFont As StdFont)
Dim OldFontHandle As Long
Set PropFont = NewFont
Call OLEFontToLogFont(NewFont, TextBoxLogFont)
OldFontHandle = TextBoxFontHandle
TextBoxFontHandle = CreateFontIndirect(TextBoxLogFont)
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_SETFONT, TextBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
Call OLEFontToLogFont(PropFont, TextBoxLogFont)
OldFontHandle = TextBoxFontHandle
TextBoxFontHandle = CreateFontIndirect(TextBoxLogFont)
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_SETFONT, TextBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Private Sub OLEFontToLogFont(ByVal Font As StdFont, ByRef LF As LOGFONT)
Dim FontName As String
With LF
FontName = Left$(Font.Name, LF_FACESIZE)
CopyMemory .LFFaceName(0), ByVal StrPtr(FontName), LenB(FontName)
.LFHeight = -MulDiv(CLng(Font.Size), DPI_Y(), 72)
If Font.Bold = True Then .LFWeight = FW_BOLD Else .LFWeight = FW_NORMAL
.LFItalic = IIf(Font.Italic = True, 1, 0)
.LFStrikeOut = IIf(Font.Strikethrough = True, 1, 0)
.LFUnderline = IIf(Font.Underline = True, 1, 0)
.LFQuality = DEFAULT_QUALITY
.LFCharset = CByte(Font.Charset And &HFF)
End With
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If TextBoxHandle <> 0 And EnabledVisualStyles() = True Then
    Select Case PropVisualStyles
        Case True
            ActivateVisualStyles TextBoxHandle
        Case False
            RemoveVisualStyles TextBoxHandle
    End Select
    Me.Refresh
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
UserControl.BackColor = Value
Me.Refresh
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
UserControl.ForeColor = Value
Me.Refresh
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If TextBoxHandle <> 0 Then
    EnableWindow TextBoxHandle, IIf(Value = True, 1, 0)
    Me.Refresh
End If
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDragMode() As VBRUN.OLEDragConstants
OLEDragMode = PropOLEDragMode
End Property

Public Property Let OLEDragMode(ByVal Value As VBRUN.OLEDragConstants)
Select Case Value
    Case vbOLEDragManual, vbOLEDragAutomatic
        PropOLEDragMode = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "OLEDragMode"
End Property

Public Property Get OLEDragDropScroll() As Boolean
Attribute OLEDragDropScroll.VB_Description = "Returns/Sets whether this object will scroll during an OLE drag/drop operation."
OLEDragDropScroll = PropOLEDragDropScroll
End Property

Public Property Let OLEDragDropScroll(ByVal Value As Boolean)
PropOLEDragDropScroll = Value
UserControl.PropertyChanged "OLEDragDropScroll"
End Property

Public Property Get OLEDropMode() As VBRUN.OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
OLEDropMode = PropOLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As VBRUN.OLEDropConstants)
Select Case Value
    Case vbOLEDropNone, vbOLEDropManual, vbOLEDropAutomatic
        PropOLEDropMode = Value
        UserControl.OLEDropMode = IIf(PropOLEDropMode = vbOLEDropAutomatic, vbOLEDropManual, Value)
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "OLEDropMode"
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
MousePointer = PropMousePointer
End Property

Public Property Let MousePointer(ByVal Value As Integer)
Select Case Value
    Case 0 To 16, 99
        PropMousePointer = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Returns/sets a custom mouse icon."
Set MouseIcon = PropMouseIcon
End Property

Public Property Let MouseIcon(ByVal Value As IPictureDisp)
Set Me.MouseIcon = Value
End Property

Public Property Set MouseIcon(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropMouseIcon = Nothing
Else
    If Value.Type = vbPicTypeIcon Or Value.Handle = 0 Then
        Set PropMouseIcon = Value
    Else
        If Ambient.UserMode = False Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
UserControl.PropertyChanged "MouseIcon"
End Property

Public Property Get BorderStyle() As CCBorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style."
Attribute BorderStyle.VB_UserMemId = -504
BorderStyle = PropBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As CCBorderStyleConstants)
Select Case Value
    Case CCBorderStyleNone, CCBorderStyleSingle, CCBorderStyleThin, CCBorderStyleSunken, CCBorderStyleRaised
        PropBorderStyle = Value
    Case Else
        Err.Raise 380
End Select
If TextBoxHandle <> 0 Then Call ComCtlsChangeBorderStyle(TextBoxHandle, PropBorderStyle)
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
Attribute Text.VB_ProcData.VB_Invoke_Property = "PPTextBoxWText"
Attribute Text.VB_UserMemId = -517
Attribute Text.VB_MemberFlags = "123c"
If TextBoxHandle <> 0 Then
    Text = String(SendMessage(TextBoxHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
    SendMessage TextBoxHandle, WM_GETTEXT, Len(Text) + 1, ByVal StrPtr(Text)
Else
    Text = PropText
End If
End Property

Public Property Let Text(ByVal Value As String)
PropText = Value
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_SETTEXT, 0, ByVal StrPtr(PropText)
UserControl.PropertyChanged "Text"
End Property

Public Property Get Default() As String
Attribute Default.VB_UserMemId = 0
Attribute Default.VB_MemberFlags = "40"
Default = Me.Text
End Property

Public Property Let Default(ByVal Value As String)
Me.Text = Value
End Property

Public Property Get Alignment() As VBRUN.AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment."
Alignment = PropAlignment
End Property

Public Property Let Alignment(ByVal Value As VBRUN.AlignmentConstants)
Select Case Value
    Case vbLeftJustify, vbCenter, vbRightJustify
        PropAlignment = Value
    Case Else
        Err.Raise 380
End Select
If TextBoxHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TextBoxHandle, GWL_STYLE)
    If (dwStyle And ES_LEFT) = ES_LEFT Then dwStyle = dwStyle And Not ES_LEFT
    If (dwStyle And ES_CENTER) = ES_CENTER Then dwStyle = dwStyle And Not ES_CENTER
    If (dwStyle And ES_RIGHT) = ES_RIGHT Then dwStyle = dwStyle And Not ES_RIGHT
    Select Case PropAlignment
        Case vbLeftJustify
            dwStyle = dwStyle Or ES_LEFT
        Case vbCenter
            dwStyle = dwStyle Or ES_CENTER
        Case vbRightJustify
            dwStyle = dwStyle Or ES_RIGHT
    End Select
    SetWindowLong TextBoxHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "Alignment"
End Property

Public Property Get AllowOnlyNumbers() As Boolean
Attribute AllowOnlyNumbers.VB_Description = "Returns/sets a value indicating if only numbers are allowed to be entered."
AllowOnlyNumbers = PropAllowOnlyNumbers
End Property

Public Property Let AllowOnlyNumbers(ByVal Value As Boolean)
PropAllowOnlyNumbers = Value
If TextBoxHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TextBoxHandle, GWL_STYLE)
    If PropAllowOnlyNumbers = True Then
        If Not (dwStyle And ES_NUMBER) = ES_NUMBER Then dwStyle = dwStyle Or ES_NUMBER
    Else
        If (dwStyle And ES_NUMBER) = ES_NUMBER Then dwStyle = dwStyle And Not ES_NUMBER
    End If
    SetWindowLong TextBoxHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "AllowOnlyNumbers"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents can be edited."
If TextBoxHandle <> 0 Then
    Locked = CBool((GetWindowLong(TextBoxHandle, GWL_STYLE) And ES_READONLY) <> 0)
Else
    Locked = PropLocked
End If
End Property

Public Property Let Locked(ByVal Value As Boolean)
PropLocked = Value
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETREADONLY, IIf(PropLocked = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "Locked"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value indicating if the selection in an edit control is hidden when the control loses focus."
HideSelection = PropHideSelection
End Property

Public Property Let HideSelection(ByVal Value As Boolean)
PropHideSelection = Value
If TextBoxHandle <> 0 Then Call ReCreateTextBox
UserControl.PropertyChanged "HideSelection"
End Property

Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
If TextBoxHandle <> 0 Then
    PasswordChar = ChrW(SendMessage(TextBoxHandle, EM_GETPASSWORDCHAR, 0, ByVal 0&))
Else
    PasswordChar = ChrW(PropPasswordChar)
End If
End Property

Public Property Let PasswordChar(ByVal Value As String)
If PropUseSystemPasswordChar = True Then Exit Property
If Value = vbNullString Or Len(Value) = 0 Then
    PropPasswordChar = 0
ElseIf Len(Value) = 1 Then
    PropPasswordChar = AscW(Value)
Else
    If Ambient.UserMode = False Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If TextBoxHandle <> 0 Then
    SendMessage TextBoxHandle, EM_SETPASSWORDCHAR, PropPasswordChar, ByVal 0&
    Me.Refresh
End If
UserControl.PropertyChanged "PasswordChar"
End Property

Public Property Get UseSystemPasswordChar() As Boolean
Attribute UseSystemPasswordChar.VB_Description = "Returns/sets a value indicating if the default system password character is used. This property has precedence over the password char property."
UseSystemPasswordChar = PropUseSystemPasswordChar
End Property

Public Property Let UseSystemPasswordChar(ByVal Value As Boolean)
PropUseSystemPasswordChar = Value
If TextBoxHandle <> 0 Then Call ReCreateTextBox
UserControl.PropertyChanged "UseSystemPasswordChar"
End Property

Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
MultiLine = PropMultiLine
End Property

Public Property Let MultiLine(ByVal Value As Boolean)
If Ambient.UserMode = True Then
    Err.Raise Number:=382, Description:="MultiLine property is read-only at run time"
Else
    PropMultiLine = Value
    If TextBoxHandle <> 0 Then Call ReCreateTextBox
End If
UserControl.PropertyChanged "MultiLine"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
MaxLength = PropMaxLength
End Property

Public Property Let MaxLength(ByVal Value As Long)
If Value < 0 Then
    If Ambient.UserMode = False Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropMaxLength = Value
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETLIMITTEXT, PropMaxLength, ByVal 0&
UserControl.PropertyChanged "MaxLength"
End Property

Public Property Get ScrollBars() As VBRUN.ScrollBarConstants
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
ScrollBars = PropScrollBars
End Property

Public Property Let ScrollBars(ByVal Value As VBRUN.ScrollBarConstants)
Select Case Value
    Case vbSBNone, vbHorizontal, vbVertical, vbBoth
        PropScrollBars = Value
        If TextBoxHandle <> 0 Then Call ReCreateTextBox
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "ScrollBars"
End Property

Public Property Get CueBanner() As String
Attribute CueBanner.VB_Description = "Returns/sets the textual cue, or tip, that is displayed to prompt the user for information. Only applicable if the multi line property is set to false. Requires comctl32.dll version 6.0 or higher."
CueBanner = PropCueBanner
End Property

Public Property Let CueBanner(ByVal Value As String)
PropCueBanner = Value
If TextBoxHandle <> 0 And PropMultiLine = False And ComCtlsSupportLevel() >= 1 Then SendMessage TextBoxHandle, EM_SETCUEBANNER, 0, ByVal StrPtr(PropCueBanner)
UserControl.PropertyChanged "CueBanner"
End Property

Public Property Get CharacterCasing() As TxtCharacterCasingConstants
Attribute CharacterCasing.VB_Description = "Returns/sets a value indicating if the text box modifies the case of characters as they are typed."
CharacterCasing = PropCharacterCasing
End Property

Public Property Let CharacterCasing(ByVal Value As TxtCharacterCasingConstants)
Select Case Value
    Case TxtCharacterCasingNormal, TxtCharacterCasingUpper, TxtCharacterCasingLower
        PropCharacterCasing = Value
    Case Else
        Err.Raise 380
End Select
If TextBoxHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TextBoxHandle, GWL_STYLE)
    If (dwStyle And ES_UPPERCASE) = ES_UPPERCASE Then dwStyle = dwStyle And Not ES_UPPERCASE
    If (dwStyle And ES_LOWERCASE) = ES_LOWERCASE Then dwStyle = dwStyle And Not ES_LOWERCASE
    Select Case PropCharacterCasing
        Case TxtCharacterCasingUpper
            dwStyle = dwStyle Or ES_UPPERCASE
        Case TxtCharacterCasingLower
            dwStyle = dwStyle Or ES_LOWERCASE
    End Select
    SetWindowLong TextBoxHandle, GWL_STYLE, dwStyle
    If Ambient.UserMode = False Then Me.Text = PropText
End If
UserControl.PropertyChanged "CharacterCasing"
End Property

Public Property Get WantReturn() As Boolean
Attribute WantReturn.VB_Description = "Returns/sets a value that determines when the user presses RETURN to perform the default button or to advance to the next line. This property applies only to a multiline text box and when there is any default button on the form."
WantReturn = PropWantReturn
End Property

Public Property Let WantReturn(ByVal Value As Boolean)
If PropWantReturn = Value Then Exit Property
PropWantReturn = Value
If TextBoxHandle <> 0 And Ambient.UserMode = True Then
    ' It is not possible (in VB6) to achieve this when specifying ES_WANTRETURN.
    Dim PropOleObject As OLEGuids.IOleObject
    Dim PropClientSite As OLEGuids.IOleClientSite
    Dim PropUnknown As IUnknown
    Dim PropControlSite As OLEGuids.IOleControlSite
    On Error Resume Next
    Set PropOleObject = Me
    Set PropClientSite = PropOleObject.GetClientSite
    Set PropUnknown = PropClientSite
    Set PropControlSite = PropUnknown
    PropControlSite.OnControlInfoChanged
    If GetFocus() = TextBoxHandle Then
        ' If focus is on the control then force the change immediately.
        PropControlSite.OnFocus 1
    End If
    On Error GoTo 0
End If
UserControl.PropertyChanged "WantReturn"
End Property

Private Sub CreateTextBox()
If TextBoxHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
Select Case PropBorderStyle
    Case CCBorderStyleSingle
        dwStyle = dwStyle Or WS_BORDER
    Case CCBorderStyleThin
        dwExStyle = dwExStyle Or WS_EX_STATICEDGE
    Case CCBorderStyleSunken
        dwExStyle = dwExStyle Or WS_EX_CLIENTEDGE
    Case CCBorderStyleRaised
        dwExStyle = dwExStyle Or WS_EX_WINDOWEDGE
        dwStyle = dwStyle Or WS_DLGFRAME
End Select
If PropAllowOnlyNumbers = True Then dwStyle = dwStyle Or ES_NUMBER
Select Case PropAlignment
    Case vbLeftJustify
        dwStyle = dwStyle Or ES_LEFT
    Case vbCenter
        dwStyle = dwStyle Or ES_CENTER
    Case vbRightJustify
        dwStyle = dwStyle Or ES_RIGHT
End Select
If PropLocked = True Then dwStyle = dwStyle Or ES_READONLY
If PropHideSelection = False Then dwStyle = dwStyle Or ES_NOHIDESEL
If PropUseSystemPasswordChar = True Then dwStyle = dwStyle Or ES_PASSWORD
If PropMultiLine = True Then
    dwStyle = dwStyle Or ES_MULTILINE
    Select Case PropScrollBars
        Case vbSBNone
            dwStyle = dwStyle Or ES_AUTOVSCROLL
        Case vbHorizontal
            dwStyle = dwStyle Or WS_HSCROLL Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL
        Case vbVertical
            dwStyle = dwStyle Or WS_VSCROLL Or ES_AUTOVSCROLL
        Case vbBoth
            dwStyle = dwStyle Or WS_HSCROLL Or WS_VSCROLL Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL
    End Select
Else
    dwStyle = dwStyle Or ES_AUTOHSCROLL
End If
Select Case PropCharacterCasing
    Case TxtCharacterCasingUpper
        dwStyle = dwStyle Or ES_UPPERCASE
    Case TxtCharacterCasingLower
        dwStyle = dwStyle Or ES_LOWERCASE
End Select
If Ambient.RightToLeft = True Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
TextBoxHandle = CreateWindowEx(dwExStyle, StrPtr("Edit"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If TextBoxHandle <> 0 Then
    If PropPasswordChar <> 0 And PropUseSystemPasswordChar = False Then SendMessage TextBoxHandle, EM_SETPASSWORDCHAR, PropPasswordChar, ByVal 0&
    SendMessage TextBoxHandle, EM_SETLIMITTEXT, PropMaxLength, ByVal 0&
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.Text = PropText
If Not PropCueBanner = vbNullString Then Me.CueBanner = PropCueBanner
If Ambient.UserMode = True Then
    If TextBoxHandle <> 0 Then Call ComCtlsSetSubclass(TextBoxHandle, Me, 1)
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
End If
End Sub

Private Sub ReCreateTextBox()
If Ambient.UserMode = True Then
    Dim Visible As Boolean
    Visible = Extender.Visible
    With Me
    If Visible = True Then SendMessage UserControl.hWnd, WM_SETREDRAW, 0, ByVal 0&
    Dim SelStart As Long, SelEnd As Long
    Dim ScrollPosHorz As Integer, ScrollPosVert As Integer
    If TextBoxHandle <> 0 Then
        SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
        If PropMultiLine = True And PropScrollBars <> vbSBNone Then
            If PropScrollBars = vbHorizontal Or PropScrollBars = vbBoth Then
                ScrollPosHorz = CUIntToInt(GetScrollPos(TextBoxHandle, SB_HORZ) And &HFFFF&)
            End If
            If PropScrollBars = vbVertical Or PropScrollBars = vbBoth Then
                ScrollPosVert = CUIntToInt(GetScrollPos(TextBoxHandle, SB_VERT) And &HFFFF&)
            End If
        End If
        Dim Buffer As String
        Buffer = String(SendMessage(TextBoxHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
        SendMessage TextBoxHandle, WM_GETTEXT, Len(Buffer) + 1, ByVal StrPtr(Buffer)
        PropText = Buffer
    End If
    Call DestroyTextBox
    Call CreateTextBox
    Call UserControl_Resize
    If TextBoxHandle <> 0 Then
        SendMessage TextBoxHandle, EM_SETSEL, SelStart, ByVal SelEnd
        If ScrollPosHorz > 0 Then SendMessage TextBoxHandle, WM_HSCROLL, MakeDWord(SB_THUMBPOSITION, ScrollPosHorz), ByVal 0&
        If ScrollPosVert > 0 Then SendMessage TextBoxHandle, WM_VSCROLL, MakeDWord(SB_THUMBPOSITION, ScrollPosVert), ByVal 0&
    End If
    If Visible = True Then SendMessage UserControl.hWnd, WM_SETREDRAW, 1, ByVal 0&
    .Refresh
    End With
Else
    Call DestroyTextBox
    Call CreateTextBox
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyTextBox()
If TextBoxHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(TextBoxHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow TextBoxHandle, SW_HIDE
SetParent TextBoxHandle, 0
DestroyWindow TextBoxHandle
TextBoxHandle = 0
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Sub Copy()
Attribute Copy.VB_Description = "Method to copy the current selection to the clipboard."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_COPY, 0, ByVal 0&
End Sub

Public Sub Cut()
Attribute Cut.VB_Description = "Method to delete (cut) the current selection and copy the deleted text to the clipboard."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_CUT, 0, ByVal 0&
End Sub

Public Sub Paste()
Attribute Paste.VB_Description = "Method to copy the current content of the clipboard at the current caret position."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_PASTE, 0, ByVal 0&
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Method to delete (clear) the current selection."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_CLEAR, 0, ByVal 0&
End Sub

Public Sub Undo()
Attribute Undo.VB_Description = "Undoes the last operation, if any."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_UNDO, 0, ByVal 0&
End Sub

Public Function CanUndo() As Boolean
Attribute CanUndo.VB_Description = "Determines whether there are any actions in the undo queue."
If TextBoxHandle <> 0 Then CanUndo = CBool(SendMessage(TextBoxHandle, EM_CANUNDO, 0, ByVal 0&) <> 0)
End Function

Public Sub ResetUndoFlag()
Attribute ResetUndoFlag.VB_Description = "Resets the undo flag."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_EMPTYUNDOBUFFER, 0, ByVal 0&
End Sub

Public Property Get Modified() As Boolean
Attribute Modified.VB_Description = "Setting the text property will reset this property to false. Any typing in the control will set the property to true."
Attribute Modified.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then Modified = CBool(SendMessage(TextBoxHandle, EM_GETMODIFY, 0, ByVal 0&) <> 0)
End Property

Public Property Let Modified(ByVal Value As Boolean)
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETMODIFY, IIf(Value = True, 1, 0), ByVal 0&
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected; indicates the position of the insertion point if no text is selected."
Attribute SelStart.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal 0&
End Property

Public Property Let SelStart(ByVal Value As Long)
If TextBoxHandle <> 0 Then
    If Value >= 0 Then
        SendMessage TextBoxHandle, EM_SETSEL, Value, ByVal Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then
    Dim SelStart As Long, SelEnd As Long
    SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    SelLength = SelEnd - SelStart
End If
End Property

Public Property Let SelLength(ByVal Value As Long)
If TextBoxHandle <> 0 Then
    If Value >= 0 Then
        Dim SelStart As Long
        SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal 0&
        SendMessage TextBoxHandle, EM_SETSEL, SelStart, ByVal SelStart + Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then
    Dim SelStart As Long, SelEnd As Long
    SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    On Error Resume Next
    SelText = Mid$(Me.Text, SelStart + 1, (SelEnd - SelStart))
    On Error GoTo 0
End If
End Property

Public Property Let SelText(ByVal Value As String)
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_REPLACESEL, 0, ByVal StrPtr(Value)
End Property

Public Function GetLine(ByVal LineNumber As Long) As String
Attribute GetLine.VB_Description = "Retrieves the text of the specified line. A value of 0 indicates that the text of the current line number (the line that contains the caret) will be retrieved."
If LineNumber < 0 Then Err.Raise 380
If TextBoxHandle <> 0 Then
    Dim FirstCharPos As Long, Length As Long
    FirstCharPos = SendMessage(TextBoxHandle, EM_LINEINDEX, LineNumber - 1, ByVal 0&)
    If FirstCharPos > -1 Then
        Length = SendMessage(TextBoxHandle, EM_LINELENGTH, FirstCharPos, ByVal 0&)
        If Length > 0 Then
            Dim Buffer As String
            Buffer = ChrW(Length) & String(Length - 1, vbNullChar)
            If LineNumber > 0 Then
                If SendMessage(TextBoxHandle, EM_GETLINE, LineNumber - 1, ByVal StrPtr(Buffer)) > 0 Then GetLine = Buffer
            Else
                If SendMessage(TextBoxHandle, EM_GETLINE, SendMessage(TextBoxHandle, EM_LINEFROMCHAR, FirstCharPos, ByVal 0&), ByVal StrPtr(Buffer)) > 0 Then GetLine = Buffer
            End If
        End If
    Else
        Err.Raise 380
    End If
End If
End Function

Public Function GetLineCount() As Long
Attribute GetLineCount.VB_Description = "Gets the number of lines."
If TextBoxHandle <> 0 Then GetLineCount = SendMessage(TextBoxHandle, EM_GETLINECOUNT, 0, ByVal 0&)
End Function

Public Sub ScrollToLine(ByVal LineNumber As Long)
Attribute ScrollToLine.VB_Description = "Scrolls to ensure the specified line is visible."
If LineNumber < 0 Then Err.Raise 380
If TextBoxHandle <> 0 Then
    If SendMessage(TextBoxHandle, EM_LINESCROLL, 0, ByVal LineNumber - 1) <> 0 Then
        Dim FirstCharPos As Long
        FirstCharPos = SendMessage(TextBoxHandle, EM_LINEINDEX, LineNumber - 1, ByVal 0&)
        If FirstCharPos > -1 Then
            Me.SelStart = FirstCharPos
            SendMessage TextBoxHandle, EM_SCROLLCARET, 0, ByVal 0&
        Else
            Err.Raise 380
        End If
    End If
End If
End Sub

Public Sub ScrollToCaret()
Attribute ScrollToCaret.VB_Description = "Scrolls the caret into view."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SCROLLCARET, 0, ByVal 0&
End Sub

Public Function CharFromPos(ByVal X As Single, ByVal Y As Single) As Long
Attribute CharFromPos.VB_Description = "Returns the character index closest to a specified point."
Dim P As POINTAPI
P.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
P.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
If TextBoxHandle <> 0 Then CharFromPos = CIntToUInt(LoWord(SendMessage(TextBoxHandle, EM_CHARFROMPOS, 0, ByVal MakeDWord(P.X, P.Y))))
End Function

Public Function GetLineFromChar(ByVal CharIndex As Long) As Long
Attribute GetLineFromChar.VB_Description = "Gets the line number that contains the specified character index. A character index of -1 retrieves either the current line or the beginning of the current selection."
If CharIndex < -1 Then Err.Raise 380
If TextBoxHandle <> 0 Then GetLineFromChar = SendMessage(TextBoxHandle, EM_LINEFROMCHAR, CharIndex, ByVal 0&) + 1
End Function

Public Function ShowBalloonTip(ByVal Text As String, Optional ByVal Title As String, Optional ByVal Icon As TxtIconConstants) As Boolean
Attribute ShowBalloonTip.VB_Description = "Displays a balloon tip. Requires comctl32.dll version 6.0 or higher."
If TextBoxHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim EDITBT As EDITBALLOONTIP
    With EDITBT
    .cbStruct = LenB(EDITBT)
    .pszText = StrPtr(Text)
    .pszTitle = StrPtr(Title)
    Select Case Icon
        Case TxtIconNone, TxtIconInfo, TxtIconWarning, TxtIconError
            .iIcon = Icon
        Case Else
            Err.Raise 380
    End Select
    If GetFocus() <> TextBoxHandle Then SetFocusAPI UserControl.hWnd
    ShowBalloonTip = CBool(SendMessage(TextBoxHandle, EM_SHOWBALLOONTIP, 0, ByVal VarPtr(EDITBT)) <> 0)
    End With
End If
End Function

Public Function HideBalloonTip() As Boolean
Attribute HideBalloonTip.VB_Description = "Hides any associated balloon tip. Requires comctl32.dll version 6.0 or higher."
If TextBoxHandle <> 0 And ComCtlsSupportLevel() >= 1 Then HideBalloonTip = CBool(SendMessage(TextBoxHandle, EM_HIDEBALLOONTIP, 0, ByVal 0&) <> 0)
End Function

Public Property Get LeftMargin() As Single
Attribute LeftMargin.VB_Description = "Returns/sets the widths of the left margin."
Attribute LeftMargin.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then LeftMargin = UserControl.ScaleX(LoWord(SendMessage(TextBoxHandle, EM_GETMARGINS, 0, ByVal 0&)), vbPixels, vbContainerSize)
End Property

Public Property Let LeftMargin(ByVal Value As Single)
If Value = EC_USEFONTINFO Or Value = -1 Then
    If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETMARGINS, EC_LEFTMARGIN, ByVal EC_USEFONTINFO
Else
    If Value < 0 Then Err.Raise 380
    Dim IntValue As Integer
    On Error Resume Next
    IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
    If Err.Number <> 0 Then IntValue = 0
    On Error GoTo 0
    If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETMARGINS, EC_LEFTMARGIN, ByVal MakeDWord(IntValue, 0)
End If
UserControl.PropertyChanged "LeftMargin"
End Property

Public Property Get RightMargin() As Single
Attribute RightMargin.VB_Description = "Returns/sets the widths of the right margin."
Attribute RightMargin.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then RightMargin = UserControl.ScaleX(HiWord(SendMessage(TextBoxHandle, EM_GETMARGINS, 0, ByVal 0&)), vbPixels, vbContainerSize)
End Property

Public Property Let RightMargin(ByVal Value As Single)
If Value = EC_USEFONTINFO Or Value = -1 Then
    If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETMARGINS, EC_RIGHTMARGIN, ByVal EC_USEFONTINFO
Else
    If Value < 0 Then Err.Raise 380
    Dim IntValue As Integer
    On Error Resume Next
    IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
    If Err.Number <> 0 Then IntValue = 0
    On Error GoTo 0
    If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETMARGINS, EC_RIGHTMARGIN, ByVal MakeDWord(0, IntValue)
End If
UserControl.PropertyChanged "RightMargin"
End Property

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            If PropOLEDragMode = vbOLEDragAutomatic Then
                Dim P3 As POINTAPI
                Dim CharPos As Long, CaretPos As Long
                Dim SelStart As Long, SelEnd As Long
                GetCursorPos P3
                ScreenToClient TextBoxHandle, P3
                CharPos = CIntToUInt(LoWord(SendMessage(TextBoxHandle, EM_CHARFROMPOS, 0, ByVal MakeDWord(P3.X, P3.Y))))
                CaretPos = SendMessage(TextBoxHandle, EM_POSFROMCHAR, CharPos, ByVal 0&)
                SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
                TextBoxAutoDragInSel = CBool(CharPos >= SelStart And CharPos <= SelEnd And CaretPos > -1 And (SelEnd - SelStart) > 0)
                If TextBoxAutoDragInSel = True Then
                    SetCursor LoadCursor(0, MousePointerID(vbArrow))
                    WindowProcControl = 1
                    Exit Function
                End If
            Else
                TextBoxAutoDragInSel = False
            End If
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(0, MousePointerID(PropMousePointer))
                WindowProcControl = 1
                Exit Function
            ElseIf PropMousePointer = 99 Then
                If Not PropMouseIcon Is Nothing Then
                    SetCursor PropMouseIcon.Handle
                    WindowProcControl = 1
                    Exit Function
                End If
            End If
        End If
    Case WM_MOUSEACTIVATE
        Static InProc As Boolean
        If GetFocus() <> TextBoxHandle Then
            If InProc = True Or LoWord(lParam) = HTBORDER Then WindowProcControl = MA_NOACTIVATEANDEAT: Exit Function
            Select Case HiWord(lParam)
                Case WM_LBUTTONDOWN
                    On Error Resume Next
                    If Extender.CausesValidation = True Then
                        InProc = True
                        Screen.ActiveForm.ValidateControls
                        InProc = False
                        If Err.Number = 380 Then
                            WindowProcControl = MA_NOACTIVATEANDEAT
                        Else
                            SetFocusAPI UserControl.hWnd
                            WindowProcControl = MA_NOACTIVATE
                        End If
                    Else
                        SetFocusAPI UserControl.hWnd
                        WindowProcControl = MA_NOACTIVATE
                    End If
                    On Error GoTo 0
                    Exit Function
            End Select
        End If
    Case WM_KEYDOWN, WM_KEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftState())
        ElseIf wMsg = WM_KEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftState())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        KeyChar = CUIntToInt(wParam And &HFFFF&)
        RaiseEvent KeyPress(KeyChar)
        If (wParam And &HFFFF&) <> 0 And KeyChar = 0 Then
            Exit Function
        Else
            wParam = CIntToUInt(KeyChar)
        End If
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK
        RaiseEvent DblClick
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftState(), X, Y)
                If PropOLEDragMode = vbOLEDragAutomatic And TextBoxAutoDragInSel = True Then
                    Dim P4 As POINTAPI
                    GetCursorPos P4
                    If DragDetect(TextBoxHandle, CInt(P4.X), CInt(P4.Y)) <> 0 Then
                        Me.OLEDrag
                    Else
                        WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
                        ReleaseCapture
                    End If
                    Exit Function
                End If
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftState(), X, Y)
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftState(), X, Y)
            Case WM_MOUSEMOVE
                RaiseEvent MouseMove(GetMouseState(), GetShiftState(), X, Y)
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
                Select Case wMsg
                    Case WM_LBUTTONUP
                        RaiseEvent MouseUp(vbLeftButton, GetShiftState(), X, Y)
                    Case WM_MBUTTONUP
                        RaiseEvent MouseUp(vbMiddleButton, GetShiftState(), X, Y)
                    Case WM_RBUTTONUP
                        RaiseEvent MouseUp(vbRightButton, GetShiftState(), X, Y)
                End Select
                Dim P1 As POINTAPI
                GetCursorPos P1
                If WindowFromPoint(P1.X, P1.Y) = hWnd Then RaiseEvent Click
        End Select
    Case WM_VSCROLL, WM_HSCROLL
        ' The notification codes EN_HSCROLL and EN_VSCROLL are not sent when clicking the scroll bar thumb itself.
        If LoWord(wParam) = SB_THUMBTRACK Then RaiseEvent Scroll
    Case WM_CONTEXTMENU
        If wParam = TextBoxHandle Then
            Dim P2 As POINTAPI, Handled As Boolean
            P2.X = Get_X_lParam(lParam)
            P2.Y = Get_Y_lParam(lParam)
            If P2.X > 0 And P2.Y > 0 Then
                ScreenToClient TextBoxHandle, P2
                RaiseEvent ContextMenu(Handled, UserControl.ScaleX(P2.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P2.Y, vbPixels, vbContainerPosition))
            ElseIf P2.X = -1 And P2.Y = -1 Then
                ' According to MSDN:
                ' If the context menu is generated from the keyboard - for example
                ' if the user types SHIFT + F10 — then the X and Y coordinates
                ' are -1 and the application should display the context menu at the
                ' location of the current selection rather than at (XPos, YPos).
                RaiseEvent ContextMenu(Handled, -1, -1)
            End If
            If Handled = True Then Exit Function
        End If
    Case WM_PASTE
        If PropAllowOnlyNumbers = True Then
            If ComCtlsSupportLevel() <= 1 Then
                If VB.Clipboard.GetFormat(vbCFText) = True Then
                    Dim Text As String
                    Text = VB.Clipboard.GetText(vbCFText)
                    If Not Text = vbNullString Then
                        Dim i As Long, InvalidText As Boolean
                        For i = 1 To Len(Text)
                            If InStr("0123456789", Mid(Text, i, 1)) = 0 Then
                                InvalidText = True
                                Exit For
                            End If
                        Next i
                        If InvalidText = True Then
                            VBA.Interaction.Beep
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        Select Case HiWord(wParam)
            Case EN_CHANGE
                UserControl.PropertyChanged "Text"
                RaiseEvent Change
            Case EN_HSCROLL, EN_VSCROLL
                ' This notification code is also sent when a keyboard event causes a change in the view area.
                RaiseEvent Scroll
        End Select
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS Then SetFocusAPI TextBoxHandle
End Function
