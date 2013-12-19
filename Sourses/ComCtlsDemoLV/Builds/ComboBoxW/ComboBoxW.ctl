VERSION 5.00
Begin VB.UserControl ComboBoxW 
   BackColor       =   &H80000005&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   1  'vbSimpleBound
   ForeColor       =   &H80000008&
   HasDC           =   0   'False
   PropertyPages   =   "ComboBoxW.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "ComboBoxW.ctx":0035
End
Attribute VB_Name = "ComboBoxW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If False Then
Private CboStyleDropDownCombo, CboStyleSimpleCombo, CboStyleDropDownList
#End If
Public Enum CboStyleConstants
CboStyleDropDownCombo = 0
CboStyleSimpleCombo = 1
CboStyleDropDownList = 2
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
Private Type COMBOBOXINFO
cbSize As Long
RCItem As RECT
RCButton As RECT
StateButton As Long
hWndCombo As Long
hWndItem As Long
hWndList As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Public Event DropDown()
Attribute DropDown.VB_Description = "Occurs when the drop-down list is about to drop down."
Public Event CloseUp()
Attribute CloseUp.VB_Description = "Occurs when the drop-down list has been closed."
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
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function DragDetect Lib "user32" (ByVal hWnd As Long, ByVal PX As Integer, ByVal PY As Integer) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const RDW_UPDATENOW As Long = &H100
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4
Private Const RDW_ALLCHILDREN As Long = &H80
Private Const HWND_DESKTOP As Long = &H0
Private Const GWL_STYLE As Long = (-16)
Private Const CF_UNICODETEXT As Long = 13
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_RTLREADING As Long = &H2000
Private Const SW_HIDE As Long = &H0
Private Const SW_SHOW As Long = &H5
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const WM_MOUSEACTIVATE As Long = &H21, MA_NOACTIVATE As Long = &H3, MA_NOACTIVATEANDEAT As Long = &H4
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_COMMAND As Long = &H111
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_IME_CHAR As Long = &H286
Private Const WM_CHARTOITEM As Long = &H2F
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
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_CTLCOLORLISTBOX As Long = &H134
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const WM_CLEAR As Long = &H303
Private Const EM_SETREADONLY As Long = &HCF, ES_READONLY As Long = &H800
Private Const EM_SETSEL As Long = &HB1
Private Const EM_REPLACESEL As Long = &HC2
Private Const CB_ERR As Long = (-1)
Private Const CB_LIMITTEXT As Long = &H141
Private Const CB_ADDSTRING As Long = &H143
Private Const CB_DELETESTRING As Long = &H144
Private Const CB_GETCOUNT As Long = &H146
Private Const CB_GETCURSEL As Long = &H147
Private Const CB_INSERTSTRING As Long = &H14A
Private Const CB_SETCURSEL As Long = &H14E
Private Const CB_GETDROPPEDCONTROLRECT As Long = &H152
Private Const CB_GETTOPINDEX As Long = &H15B
Private Const CB_SETTOPINDEX As Long = &H15C
Private Const CB_GETHORIZONTALEXTENT As Long = &H15D
Private Const CB_SETHORIZONTALEXTENT As Long = &H15E
Private Const CB_GETLBTEXT As Long = &H148
Private Const CB_GETLBTEXTLEN As Long = &H149
Private Const CB_GETEDITSEL As Long = &H140
Private Const CB_SETEDITSEL As Long = &H142
Private Const CB_RESETCONTENT As Long = &H14B
Private Const CB_SETITEMHEIGHT As Long = &H153
Private Const CB_GETITEMHEIGHT As Long = &H154
Private Const CB_GETDROPPEDSTATE As Long = &H157
Private Const CB_GETCOMBOBOXINFO As Long = &H164
Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Const CB_GETITEMDATA As Long = &H150
Private Const CB_SETITEMDATA As Long = &H151
Private Const CB_SETEXTENDEDUI As Long = &H155
Private Const CB_GETEXTENDEDUI As Long = &H156
Private Const CBM_FIRST As Long = &H1700
Private Const CB_SETMINVISIBLE As Long = (CBM_FIRST + 1)
Private Const CB_GETMINVISIBLE As Long = (CBM_FIRST + 2)
Private Const CB_SETCUEBANNER As Long = (CBM_FIRST + 3)
Private Const CB_GETCUEBANNER As Long = (CBM_FIRST + 4)
Private Const EM_GETSEL As Long = &HB0
Private Const EM_POSFROMCHAR As Long = &HD6
Private Const EM_CHARFROMPOS As Long = &HD7
Private Const CBS_AUTOHSCROLL As Long = &H40
Private Const CBS_SIMPLE As Long = &H1
Private Const CBS_DROPDOWN As Long = &H2
Private Const CBS_DROPDOWNLIST As Long = &H3
Private Const CBS_SORT As Long = &H100
Private Const CBS_HASSTRINGS As Long = &H200
Private Const CBS_DISABLENOSCROLL As Long = &H800
Private Const CBS_NOINTEGRALHEIGHT As Long = &H400
Private Const CBS_UPPERCASE As Long = &H2000
Private Const CBS_LOWERCASE As Long = &H4000
Private Const CBN_SELCHANGE As Long = 1
Private Const CBN_DBLCLK As Long = 2
Private Const CBN_EDITCHANGE As Long = 5
Private Const CBN_EDITUPDATE As Long = 6
Private Const CBN_DROPDOWN As Long = 7
Private Const CBN_CLOSEUP As Long = 8
Private Const CBN_SELENDOK As Long = 9
Private Const CBN_SELENDCANCEL As Long = 10
Implements ISubclass
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private ComboBoxHandle As Long, ComboBoxEditHandle As Long
Private ComboBoxListBackColorBrush As Long
Private ComboBoxFontHandle As Long
Private ComboBoxLogFont As LOGFONT
Private ComboBoxNewIndex As Long
Private ComboBoxAutoDragInSel As Boolean, ComboBoxAutoDragIsActive As Boolean
Private ComboBoxAutoDragSelStart As Integer, ComboBoxAutoDragSelEnd As Integer
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropOLEDragMode As VBRUN.OLEDragConstants
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropRedraw As Boolean
Private PropStyle As CboStyleConstants
Private PropLocked As Boolean
Private PropText As String
Private PropExtendedUI As Boolean
Private PropMaxDropDownItems As Integer
Private PropIntegralHeight As Boolean
Private PropMaxLength As Long
Private PropCueBanner As String
Private PropUseListBackColor As Boolean
Private PropUseListForeColor As Boolean
Private PropListBackColor As OLE_COLOR
Private PropListForeColor As OLE_COLOR
Private PropSorted As Boolean
Private PropHorizontalExtent As Long

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
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyTab, vbKeyReturn, vbKeyEscape
            If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
                If SendMessage(ComboBoxHandle, CB_GETDROPPEDSTATE, 0, ByVal 0&) = 0 Or PropStyle = CboStyleSimpleCombo Then
                    If IsInputKey = False Then Exit Sub
                Else
                    If PropStyle = CboStyleDropDownCombo Then SendMessage ComboBoxHandle, CB_SHOWDROPDOWN, 0, ByVal 0&
                End If
            ElseIf KeyCode = vbKeyTab Then
                If IsInputKey = False Then Exit Sub
            End If
            Dim hWnd As Long
            hWnd = GetFocus()
            If hWnd <> 0 Then
                Select Case hWnd
                    Case ComboBoxHandle, ComboBoxEditHandle
                        SendMessage hWnd, wMsg, wParam, ByVal lParam
                        Handled = True
                End Select
            End If
    End Select
End If
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
Call SetVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
DispIDMousePointer = GetDispID(Me, "MousePointer")
End Sub

Private Sub UserControl_InitProperties()
Set PropFont = Ambient.Font
PropVisualStyles = True
PropOLEDragMode = vbOLEDragManual
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropRedraw = True
PropStyle = CboStyleDropDownCombo
PropLocked = False
PropText = Ambient.DisplayName
PropExtendedUI = False
PropMaxDropDownItems = 9
PropIntegralHeight = True
PropMaxLength = 0
PropCueBanner = vbNullString
PropUseListBackColor = False
PropUseListForeColor = False
PropListBackColor = vbWindowBackground
PropListForeColor = vbWindowText
PropSorted = False
PropHorizontalExtent = 0
Call CreateComboBox
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
Set PropFont = .ReadProperty("Font", Ambient.Font)
PropVisualStyles = .ReadProperty("VisualStyles", True)
PropOLEDragMode = .ReadProperty("OLEDragMode", vbOLEDragManual)
Me.BackColor = .ReadProperty("BackColor", vbWindowBackground)
Me.ForeColor = .ReadProperty("ForeColor", vbWindowText)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropRedraw = .ReadProperty("Redraw", True)
PropStyle = .ReadProperty("Style", CboStyleDropDownCombo)
PropLocked = .ReadProperty("Locked", False)
PropText = VarToStr(.ReadProperty("Text", vbNullString))
PropExtendedUI = .ReadProperty("ExtendedUI", False)
PropMaxDropDownItems = .ReadProperty("MaxDropDownItems", 9)
PropIntegralHeight = .ReadProperty("IntegralHeight", True)
PropMaxLength = .ReadProperty("MaxLength", 0)
PropCueBanner = VarToStr(.ReadProperty("CueBanner", vbNullString))
PropUseListBackColor = .ReadProperty("UseListBackColor", False)
PropUseListForeColor = .ReadProperty("UseListForeColor", False)
PropListBackColor = .ReadProperty("ListBackColor", vbWindowBackground)
PropListForeColor = .ReadProperty("ListForeColor", vbWindowText)
PropSorted = .ReadProperty("Sorted", False)
PropHorizontalExtent = .ReadProperty("HorizontalExtent", 0)
End With
Call CreateComboBox
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", PropFont, Ambient.Font
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", Me.BackColor, vbWindowBackground
.WriteProperty "ForeColor", Me.ForeColor, vbWindowText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDragMode", PropOLEDragMode, vbOLEDragManual
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "Redraw", PropRedraw, True
.WriteProperty "Style", PropStyle, CboStyleDropDownCombo
.WriteProperty "Locked", PropLocked, False
.WriteProperty "Text", StrToVar(PropText), vbNullString
.WriteProperty "ExtendedUI", PropExtendedUI, False
.WriteProperty "MaxDropDownItems", PropMaxDropDownItems, 9
.WriteProperty "IntegralHeight", PropIntegralHeight, True
.WriteProperty "MaxLength", PropMaxLength, 0
.WriteProperty "CueBanner", StrToVar(PropCueBanner), vbNullString
.WriteProperty "UseListBackColor", PropUseListBackColor, False
.WriteProperty "UseListForeColor", PropUseListForeColor, False
.WriteProperty "ListBackColor", PropListBackColor, vbWindowBackground
.WriteProperty "ListForeColor", PropListForeColor, vbWindowText
.WriteProperty "Sorted", PropSorted, False
.WriteProperty "HorizontalExtent", PropHorizontalExtent, 0
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
If PropOLEDragMode = vbOLEDragAutomatic And ComboBoxAutoDragIsActive = True And Effect = vbDropEffectMove Then
    If ComboBoxEditHandle <> 0 Then
        SendMessage ComboBoxEditHandle, EM_SETSEL, ComboBoxAutoDragSelStart, ByVal ComboBoxAutoDragSelEnd
        SendMessage ComboBoxEditHandle, WM_CLEAR, 0, ByVal 0&
    End If
End If
RaiseEvent OLECompleteDrag(Effect)
ComboBoxAutoDragIsActive = False
ComboBoxAutoDragSelStart = 0
ComboBoxAutoDragSelEnd = 0
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
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
    ComboBoxAutoDragIsActive = True
End If
RaiseEvent OLEStartDrag(Data, AllowedEffects)
If AllowedEffects = vbDropEffectNone Then ComboBoxAutoDragIsActive = False
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
UserControl.OLEDrag
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
If Ambient.UserMode = False And PropertyName = "DisplayName" And PropStyle = CboStyleDropDownList Then
    If ComboBoxHandle <> 0 Then
        If SendMessage(ComboBoxHandle, CB_GETCOUNT, 0, ByVal 0&) > 0 Then
            Dim Buffer As String
            Buffer = Ambient.DisplayName
            SendMessage ComboBoxHandle, CB_RESETCONTENT, 0, ByVal 0&
            SendMessage ComboBoxHandle, CB_ADDSTRING, 0, ByVal StrPtr(Buffer)
            SendMessage ComboBoxHandle, CB_SETCURSEL, 0, ByVal 0&
        End If
    End If
End If
End Sub

Private Sub UserControl_Resize()
Static InProc As Boolean
If InProc = True Then Exit Sub
If ComboBoxHandle = 0 Then Exit Sub
With UserControl
Dim WndRect As RECT
If PropStyle <> CboStyleSimpleCombo Then
    MoveWindow ComboBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    GetWindowRect ComboBoxHandle, WndRect
    If (WndRect.Bottom - WndRect.Top) <> .ScaleHeight Or (WndRect.Right - WndRect.Left) <> .ScaleWidth Then
        InProc = True
        .Size .ScaleX((WndRect.Right - WndRect.Left), vbPixels, vbTwips), .ScaleY((WndRect.Bottom - WndRect.Top), vbPixels, vbTwips)
        InProc = False
    End If
Else
    If PropIntegralHeight = True Then
        Dim ListRect As RECT, EditHeight As Long, Height As Long
        MoveWindow ComboBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight + IIf(PropIntegralHeight = True, 2, 0), 1
        GetWindowRect ComboBoxHandle, WndRect
        GetWindowRect Me.hWndList, ListRect
        MapWindowPoints HWND_DESKTOP, ComboBoxHandle, ListRect, 2
        EditHeight = ListRect.Top
        Const SM_CYEDGE As Long = 46
        If (ListRect.Bottom - ListRect.Top) > (GetSystemMetrics(SM_CYEDGE) * 2) Then
            Height = EditHeight + (ListRect.Bottom - ListRect.Top)
        Else
            Height = EditHeight
        End If
        InProc = True
        .Size .ScaleX((WndRect.Right - WndRect.Left), vbPixels, vbTwips), .ScaleY(Height, vbPixels, vbTwips)
        InProc = False
    End If
    MoveWindow ComboBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    Me.Refresh
End If
End With
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableSubclass(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyComboBox
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
hWnd = ComboBoxHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get hWndEdit() As Long
Attribute hWndEdit.VB_Description = "Returns a handle to a control."
hWndEdit = ComboBoxEditHandle
End Property

Public Property Get hWndList() As Long
Attribute hWndList.VB_Description = "Returns a handle to a control."
If ComboBoxHandle <> 0 Then
    Dim CBI As COMBOBOXINFO
    CBI.cbSize = LenB(CBI)
    SendMessage ComboBoxHandle, CB_GETCOMBOBOXINFO, 0, ByVal VarPtr(CBI)
    hWndList = CBI.hWndList
End If
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
Call OLEFontToLogFont(NewFont, ComboBoxLogFont)
OldFontHandle = ComboBoxFontHandle
ComboBoxFontHandle = CreateFontIndirect(ComboBoxLogFont)
If ComboBoxHandle <> 0 Then SendMessage ComboBoxHandle, WM_SETFONT, ComboBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
Call OLEFontToLogFont(PropFont, ComboBoxLogFont)
OldFontHandle = ComboBoxFontHandle
ComboBoxFontHandle = CreateFontIndirect(ComboBoxLogFont)
If ComboBoxHandle <> 0 Then SendMessage ComboBoxHandle, WM_SETFONT, ComboBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
Call UserControl_Resize
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
If ComboBoxHandle <> 0 And EnabledVisualStyles() = True Then
    Select Case PropVisualStyles
        Case True
            ActivateVisualStyles ComboBoxHandle
        Case False
            RemoveVisualStyles ComboBoxHandle
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
If ComboBoxHandle <> 0 Then EnableWindow ComboBoxHandle, IIf(Value = True, 1, 0)
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDragMode() As VBRUN.OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this control can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
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

Public Property Get OLEDropMode() As OLEDropModeConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As OLEDropModeConstants)
Select Case Value
    Case OLEDropModeNone, OLEDropModeManual
        UserControl.OLEDropMode = Value
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

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Returns/sets a value that determines whether or not the combo box redraws when changing the items. You can speed up the creation of large lists by disabling this property before adding the items."
Redraw = PropRedraw
End Property

Public Property Let Redraw(ByVal Value As Boolean)
PropRedraw = Value
If ComboBoxHandle <> 0 And Ambient.UserMode = True Then
    SendMessage ComboBoxHandle, WM_SETREDRAW, IIf(PropRedraw = True, 1, 0), ByVal 0&
    If PropRedraw = True Then Me.Refresh
End If
End Property

Public Property Get Style() As CboStyleConstants
Attribute Style.VB_Description = "Returns/sets a value that determines the type of control and the behavior of its list box portion."
Style = PropStyle
End Property

Public Property Let Style(ByVal Value As CboStyleConstants)
Select Case Value
    Case CboStyleDropDownCombo, CboStyleSimpleCombo, CboStyleDropDownList
        If Ambient.UserMode = True Then
            Err.Raise Number:=382, Description:="Style property is read-only at run time"
        Else
            PropStyle = Value
            If ComboBoxHandle <> 0 Then Call ReCreateComboBox
        End If
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "Style"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents can be edited. This property does not have any effect if the style property is set to 'DropDownList'."
If PropStyle = CboStyleDropDownList Then Exit Property
If ComboBoxHandle <> 0 And Ambient.UserMode = True Then
    Locked = CBool((GetWindowLong(ComboBoxHandle, GWL_STYLE) And ES_READONLY) <> 0)
Else
    Locked = PropLocked
End If
End Property

Public Property Let Locked(ByVal Value As Boolean)
If PropStyle = CboStyleDropDownList Then Exit Property
PropLocked = Value
If ComboBoxHandle <> 0 And Ambient.UserMode = True Then SendMessage ComboBoxEditHandle, EM_SETREADONLY, IIf(PropLocked = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "Locked"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
Attribute Text.VB_UserMemId = -517
Attribute Text.VB_MemberFlags = "123c"
Select Case PropStyle
    Case CboStyleDropDownCombo, CboStyleSimpleCombo
        If ComboBoxHandle <> 0 Then
            Text = String(SendMessage(ComboBoxEditHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
            SendMessage ComboBoxEditHandle, WM_GETTEXT, Len(Text) + 1, ByVal StrPtr(Text)
        Else
            Text = PropText
        End If
    Case CboStyleDropDownList
        If Ambient.UserMode = True Then
            Dim SelIndex As Long
            SelIndex = SendMessage(ComboBoxHandle, CB_GETCURSEL, 0, ByVal 0&)
            If Not SelIndex = CB_ERR Then Text = Me.List(SelIndex)
        Else
            Text = Ambient.DisplayName
        End If
End Select
End Property

Public Property Let Text(ByVal Value As String)
Select Case PropStyle
    Case CboStyleDropDownCombo, CboStyleSimpleCombo
        PropText = Value
        If ComboBoxHandle <> 0 Then SendMessage ComboBoxEditHandle, WM_SETTEXT, 0, ByVal StrPtr(PropText)
    Case CboStyleDropDownList
        Exit Property
End Select
UserControl.PropertyChanged "Text"
End Property

Public Property Get ExtendedUI() As Boolean
Attribute ExtendedUI.VB_Description = "Returns/sets a value that determines whether the default UI or the extended UI is used."
If ComboBoxHandle <> 0 Then
    ExtendedUI = CBool(SendMessage(ComboBoxHandle, CB_GETEXTENDEDUI, 0, ByVal 0&) = 1)
Else
    ExtendedUI = PropExtendedUI
End If
End Property

Public Property Let ExtendedUI(ByVal Value As Boolean)
PropExtendedUI = Value
If ComboBoxHandle <> 0 Then SendMessage ComboBoxHandle, CB_SETEXTENDEDUI, IIf(PropExtendedUI = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "ExtendedUI"
End Property

Public Property Get MaxDropDownItems() As Integer
Attribute MaxDropDownItems.VB_Description = "Returns/sets the maximum number of items to be shown in the drop-down list."
MaxDropDownItems = PropMaxDropDownItems
End Property

Public Property Let MaxDropDownItems(ByVal Value As Integer)
If Value < 1 Or Value > 30 Then
    If Ambient.UserMode = False Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropMaxDropDownItems = Value
Call SetDropListHeight(True)
UserControl.PropertyChanged "MaxDropDownItems"
End Property

Public Property Get IntegralHeight() As Boolean
Attribute IntegralHeight.VB_Description = "Returns/sets a value indicating whether the control displays partial items."
IntegralHeight = PropIntegralHeight
End Property

Public Property Let IntegralHeight(ByVal Value As Boolean)
If Ambient.UserMode = True Then
    Err.Raise Number:=382, Description:="IntegralHeight property is read-only at run time"
Else
    PropIntegralHeight = Value
    If ComboBoxHandle <> 0 Then Call ReCreateComboBox
End If
UserControl.PropertyChanged "IntegralHeight"
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
If ComboBoxHandle <> 0 Then SendMessage ComboBoxHandle, CB_LIMITTEXT, IIf(PropMaxLength = 0, 255, PropMaxLength), ByVal 0&
UserControl.PropertyChanged "MaxLength"
End Property

Public Property Get CueBanner() As String
Attribute CueBanner.VB_Description = "Returns/sets the textual cue, or tip, that is displayed to prompt the user for information. Requires comctl32.dll version 6.1 or higher."
CueBanner = PropCueBanner
End Property

Public Property Let CueBanner(ByVal Value As String)
PropCueBanner = Value
If ComboBoxHandle <> 0 And ComCtlsSupportLevel() >= 2 Then SendMessage ComboBoxHandle, CB_SETCUEBANNER, 0, ByVal StrPtr(PropCueBanner)
UserControl.PropertyChanged "CueBanner"
End Property

Public Property Get UseListBackColor() As Boolean
Attribute UseListBackColor.VB_Description = "Returns/sets a value which determines if the combo box control will use the list back color property."
UseListBackColor = PropUseListBackColor
End Property

Public Property Let UseListBackColor(ByVal Value As Boolean)
PropUseListBackColor = Value
Me.Refresh
UserControl.PropertyChanged "UseListBackColor"
End Property

Public Property Get UseListForeColor() As Boolean
Attribute UseListForeColor.VB_Description = "Returns/sets a value which determines if the combo box control will use the list fore color property."
UseListForeColor = PropUseListForeColor
End Property

Public Property Let UseListForeColor(ByVal Value As Boolean)
PropUseListForeColor = Value
Me.Refresh
UserControl.PropertyChanged "UseListForeColor"
End Property

Public Property Get ListBackColor() As OLE_COLOR
Attribute ListBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in the control's list portion. This property is ignored at design time."
ListBackColor = PropListBackColor
End Property

Public Property Let ListBackColor(ByVal Value As OLE_COLOR)
PropListBackColor = Value
If ComboBoxHandle <> 0 Then
    If ComboBoxListBackColorBrush <> 0 Then DeleteObject ComboBoxListBackColorBrush
    ComboBoxListBackColorBrush = CreateSolidBrush(WinColor(PropListBackColor))
    Me.Refresh
End If
UserControl.PropertyChanged "ListBackColor"
End Property

Public Property Get ListForeColor() As OLE_COLOR
Attribute ListForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in the control's list portion. This property is ignored at design time."
ListForeColor = PropListForeColor
End Property

Public Property Let ListForeColor(ByVal Value As OLE_COLOR)
PropListForeColor = Value
If ComboBoxHandle <> 0 Then Me.Refresh
UserControl.PropertyChanged "ListForeColor"
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
Sorted = PropSorted
End Property

Public Property Let Sorted(ByVal Value As Boolean)
PropSorted = Value
If ComboBoxHandle <> 0 Then Call ReCreateComboBox
UserControl.PropertyChanged "Sorted"
End Property

Public Property Get HorizontalExtent() As Single
Attribute HorizontalExtent.VB_Description = "Returns/sets the width by which a drop-down list can be scrolled horizontally."
If ComboBoxHandle <> 0 Then
    HorizontalExtent = UserControl.ScaleX(SendMessage(ComboBoxHandle, CB_GETHORIZONTALEXTENT, 0, ByVal 0&), vbPixels, vbContainerSize)
Else
    HorizontalExtent = UserControl.ScaleX(PropHorizontalExtent, vbPixels, vbContainerSize)
End If
End Property

Public Property Let HorizontalExtent(ByVal Value As Single)
If Value < 0 Then
    If Ambient.UserMode = False Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropHorizontalExtent = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If ComboBoxHandle <> 0 Then SendMessage ComboBoxHandle, CB_SETHORIZONTALEXTENT, PropHorizontalExtent, ByVal 0&
UserControl.PropertyChanged "HorizontalExtent"
End Property

Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to the combo box."
If ComboBoxHandle <> 0 Then
    Dim RetVal As Long
    If IsMissing(Index) = True Then
        RetVal = SendMessage(ComboBoxHandle, CB_ADDSTRING, 0, ByVal StrPtr(Item))
    Else
        Dim IndexLong As Long
        Select Case VarType(Index)
            Case vbLong, vbInteger, vbByte
                If Index >= 0 Then
                    IndexLong = Index
                Else
                    Err.Raise 5
                End If
            Case vbString
                IndexLong = CLng(Index)
                If IndexLong < 0 Then Err.Raise 5
            Case Else
                Err.Raise 13
        End Select
        RetVal = SendMessage(ComboBoxHandle, CB_INSERTSTRING, IndexLong, ByVal StrPtr(Item))
    End If
    If Not RetVal = CB_ERR Then
        ComboBoxNewIndex = RetVal
    Else
        Err.Raise 5
    End If
    Call SetDropListHeight(False)
End If
End Sub

Public Sub RemoveItem(ByVal Index As Long)
Attribute RemoveItem.VB_Description = "Removes an item from the combo box."
If ComboBoxHandle <> 0 Then
    If Index >= 0 Then
        If Not SendMessage(ComboBoxHandle, CB_DELETESTRING, Index, ByVal 0&) = CB_ERR Then
            ComboBoxNewIndex = -1
        Else
            Err.Raise 5
        End If
    Else
        Err.Raise 5
    End If
End If
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of the combo box."
If ComboBoxHandle <> 0 Then
    SendMessage ComboBoxHandle, CB_RESETCONTENT, 0, ByVal 0&
    ComboBoxNewIndex = -1
End If
End Sub

Public Property Get ListCount() As Long
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
If ComboBoxHandle <> 0 Then ListCount = SendMessage(ComboBoxHandle, CB_GETCOUNT, 0, ByVal 0&)
End Property

Public Property Get List(ByVal Index As Long) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
Attribute List.VB_MemberFlags = "400"
If ComboBoxHandle <> 0 Then
    Dim Length As Long
    Length = SendMessage(ComboBoxHandle, CB_GETLBTEXTLEN, Index, ByVal 0&)
    If Not Length = CB_ERR Then
        List = String(Length, vbNullChar)
        SendMessage ComboBoxHandle, CB_GETLBTEXT, Index, ByVal StrPtr(List)
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Let List(ByVal Index As Long, ByVal Value As String)
If ComboBoxHandle <> 0 Then
    If Index > -1 Then
        Dim SelIndex As Long, ItemData As Long
        SelIndex = SendMessage(ComboBoxHandle, CB_GETCURSEL, 0, ByVal 0&)
        ItemData = SendMessage(ComboBoxHandle, CB_GETITEMDATA, Index, ByVal 0&)
        If Not SendMessage(ComboBoxHandle, CB_DELETESTRING, Index, ByVal 0&) = CB_ERR Then
            SendMessage ComboBoxHandle, CB_INSERTSTRING, Index, ByVal StrPtr(Value)
            SendMessage ComboBoxHandle, CB_SETCURSEL, SelIndex, ByVal 0&
            SendMessage ComboBoxHandle, CB_SETITEMDATA, Index, ByVal ItemData
        Else
            Err.Raise 5
        End If
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
Attribute ListIndex.VB_MemberFlags = "400"
If ComboBoxHandle <> 0 Then ListIndex = SendMessage(ComboBoxHandle, CB_GETCURSEL, 0, ByVal 0&)
End Property

Public Property Let ListIndex(ByVal Value As Long)
If ComboBoxHandle <> 0 Then
    If Not Value = -1 Then
        If SendMessage(ComboBoxHandle, CB_SETCURSEL, Value, ByVal 0&) = CB_ERR Then Err.Raise 380
    Else
        SendMessage ComboBoxHandle, CB_SETCURSEL, -1, ByVal 0&
    End If
End If
End Property

Public Property Get ItemData(ByVal Index As Long) As Long
Attribute ItemData.VB_Description = "Returns/sets a specific number for each item in a combo box."
If ComboBoxHandle <> 0 Then
    If Not SendMessage(ComboBoxHandle, CB_GETLBTEXTLEN, Index, ByVal 0&) = CB_ERR Then
        ItemData = SendMessage(ComboBoxHandle, CB_GETITEMDATA, Index, ByVal 0&)
    Else
        Err.Raise 381
    End If
End If
End Property

Public Property Let ItemData(ByVal Index As Long, ByVal Value As Long)
If ComboBoxHandle <> 0 Then
    If Not SendMessage(ComboBoxHandle, CB_GETLBTEXTLEN, Index, ByVal 0&) = CB_ERR Then
        SendMessage ComboBoxHandle, CB_SETITEMDATA, Index, ByVal Value
    Else
        Err.Raise 381
    End If
End If
End Property

Private Sub CreateComboBox()
If ComboBoxHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or CBS_AUTOHSCROLL Or WS_VSCROLL Or WS_HSCROLL
Select Case PropStyle
    Case CboStyleDropDownCombo
        dwStyle = dwStyle Or CBS_DROPDOWN
    Case CboStyleSimpleCombo
        dwStyle = dwStyle Or CBS_SIMPLE
    Case CboStyleDropDownList
        PropLocked = False
        dwStyle = dwStyle Or CBS_DROPDOWNLIST
End Select
If PropIntegralHeight = False Then dwStyle = dwStyle Or CBS_NOINTEGRALHEIGHT
If PropSorted = True Then dwStyle = dwStyle Or CBS_SORT
If Ambient.RightToLeft = True Then dwExStyle = WS_EX_RTLREADING
ComboBoxHandle = CreateWindowEx(dwExStyle, StrPtr("ComboBox"), StrPtr("Combo Box"), dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If ComboBoxHandle <> 0 Then
    If PropStyle = CboStyleDropDownCombo Then
        Dim CBI As COMBOBOXINFO
        CBI.cbSize = LenB(CBI)
        SendMessage ComboBoxHandle, CB_GETCOMBOBOXINFO, 0, ByVal VarPtr(CBI)
        ComboBoxEditHandle = CBI.hWndItem
    ElseIf PropStyle = CboStyleSimpleCombo Then
        ComboBoxEditHandle = FindWindowEx(ComboBoxHandle, 0, StrPtr("Edit"), 0)
        If PropIntegralHeight = False Then MoveWindow ComboBoxHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight + 2, 1
    End If
    If (ComboBoxEditHandle = 0 And PropStyle <> CboStyleDropDownList) Then
        ShowWindow ComboBoxHandle, SW_HIDE
        SetParent ComboBoxHandle, 0
        DestroyWindow ComboBoxHandle
        ComboBoxHandle = 0
        ComboBoxEditHandle = 0
        Exit Sub
    End If
    SendMessage ComboBoxHandle, CB_LIMITTEXT, IIf(PropMaxLength = 0, 255, PropMaxLength), ByVal 0&
    If PropHorizontalExtent > 0 Then SendMessage ComboBoxHandle, CB_SETHORIZONTALEXTENT, PropHorizontalExtent, ByVal 0&
    ComboBoxNewIndex = -1
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
If PropRedraw = False Then Me.Redraw = False
If PropLocked = True Then Me.Locked = PropLocked
Me.Text = PropText
Me.ExtendedUI = PropExtendedUI
Me.MaxDropDownItems = PropMaxDropDownItems
If Not PropCueBanner = vbNullString Then Me.CueBanner = PropCueBanner
If Ambient.UserMode = True Then
    If ComboBoxHandle <> 0 Then
        If ComboBoxListBackColorBrush = 0 Then ComboBoxListBackColorBrush = CreateSolidBrush(WinColor(PropListBackColor))
        Call ComCtlsSetSubclass(ComboBoxHandle, Me, 1)
        If ComboBoxEditHandle <> 0 Then Call ComCtlsSetSubclass(ComboBoxEditHandle, Me, 2)
    End If
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 3)
Else
    If PropStyle = CboStyleDropDownList Then
        If ComboBoxHandle <> 0 Then
            Dim Buffer As String
            Buffer = Ambient.DisplayName
            SendMessage ComboBoxHandle, CB_ADDSTRING, 0, ByVal StrPtr(Buffer)
            SendMessage ComboBoxHandle, CB_SETCURSEL, 0, ByVal 0&
        End If
    End If
End If
End Sub

Private Sub ReCreateComboBox()
If Ambient.UserMode = True Then
    Dim Visible As Boolean
    Visible = Extender.Visible
    With Me
    If Visible = True Then SendMessage UserControl.hWnd, WM_SETREDRAW, 0, ByVal 0&
    Dim ListArr() As String, ItemDataArr() As Long
    Dim ItemHeight As Long, ListIndex As Long, Text As String, SelStart As Long, SelEnd As Long, NewIndex As Long
    Dim Count As Long, i As Long
    If ComboBoxHandle <> 0 Then
        Count = SendMessage(ComboBoxHandle, CB_GETCOUNT, 0, ByVal 0&)
        If Count > 0 Then
            ReDim ListArr(0 To (Count - 1)) As String
            ReDim ItemDataArr(0 To (Count - 1)) As Long
            For i = 0 To (Count - 1)
                ListArr(i) = .List(i)
                ItemDataArr(i) = SendMessage(ComboBoxHandle, CB_GETITEMDATA, i, ByVal 0&)
            Next i
        End If
        ItemHeight = SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, 0, ByVal 0&)
        ListIndex = .ListIndex
        Text = .Text
        If ComboBoxEditHandle <> 0 Then SendMessage ComboBoxHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    End If
    NewIndex = ComboBoxNewIndex
    Call DestroyComboBox
    Call CreateComboBox
    Call UserControl_Resize
    If Count > 0 Then
        SendMessage ComboBoxHandle, WM_SETREDRAW, 0, ByVal 0&
        For i = 0 To (Count - 1)
            SendMessage ComboBoxHandle, CB_INSERTSTRING, i, ByVal StrPtr(ListArr(i))
            SendMessage ComboBoxHandle, CB_SETITEMDATA, i, ByVal ItemDataArr(i)
        Next i
        SendMessage ComboBoxHandle, WM_SETREDRAW, 1, ByVal 0&
    End If
    If ComboBoxHandle <> 0 Then
        SendMessage ComboBoxHandle, CB_SETITEMHEIGHT, 0, ByVal ItemHeight
        .ListIndex = ListIndex
        .Text = Text
        If ComboBoxEditHandle <> 0 Then SendMessage ComboBoxEditHandle, EM_SETSEL, SelStart, ByVal SelEnd
    End If
    ComboBoxNewIndex = NewIndex
    If Visible = True Then SendMessage UserControl.hWnd, WM_SETREDRAW, 1, ByVal 0&
    .Refresh
    If PropRedraw = False Then .Redraw = PropRedraw
    End With
Else
    Call DestroyComboBox
    Call CreateComboBox
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyComboBox()
If ComboBoxHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(ComboBoxHandle)
If ComboBoxEditHandle <> 0 Then Call ComCtlsRemoveSubclass(ComboBoxEditHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow ComboBoxHandle, SW_HIDE
SetParent ComboBoxHandle, 0
DestroyWindow ComboBoxHandle
ComboBoxHandle = 0
ComboBoxEditHandle = 0
If ComboBoxListBackColorBrush <> 0 Then
    DeleteObject ComboBoxListBackColorBrush
    ComboBoxListBackColorBrush = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected; indicates the position of the insertion point if no text is selected."
Attribute SelStart.VB_MemberFlags = "400"
If ComboBoxHandle <> 0 And ComboBoxEditHandle <> 0 Then SendMessage ComboBoxHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal 0&
End Property

Public Property Let SelStart(ByVal Value As Long)
If ComboBoxHandle <> 0 And ComboBoxEditHandle <> 0 Then
    If Value >= 0 Then
        SendMessage ComboBoxEditHandle, EM_SETSEL, Value, ByVal Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
If ComboBoxHandle <> 0 And ComboBoxEditHandle <> 0 Then
    Dim SelStart As Long, SelEnd As Long
    SendMessage ComboBoxHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    SelLength = SelEnd - SelStart
End If
End Property

Public Property Let SelLength(ByVal Value As Long)
If ComboBoxHandle <> 0 And ComboBoxEditHandle <> 0 Then
    If Value >= 0 Then
        Dim SelStart As Long
        SendMessage ComboBoxHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal 0&
        SendMessage ComboBoxEditHandle, EM_SETSEL, SelStart, ByVal SelStart + Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_MemberFlags = "400"
If ComboBoxHandle <> 0 Then
    If ComboBoxEditHandle <> 0 Then
        Dim SelStart As Long, SelEnd As Long
        SendMessage ComboBoxHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
        On Error Resume Next
        SelText = Mid$(Me.Text, SelStart + 1, (SelEnd - SelStart))
        On Error GoTo 0
    Else
        SelText = Me.Text
    End If
End If
End Property

Public Property Let SelText(ByVal Value As String)
If ComboBoxHandle <> 0 Then
    If ComboBoxEditHandle <> 0 Then
        SendMessage ComboBoxEditHandle, EM_REPLACESEL, 0, ByVal StrPtr(Value)
    Else
        Me.Text = Value
    End If
End If
End Property

Public Property Get ItemHeight() As Single
Attribute ItemHeight.VB_Description = "Returns/sets the height of an item in the drop-down list."
Attribute ItemHeight.VB_MemberFlags = "400"
If ComboBoxHandle <> 0 Then ItemHeight = UserControl.ScaleY(SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, 0, ByVal 0&), vbPixels, vbContainerSize)
End Property

Public Property Let ItemHeight(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If ComboBoxHandle <> 0 Then SendMessage ComboBoxHandle, CB_SETITEMHEIGHT, 0, ByVal CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels))
Call SetDropListHeight(True)
End Property

Public Property Get DroppedDown() As Boolean
Attribute DroppedDown.VB_Description = "Returns/sets a value that determines whether the drop-down list is dropped down or not."
Attribute DroppedDown.VB_MemberFlags = "400"
If ComboBoxHandle <> 0 Then DroppedDown = CBool(SendMessage(ComboBoxHandle, CB_GETDROPPEDSTATE, 0, ByVal 0&) <> 0)
End Property

Public Property Let DroppedDown(ByVal Value As Boolean)
If ComboBoxHandle <> 0 Then SendMessage ComboBoxHandle, CB_SHOWDROPDOWN, IIf(Value = True, 1, 0), ByVal 0&
End Property

Public Property Get NewIndex() As Long
Attribute NewIndex.VB_Description = "Returns the index of the item most recently added to a control."
Attribute NewIndex.VB_MemberFlags = "400"
NewIndex = ComboBoxNewIndex
End Property

Public Property Get TopIndex() As Long
Attribute TopIndex.VB_Description = "Returns/sets which item in a control is displayed in the topmost position."
Attribute TopIndex.VB_MemberFlags = "400"
If ComboBoxHandle <> 0 Then TopIndex = SendMessage(ComboBoxHandle, CB_GETTOPINDEX, 0, ByVal 0&)
End Property

Public Property Let TopIndex(ByVal Value As Long)
If ComboBoxHandle <> 0 Then
    If Value >= 0 Then
        If SendMessage(ComboBoxHandle, CB_SETTOPINDEX, Value, ByVal 0&) = CB_ERR Then Err.Raise 380
    Else
        Err.Raise 380
    End If
End If
End Property

Private Sub SetDropListHeight(ByVal Calculate As Boolean)
If ComboBoxHandle <> 0 Then
    Static LastCount As Long, ItemHeight As Long
    Dim Count As Long
    Count = SendMessage(ComboBoxHandle, CB_GETCOUNT, 0, ByVal 0&)
    Select Case Count
        Case 0
            Count = 1
        Case Is > PropMaxDropDownItems
            Count = PropMaxDropDownItems
    End Select
    If Calculate = False Then
        If Count = LastCount Then Exit Sub
    Else
        ItemHeight = SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, 0, ByVal 0&)
    End If
    If PropStyle <> CboStyleSimpleCombo Then
        MoveWindow ComboBoxHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight + (ItemHeight * Count) + 2, 1
        If PropIntegralHeight = True And ComCtlsSupportLevel() >= 1 Then SendMessage ComboBoxHandle, CB_SETMINVISIBLE, PropMaxDropDownItems, ByVal 0&
    Else
        RedrawWindow ComboBoxHandle, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
    End If
    LastCount = Count
End If
End Sub

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcEdit(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_MOUSEACTIVATE
        Static InProc As Boolean
        If GetFocus() <> ComboBoxHandle And (GetFocus() <> ComboBoxEditHandle Or ComboBoxEditHandle = 0) Then
            If InProc = True Then WindowProcControl = MA_NOACTIVATEANDEAT: Exit Function
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
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
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
    Case WM_CTLCOLORLISTBOX
        WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        If PropUseListBackColor = True Or PropUseListForeColor = True Then SetBkMode wParam, 1
        If PropUseListForeColor = True Then SetTextColor wParam, WinColor(Me.ListForeColor)
        If PropUseListBackColor = True And ComboBoxListBackColorBrush <> 0 Then WindowProcControl = ComboBoxListBackColorBrush
        Exit Function
    Case WM_KEYDOWN, WM_KEYUP
        If PropStyle = CboStyleDropDownList Then
            Dim KeyCode As Integer
            KeyCode = wParam And &HFF&
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftState())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftState())
            End If
            wParam = KeyCode
        End If
    Case WM_CHAR
        If PropStyle = CboStyleDropDownList Then
            Dim KeyChar As Integer
            KeyChar = CUIntToInt(wParam And &HFFFF&)
            RaiseEvent KeyPress(KeyChar)
            wParam = CIntToUInt(KeyChar)
        End If
    Case WM_IME_CHAR
        If PropStyle = CboStyleDropDownList Then
            SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
            Exit Function
        End If
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftState(), X, Y)
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
        End Select
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcEdit(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> ComboBoxHandle Then SetFocusAPI UserControl.hWnd: Exit Function
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            If PropOLEDragMode = vbOLEDragAutomatic Then
                Dim P1 As POINTAPI
                Dim CharPos As Long, CaretPos As Long
                Dim SelStart As Long, SelEnd As Long
                GetCursorPos P1
                ScreenToClient ComboBoxEditHandle, P1
                CharPos = LoWord(SendMessage(ComboBoxEditHandle, EM_CHARFROMPOS, 0, ByVal MakeDWord(P1.X, P1.Y)))
                CaretPos = SendMessage(ComboBoxEditHandle, EM_POSFROMCHAR, CharPos, ByVal 0&)
                SendMessage ComboBoxEditHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
                ComboBoxAutoDragInSel = CBool(CharPos >= SelStart And CharPos <= SelEnd And CaretPos > -1 And (SelEnd - SelStart) > 0)
                If ComboBoxAutoDragInSel = True Then
                    ComboBoxAutoDragSelStart = SelStart
                    ComboBoxAutoDragSelEnd = SelEnd
                    SetCursor LoadCursor(0, MousePointerID(vbArrow))
                    WindowProcEdit = 1
                    Exit Function
                End If
            Else
                ComboBoxAutoDragInSel = False
            End If
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
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftState(), X, Y)
                If PropOLEDragMode = vbOLEDragAutomatic And ComboBoxAutoDragInSel = True Then
                    Dim P2 As POINTAPI
                    GetCursorPos P2
                    If DragDetect(ComboBoxEditHandle, CInt(P2.X), CInt(P2.Y)) <> 0 Then
                        Me.OLEDrag
                    Else
                        WindowProcEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
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
        End Select
End Select
WindowProcEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        Dim hWndFocus As Long
        Select Case HiWord(wParam)
            Case CBN_SELCHANGE
                Dim SelIndex As Long
                SelIndex = SendMessage(lParam, CB_GETCURSEL, 0, ByVal 0&)
                If Not SelIndex = CB_ERR Then
                    Me.Text = Me.List(SelIndex)
                    RaiseEvent Click
                End If
            Case CBN_DBLCLK
                RaiseEvent DblClick
            Case CBN_EDITCHANGE
                UserControl.PropertyChanged "Text"
                RaiseEvent Change
            Case CBN_DROPDOWN
                RaiseEvent DropDown
            Case CBN_CLOSEUP
                RaiseEvent CloseUp
        End Select
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS Then SetFocusAPI ComboBoxHandle
End Function
