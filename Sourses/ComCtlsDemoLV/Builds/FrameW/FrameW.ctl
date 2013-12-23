VERSION 5.00
Begin VB.UserControl FrameW 
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ControlContainer=   -1  'True
   HasDC           =   0   'False
   PropertyPages   =   "FrameW.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "FrameW.ctx":0035
End
Attribute VB_Name = "FrameW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
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
Private Type SIZEAPI
CX As Long
CY As Long
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
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
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
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const RDW_UPDATENOW As Long = &H100
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4
Private Const RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_EX_TRANSPARENT As Long = &H20
Private Const WS_EX_RTLREADING As Long = &H2000
Private Const SW_HIDE As Long = &H0
Private Const SW_SHOW As Long = &H5
Private Const WM_SETFONT As Long = &H30
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const WM_PAINT As Long = &HF
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const BS_TEXT As Long = &H0
Private Const BS_GROUPBOX As Long = &H7
Private Const BS_FLAT As Long = &H8000&
Implements ISubclass
Implements OLEGuids.IPerPropertyBrowsingVB
Private FrameGroupBoxHandle As Long
Private FrameTransparentBrush As Long
Private FrameFontHandle As Long
Private FrameLogFont As LOGFONT
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropEnabled As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropBorder As Boolean
Private PropAppearance As CCAppearanceConstants
Private PropCaption As String
Private PropTransparent As Boolean

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
Call SetVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
DispIDMousePointer = GetDispID(Me, "MousePointer")
End Sub

Private Sub UserControl_InitProperties()
Set PropFont = Ambient.Font
Set UserControl.Font = PropFont
PropVisualStyles = True
PropEnabled = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropBorder = True
PropAppearance = UserControl.Appearance
PropCaption = Ambient.DisplayName
PropTransparent = False
Call CreateFrame
End Sub

Private Sub UserControl_Paint()
If FrameGroupBoxHandle <> 0 Then RedrawWindow FrameGroupBoxHandle, 0, 0, RDW_INVALIDATE
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
Set PropFont = .ReadProperty("Font", Ambient.Font)
Set UserControl.Font = PropFont
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.BackColor = .ReadProperty("BackColor", vbButtonFace)
Me.ForeColor = .ReadProperty("ForeColor", vbButtonText)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
Me.MousePointer = PropMousePointer
PropBorder = .ReadProperty("Border", True)
PropAppearance = .ReadProperty("Appearance", CCAppearance3D)
PropCaption = VarToStr(.ReadProperty("Caption", vbNullString))
PropTransparent = .ReadProperty("Transparent", False)
End With
Call CreateFrame
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", PropFont, Ambient.Font
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", Me.BackColor, vbButtonFace
.WriteProperty "ForeColor", Me.ForeColor, vbButtonText
.WriteProperty "Enabled", PropEnabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "Border", PropBorder, True
.WriteProperty "Appearance", PropAppearance, CCAppearance3D
.WriteProperty "Caption", StrToVar(PropCaption), vbNullString
.WriteProperty "Transparent", PropTransparent, False
End With
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
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
RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
UserControl.OLEDrag
End Sub

Private Sub UserControl_Resize()
With UserControl
If FrameGroupBoxHandle <> 0 Then MoveWindow FrameGroupBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyFrame
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
hWnd = UserControl.hWnd
End Property

Public Property Get hWndGroupBox() As Long
Attribute hWndGroupBox.VB_Description = "Returns a handle to a control."
hWndGroupBox = FrameGroupBoxHandle
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
Call OLEFontToLogFont(NewFont, FrameLogFont)
OldFontHandle = FrameFontHandle
FrameFontHandle = CreateFontIndirect(FrameLogFont)
If FrameGroupBoxHandle <> 0 Then SendMessage FrameGroupBoxHandle, WM_SETFONT, FrameFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
Me.Refresh
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
Call OLEFontToLogFont(PropFont, FrameLogFont)
OldFontHandle = FrameFontHandle
FrameFontHandle = CreateFontIndirect(FrameLogFont)
If FrameGroupBoxHandle <> 0 Then SendMessage FrameGroupBoxHandle, WM_SETFONT, FrameFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
Me.Refresh
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
If FrameGroupBoxHandle <> 0 And EnabledVisualStyles() = True Then
    Select Case PropVisualStyles
        Case True
            ActivateVisualStyles FrameGroupBoxHandle
        Case False
            RemoveVisualStyles FrameGroupBoxHandle
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
Enabled = PropEnabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
PropEnabled = Value
If Ambient.UserMode = True Then UserControl.Enabled = PropEnabled
If FrameGroupBoxHandle <> 0 Then EnableWindow FrameGroupBoxHandle, IIf(PropEnabled = True, 1, 0)
UserControl.PropertyChanged "Enabled"
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
If Ambient.UserMode = True Then
    Select Case PropMousePointer
        Case vbIconPointer, 16, vbCustom
            If PropMousePointer = vbCustom Then
                Set UserControl.MouseIcon = PropMouseIcon
            Else
                Set UserControl.MouseIcon = PictureFromHandle(LoadCursor(0, MousePointerID(PropMousePointer)), vbPicTypeIcon)
            End If
            UserControl.MousePointer = vbCustom
        Case Else
            UserControl.MousePointer = PropMousePointer
    End Select
End If
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
Me.MousePointer = PropMousePointer
UserControl.PropertyChanged "MouseIcon"
End Property

Public Property Get Border() As Boolean
Attribute Border.VB_Description = "Returns/sets a value that determines whether or not the control's border and caption are drawn."
Border = PropBorder
End Property

Public Property Let Border(ByVal Value As Boolean)
PropBorder = Value
If FrameGroupBoxHandle <> 0 Then ShowWindow FrameGroupBoxHandle, IIf(PropBorder = True, SW_SHOW, SW_HIDE)
UserControl.PropertyChanged "Border"
End Property

Public Property Get Appearance() As CCAppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets a value that determines whether an object is painted two-dimensional or with 3-D effects."
Attribute Appearance.VB_UserMemId = -520
Appearance = PropAppearance
End Property

Public Property Let Appearance(ByVal Value As CCAppearanceConstants)
Select Case Value
    Case CCAppearanceFlat, CCAppearance3D
        PropAppearance = Value
    Case Else
        Err.Raise 380
End Select
UserControl.Appearance = PropAppearance
UserControl.ForeColor = IIf(PropAppearance = CCAppearanceFlat, vbWindowText, vbButtonText)
If FrameGroupBoxHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(FrameGroupBoxHandle, GWL_STYLE)
    If PropAppearance = CCAppearanceFlat Then
        If Not (dwStyle And BS_FLAT) = BS_FLAT Then dwStyle = dwStyle Or BS_FLAT
    Else
        If (dwStyle And BS_FLAT) = BS_FLAT Then dwStyle = dwStyle And Not BS_FLAT
    End If
    SetWindowLong FrameGroupBoxHandle, GWL_STYLE, dwStyle
End If
Me.Refresh
UserControl.PropertyChanged "Appearance"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_UserMemId = -518
If FrameGroupBoxHandle <> 0 Then
    Caption = String(SendMessage(FrameGroupBoxHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
    SendMessage FrameGroupBoxHandle, WM_GETTEXT, Len(Caption) + 1, ByVal StrPtr(Caption)
Else
    Caption = PropCaption
End If
End Property

Public Property Let Caption(ByVal Value As String)
PropCaption = Value
If FrameGroupBoxHandle <> 0 Then SendMessage FrameGroupBoxHandle, WM_SETTEXT, 0, ByVal StrPtr(PropCaption)
Me.Refresh
UserControl.PropertyChanged "Caption"
End Property

Public Property Get Transparent() As Boolean
Attribute Transparent.VB_Description = "Returns/sets a value indicating if the background is a replica of the underlying background to simulate transparency. This property is ignored at design time."
Transparent = PropTransparent
End Property

Public Property Let Transparent(ByVal Value As Boolean)
PropTransparent = Value
Me.Refresh
UserControl.PropertyChanged "Transparent"
End Property

Private Sub CreateFrame()
If FrameGroupBoxHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or BS_TEXT Or BS_GROUPBOX
If PropAppearance = CCAppearanceFlat Then dwStyle = dwStyle Or BS_FLAT
dwExStyle = WS_EX_TRANSPARENT
If Ambient.RightToLeft = True Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
FrameGroupBoxHandle = CreateWindowEx(dwExStyle, StrPtr("Button"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = PropEnabled
Me.Border = PropBorder
Me.Caption = PropCaption
If Ambient.UserMode = True Then Call ComCtlsSetSubclass(UserControl.hWnd, Me, 0)
End Sub

Private Sub DestroyFrame()
If FrameGroupBoxHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow FrameGroupBoxHandle, SW_HIDE
SetParent FrameGroupBoxHandle, 0
DestroyWindow FrameGroupBoxHandle
FrameGroupBoxHandle = 0
If FrameTransparentBrush <> 0 Then
    DeleteObject FrameTransparentBrush
    FrameTransparentBrush = 0
    Set UserControl.Picture = Nothing
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
If FrameTransparentBrush <> 0 Then
    DeleteObject FrameTransparentBrush
    FrameTransparentBrush = 0
    Set UserControl.Picture = Nothing
End If
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get ContainedControls() As Object
Attribute ContainedControls.VB_Description = "Returns a collection that allows access to the controls contained within the control that were added to the control by the developer who uses the control."
Set ContainedControls = UserControl.ContainedControls
End Property

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_CTLCOLORSTATIC
        WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        If PropTransparent = True Then
            SetBkMode wParam, 1
            Dim hDCScreen As Long, hDCBmp As Long
            Dim hBmp As Long, hBmpOld As Long
            With UserControl
            If FrameTransparentBrush = 0 Then
                hDCScreen = GetDC(0)
                If hDCScreen <> 0 Then
                    hDCBmp = CreateCompatibleDC(hDCScreen)
                    If hDCBmp <> 0 Then
                        hBmp = CreateCompatibleBitmap(hDCScreen, .ScaleWidth, .ScaleHeight)
                        If hBmp <> 0 Then
                            hBmpOld = SelectObject(hDCBmp, hBmp)
                            Dim P As POINTAPI
                            ClientToScreen hWnd, P
                            ScreenToClient .ContainerHwnd, P
                            SetViewportOrgEx hDCBmp, -P.X, -P.Y, P
                            SendMessage .ContainerHwnd, WM_PAINT, hDCBmp, ByVal 0&
                            SetViewportOrgEx hDCBmp, P.X, P.Y, P
                            FrameTransparentBrush = CreatePatternBrush(hBmp)
                            Dim hDCBmp2 As Long
                            Dim hBmp2 As Long, hBmpOld2 As Long
                            hDCBmp2 = CreateCompatibleDC(hDCBmp)
                            If hDCBmp2 <> 0 Then
                                hBmp2 = CreateCompatibleBitmap(hDCBmp, .ScaleWidth, .ScaleHeight)
                                If hBmp2 <> 0 Then
                                    hBmpOld2 = SelectObject(hDCBmp2, hBmp2)
                                    BitBlt hDCBmp2, 0, 0, .ScaleWidth, .ScaleHeight, hDCBmp, 0, 0, vbSrcCopy
                                    hBmp2 = SelectObject(hDCBmp2, hBmpOld2)
                                    DeleteDC hDCBmp2
                                End If
                            End If
                            Set UserControl.Picture = PictureFromHandle(hBmp2, vbPicTypeBitmap)
                            If FrameGroupBoxHandle <> 0 Then RedrawWindow FrameGroupBoxHandle, 0, 0, RDW_INVALIDATE
                            SelectObject hDCBmp, hBmpOld
                            DeleteObject hBmp
                        End If
                        DeleteDC hDCBmp
                    End If
                    ReleaseDC 0, hDCScreen
                End If
            End If
            If FrameTransparentBrush <> 0 Then WindowProcUserControl = FrameTransparentBrush
            End With
        End If
        Exit Function
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_PAINT And wParam <> 0 Then
    If FrameGroupBoxHandle <> 0 Then SendMessage FrameGroupBoxHandle, WM_PAINT, wParam, ByVal 0&
End If
End Function
