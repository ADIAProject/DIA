VERSION 5.00
Begin VB.UserControl OptionButtonW 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   HasDC           =   0   'False
   PropertyPages   =   "OptionButtonW.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "OptionButtonW.ctx":0035
   Begin VB.Timer TimerImageList 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "OptionButtonW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If False Then
Private OptImageListAlignmentLeft, OptImageListAlignmentRight, OptImageListAlignmentTop, OptImageListAlignmentBottom, OptImageListAlignmentCenter
#End If
Private Const BUTTON_IMAGELIST_ALIGN_LEFT As Long = 0
Private Const BUTTON_IMAGELIST_ALIGN_RIGHT As Long = 1
Private Const BUTTON_IMAGELIST_ALIGN_TOP As Long = 2
Private Const BUTTON_IMAGELIST_ALIGN_BOTTOM As Long = 3
Private Const BUTTON_IMAGELIST_ALIGN_CENTER As Long = 4
Public Enum OptImageListAlignmentConstants
OptImageListAlignmentLeft = BUTTON_IMAGELIST_ALIGN_LEFT
OptImageListAlignmentRight = BUTTON_IMAGELIST_ALIGN_RIGHT
OptImageListAlignmentTop = BUTTON_IMAGELIST_ALIGN_TOP
OptImageListAlignmentBottom = BUTTON_IMAGELIST_ALIGN_BOTTOM
OptImageListAlignmentCenter = BUTTON_IMAGELIST_ALIGN_CENTER
End Enum
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
Private Type BUTTON_IMAGELIST
hImageList As Long
RCMargin As RECT
uAlign As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Type NMBCHOTITEM
hdr As NMHDR
dwFlags As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event HotChanged()
Attribute HotChanged.VB_Description = "Occurrs when the option button control's hot state changes. Requires comctl32.dll version 6.0 or higher."
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
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
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
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const RDW_UPDATENOW As Long = &H100
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4
Private Const RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_RTLREADING As Long = &H2000
Private Const SW_HIDE As Long = &H0
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_MOUSEACTIVATE As Long = &H21, MA_NOACTIVATE As Long = &H3, MA_NOACTIVATEANDEAT As Long = &H4
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_UNICHAR As Long = &H109, UNICODE_NOCHAR As Long = &HFFFF&
Private Const WM_IME_CHAR As Long = &H286
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_COMMAND As Long = &H111
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const WM_CTLCOLORBTN As Long = &H135
Private Const WM_PAINT As Long = &HF
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const BS_TEXT As Long = &H0
Private Const BS_RADIOBUTTON As Long = &H4
Private Const BS_RIGHTBUTTON As Long = &H20
Private Const BS_ICON As Long = &H40
Private Const BS_BITMAP As Long = &H80
Private Const BS_LEFT As Long = &H100
Private Const BS_RIGHT As Long = &H200
Private Const BS_CENTER As Long = &H300
Private Const BS_TOP As Long = &H400
Private Const BS_VCENTER As Long = &HC00
Private Const BS_BOTTOM As Long = &H800
Private Const BS_PUSHLIKE As Long = &H1000
Private Const BS_MULTILINE As Long = &H2000
Private Const BS_NOTIFY As Long = &H4000
Private Const BS_FLAT As Long = &H8000&
Private Const BM_GETCHECK As Long = &HF0
Private Const BM_SETCHECK As Long = &HF1
Private Const BM_GETSTATE As Long = &HF2
Private Const BM_SETSTATE As Long = &HF3
Private Const BM_SETIMAGE As Long = &HF7
Private Const BCM_FIRST As Long = &H1600
Private Const BCM_SETIMAGELIST As Long = (BCM_FIRST + 2)
Private Const BCM_GETIMAGELIST As Long = (BCM_FIRST + 3)
Private Const BST_UNCHECKED As Long = &H0
Private Const BST_CHECKED As Long = &H1
Private Const BST_INDETERMINATE As Long = &H2
Private Const BST_PUSHED As Long = &H4
Private Const BST_HOT As Long = &H200
Private Const BCCL_NOGLYPH As Long = (-1) ' Contrary to MSDN it even works on Windows XP
Private Const BN_CLICKED As Long = 0
Private Const BN_DOUBLECLICKED As Long = 5
Private Const BCN_FIRST As Long = -1250
Private Const BCN_HOTITEMCHANGE As Long = (BCN_FIRST + 1)
Private Const HICF_MOUSE As Long = &H1
Private Const HICF_ENTERING As Long = &H10
Private Const HICF_LEAVING As Long = &H20
Private Const IMAGE_BITMAP As Long = 0
Private Const IMAGE_ICON As Long = 1
Implements ISubclass
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private OptionButtonHandle As Long
Private OptionButtonTransparentBrush As Long
Private OptionButtonIgnoreClick As Boolean
Private OptionButtonFontHandle As Long
Private OptionButtonCharCodeCache As Long
Private DispIDMousePointer As Long
Private DispIDImageList As Long, ImageListArray() As String
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropImageListName As String, PropImageListControl As Object, PropImageListInit As Boolean
Private PropImageListAlignment As OptImageListAlignmentConstants
Private PropImageListMargin As Long
Private PropValue As OLE_OPTEXCLUSIVE
Private PropCaption As String
Private PropAlignment As CCLeftRightAlignmentConstants
Private PropTextAlignment As VBRUN.AlignmentConstants
Private PropPushLike As Boolean
Private PropPicture As IPictureDisp
Private PropWordWrap As Boolean
Private PropTransparent As Boolean
Private PropAppearance As CCAppearanceConstants
Private PropVerticalAlignment As CCVerticalAlignmentConstants

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
            If IsInputKey = True Then
                If OptionButtonHandle <> 0 Then
                    SendMessage OptionButtonHandle, wMsg, wParam, ByVal lParam
                    Handled = True
                End If
            End If
    End Select
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispID As Long, ByRef DisplayName As String)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetDisplayStringMousePointer(PropMousePointer, DisplayName)
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetPredefinedStringsMousePointer(StringsOut(), CookiesOut())
    Handled = True
ElseIf DispID = DispIDImageList Then
    Dim ControlEnum As Object
    Dim PropUBound As Long
    On Error GoTo CATCH_EXCEPTION
    PropUBound = UBound(StringsOut())
    ReDim Preserve StringsOut(PropUBound + 1) As String
    ReDim Preserve CookiesOut(PropUBound + 1) As Long
    StringsOut(PropUBound) = "(None)"
    CookiesOut(PropUBound) = PropUBound
    For Each ControlEnum In UserControl.ParentControls
        If TypeName(ControlEnum) = "ImageList" Then
            PropUBound = UBound(StringsOut())
            ReDim Preserve StringsOut(PropUBound + 1) As String
            ReDim Preserve CookiesOut(PropUBound + 1) As Long
            StringsOut(PropUBound) = ProperControlName(ControlEnum)
            CookiesOut(PropUBound) = PropUBound
        End If
    Next ControlEnum
    On Error GoTo 0
    Dim i As Long
    ReDim ImageListArray(0 To UBound(StringsOut()))
    For i = 0 To UBound(StringsOut())
        ImageListArray(i) = StringsOut(i)
    Next i
    Handled = True
End If
Exit Sub
CATCH_EXCEPTION:
Handled = False
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispID As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispID = DispIDMousePointer Then
    Value = Cookie
    Handled = True
ElseIf DispID = DispIDImageList Then
    If Cookie < UBound(ImageListArray()) Then Value = ImageListArray(Cookie)
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_STANDARD_CLASSES)
Call SetVTableSubclass(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
DispIDMousePointer = GetDispID(Me, "MousePointer")
DispIDImageList = GetDispID(Me, "ImageList")
ReDim ImageListArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropImageListName = "(None)": Set PropImageListControl = Nothing
PropImageListAlignment = OptImageListAlignmentLeft
PropImageListMargin = 0
PropValue = False
PropCaption = Ambient.DisplayName
PropAlignment = CCLeftRightAlignmentLeft
PropTextAlignment = vbLeftJustify
PropPushLike = False
Set PropPicture = Nothing
PropWordWrap = True
PropTransparent = False
PropAppearance = UserControl.Appearance
PropVerticalAlignment = CCVerticalAlignmentCenter
Call CreateOptionButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
Set PropFont = .ReadProperty("Font", Ambient.Font)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.BackColor = .ReadProperty("BackColor", vbButtonFace)
Me.ForeColor = .ReadProperty("ForeColor", vbButtonText)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropImageListName = .ReadProperty("ImageList", "(None)")
PropImageListAlignment = .ReadProperty("ImageListAlignment", OptImageListAlignmentLeft)
PropImageListMargin = .ReadProperty("ImageListMargin", 0)
PropValue = .ReadProperty("Value", True)
PropCaption = VarToStr(.ReadProperty("Caption", vbNullString))
PropAlignment = .ReadProperty("Alignment", CCLeftRightAlignmentLeft)
PropTextAlignment = .ReadProperty("TextAlignment", vbLeftJustify)
PropPushLike = .ReadProperty("PushLike", False)
Set PropPicture = .ReadProperty("Picture", Nothing)
PropWordWrap = .ReadProperty("WordWrap", True)
PropTransparent = .ReadProperty("Transparent", False)
PropAppearance = .ReadProperty("Appearance", CCAppearance3D)
PropVerticalAlignment = .ReadProperty("VerticalAlignment", CCVerticalAlignmentCenter)
End With
Call CreateOptionButton
If Not PropImageListName = "(None)" Then TimerImageList.Enabled = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", PropFont, Ambient.Font
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", Me.BackColor, vbButtonFace
.WriteProperty "ForeColor", Me.ForeColor, vbButtonText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "ImageList", PropImageListName, "(None)"
.WriteProperty "ImageListAlignment", PropImageListAlignment, OptImageListAlignmentLeft
.WriteProperty "ImageListMargin", PropImageListMargin, 0
.WriteProperty "Value", PropValue, True
.WriteProperty "Caption", StrToVar(PropCaption), vbNullString
.WriteProperty "Alignment", PropAlignment, CCLeftRightAlignmentLeft
.WriteProperty "TextAlignment", PropTextAlignment, vbLeftJustify
.WriteProperty "PushLike", PropPushLike, False
.WriteProperty "Picture", PropPicture, Nothing
.WriteProperty "WordWrap", PropWordWrap, True
.WriteProperty "Transparent", PropTransparent, False
.WriteProperty "Appearance", PropAppearance, CCAppearance3D
.WriteProperty "VerticalAlignment", PropVerticalAlignment, CCVerticalAlignmentCenter
End With
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
If OptionButtonHandle <> 0 Then MoveWindow OptionButtonHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableSubclass(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyOptionButton
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerImageList_Timer()
If PropImageListInit = False Then
    Me.ImageList = PropImageListName
    PropImageListInit = True
End If
TimerImageList.Enabled = False
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

Public Property Get Container() As Object
Attribute Container.VB_Description = "Returns the container of an object."
Set Container = Extender.Container
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
Left = Extender.Left
End Property

Public Property Let Left(ByVal Value As Single)
Extender.Left = Value
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns/sets the distance between the internal top edge of an object and the top edge of its container."
Top = Extender.Top
End Property

Public Property Let Top(ByVal Value As Single)
Extender.Top = Value
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

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value that determines whether an object is visible or hidden."
Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
Extender.Visible = Value
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
hWnd = OptionButtonHandle
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
OldFontHandle = OptionButtonFontHandle
OptionButtonFontHandle = CreateFontFromOLEFont(PropFont)
If OptionButtonHandle <> 0 Then SendMessage OptionButtonHandle, WM_SETFONT, OptionButtonFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = OptionButtonFontHandle
OptionButtonFontHandle = CreateFontFromOLEFont(PropFont)
If OptionButtonHandle <> 0 Then SendMessage OptionButtonHandle, WM_SETFONT, OptionButtonFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If OptionButtonHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles OptionButtonHandle
    Else
        RemoveVisualStyles OptionButtonHandle
    End If
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
If OptionButtonHandle <> 0 Then EnableWindow OptionButtonHandle, IIf(Value = True, 1, 0)
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

Public Property Get ImageList() As Variant
Attribute ImageList.VB_Description = "Returns/sets the image list control to be used. The image list should contain either a single image to be used for all states or individual images for each state. Requires comctl32.dll version 6.0 or higher."
If Ambient.UserMode = True Then
    If PropImageListInit = False And PropImageListControl Is Nothing Then
        If Not PropImageListName = "(None)" Then Me.ImageList = PropImageListName
        PropImageListInit = True
    End If
    Set ImageList = PropImageListControl
Else
    ImageList = PropImageListName
End If
End Property

Public Property Set ImageList(ByVal Value As Variant)
Me.ImageList = Value
End Property

Public Property Let ImageList(ByVal Value As Variant)
If OptionButtonHandle <> 0 Then
    ' The image list should contain either a single image to be used for all states or
    ' individual images for each state. The following states are defined as following:
    ' PBS_NORMAL = 1
    ' PBS_HOT = 2
    ' PBS_PRESSED = 3
    ' PBS_DISABLED = 4
    ' PBS_DEFAULTED = 5
    ' PBS_STYLUSHOT = 6
    Dim Success As Boolean, Handle As Long
    On Error Resume Next
    If IsObject(Value) Then
        If TypeName(Value) = "ImageList" Then
            Handle = Value.hImageList
            Success = CBool(Err.Number = 0 And Handle <> 0)
        End If
        If Success = True Then
            Call SetImageList(Handle)
            PropImageListName = ProperControlName(Value)
            If Ambient.UserMode = True Then Set PropImageListControl = Value
        End If
    ElseIf VarType(Value) = vbString Then
        Dim ControlEnum As Object, CompareName As String
        For Each ControlEnum In UserControl.ParentControls
            If TypeName(ControlEnum) = "ImageList" Then
                CompareName = ProperControlName(ControlEnum)
                If CompareName = Value And Not CompareName = vbNullString Then
                    Err.Clear
                    Handle = ControlEnum.hImageList
                    Success = CBool(Err.Number = 0 And Handle <> 0)
                    If Success = True Then
                        Call SetImageList(Handle)
                        PropImageListName = Value
                        If Ambient.UserMode = True Then Set PropImageListControl = ControlEnum
                        Exit For
                    ElseIf Ambient.UserMode = False Then
                        PropImageListName = Value
                        Success = True
                        Exit For
                    End If
                End If
            End If
        Next ControlEnum
    End If
    On Error GoTo 0
    If Success = False Then
        Call SetImageList(BCCL_NOGLYPH)
        PropImageListName = "(None)"
        Set PropImageListControl = Nothing
    End If
End If
UserControl.PropertyChanged "ImageList"
End Property

Public Property Get ImageListAlignment() As OptImageListAlignmentConstants
Attribute ImageListAlignment.VB_Description = "Returns/sets the alignment used to the image in the image list control. Requires comctl32.dll version 6.0 or higher."
ImageListAlignment = PropImageListAlignment
End Property

Public Property Let ImageListAlignment(ByVal Value As OptImageListAlignmentConstants)
Select Case Value
    Case OptImageListAlignmentLeft, OptImageListAlignmentRight, OptImageListAlignmentTop, OptImageListAlignmentBottom, OptImageListAlignmentCenter
        PropImageListAlignment = Value
    Case Else
        Err.Raise 380
End Select
If OptionButtonHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If Not PropImageListControl Is Nothing Then
        Me.ImageList = PropImageListControl
    ElseIf Not PropImageListName = "(None)" Then
        Me.ImageList = PropImageListName
    End If
End If
UserControl.PropertyChanged "ImageListAlignment"
End Property

Public Property Get ImageListMargin() As Single
Attribute ImageListMargin.VB_Description = "Returns/sets the margin (related to the alignment) used to the image in the image list control. Requires comctl32.dll version 6.0 or higher."
ImageListMargin = UserControl.ScaleX(PropImageListMargin, vbPixels, vbContainerSize)
End Property

Public Property Let ImageListMargin(ByVal Value As Single)
If Value < 0 Then
    If Ambient.UserMode = False Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropImageListMargin = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If OptionButtonHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If Not PropImageListControl Is Nothing Then
        Me.ImageList = PropImageListControl
    ElseIf Not PropImageListName = "(None)" Then
        Me.ImageList = PropImageListName
    End If
End If
UserControl.PropertyChanged "ImageListMargin"
End Property

Public Property Get Value() As OLE_OPTEXCLUSIVE
Attribute Value.VB_Description = "Returns/sets the value of an object."
Attribute Value.VB_UserMemId = 0
Value = PropValue
End Property

Public Property Let Value(ByVal NewValue As OLE_OPTEXCLUSIVE)
PropValue = NewValue
If OptionButtonHandle <> 0 Then SendMessage OptionButtonHandle, BM_SETCHECK, IIf(PropValue = True, BST_CHECKED, BST_UNCHECKED), ByVal 0&
UserControl.PropertyChanged "Value"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
If OptionButtonHandle <> 0 Then
    Caption = String(SendMessage(OptionButtonHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
    SendMessage OptionButtonHandle, WM_GETTEXT, Len(Caption) + 1, ByVal StrPtr(Caption)
Else
    Caption = PropCaption
End If
End Property

Public Property Let Caption(ByVal Value As String)
PropCaption = Value
UserControl.AccessKeys = ChrW(AccelCharCode(PropCaption))
If OptionButtonHandle <> 0 Then SendMessage OptionButtonHandle, WM_SETTEXT, 0, ByVal StrPtr(PropCaption)
UserControl.PropertyChanged "Caption"
End Property

Public Property Get Alignment() As CCLeftRightAlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment."
Alignment = PropAlignment
End Property

Public Property Let Alignment(ByVal Value As CCLeftRightAlignmentConstants)
Select Case Value
    Case CCLeftRightAlignmentLeft, CCLeftRightAlignmentRight
        PropAlignment = Value
    Case Else
        Err.Raise 380
End Select
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If PropAlignment = CCLeftRightAlignmentRight Then
        If Not (dwStyle And BS_RIGHTBUTTON) = BS_RIGHTBUTTON Then dwStyle = dwStyle Or BS_RIGHTBUTTON
    ElseIf PropAlignment = CCLeftRightAlignmentLeft Then
        If (dwStyle And BS_RIGHTBUTTON) = BS_RIGHTBUTTON Then dwStyle = dwStyle And Not BS_RIGHTBUTTON
    End If
    SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "Alignment"
End Property

Public Property Get TextAlignment() As VBRUN.AlignmentConstants
Attribute TextAlignment.VB_Description = "Returns/sets the alignment of the displayed text."
TextAlignment = PropTextAlignment
End Property

Public Property Let TextAlignment(ByVal Value As VBRUN.AlignmentConstants)
Select Case Value
    Case vbLeftJustify, vbCenter, vbRightJustify
        PropTextAlignment = Value
    Case Else
        Err.Raise 380
End Select
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If (dwStyle And BS_LEFT) = BS_LEFT Then dwStyle = dwStyle And Not BS_LEFT
    If (dwStyle And BS_CENTER) = BS_CENTER Then dwStyle = dwStyle And Not BS_CENTER
    If (dwStyle And BS_RIGHT) = BS_RIGHT Then dwStyle = dwStyle And Not BS_RIGHT
    Select Case PropTextAlignment
        Case vbLeftJustify
            dwStyle = dwStyle Or BS_LEFT
        Case vbCenter
            dwStyle = dwStyle Or BS_CENTER
        Case vbRightJustify
            dwStyle = dwStyle Or BS_RIGHT
    End Select
    SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "TextAlignment"
End Property

Public Property Get PushLike() As Boolean
Attribute PushLike.VB_Description = "Returns/sets a value that determines whether or not the control look and act like a push button."
PushLike = PropPushLike
End Property

Public Property Let PushLike(ByVal Value As Boolean)
PropPushLike = Value
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If PropPushLike = True Then
        If Not (dwStyle And BS_PUSHLIKE) = BS_PUSHLIKE Then dwStyle = dwStyle Or BS_PUSHLIKE
    Else
        If (dwStyle And BS_PUSHLIKE) = BS_PUSHLIKE Then dwStyle = dwStyle And Not BS_PUSHLIKE
    End If
    SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "PushLike"
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
Set Picture = PropPicture
End Property

Public Property Let Picture(ByVal Value As IPictureDisp)
Set Me.Picture = Value
End Property

Public Property Set Picture(ByVal Value As IPictureDisp)
Dim dwStyle As Long
If Value Is Nothing Then
    Set PropPicture = Nothing
    If OptionButtonHandle <> 0 Then
        dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
        If (dwStyle And BS_ICON) = BS_ICON Then dwStyle = dwStyle And Not BS_ICON
        If (dwStyle And BS_BITMAP) = BS_BITMAP Then dwStyle = dwStyle And Not BS_BITMAP
        SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
        SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
        SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
        Me.Refresh
    End If
Else
    Set UserControl.Picture = Value
    Set PropPicture = UserControl.Picture
    Set UserControl.Picture = Nothing
    If OptionButtonHandle <> 0 Then
        dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
        If (dwStyle And BS_ICON) = BS_ICON Then dwStyle = dwStyle And Not BS_ICON
        If (dwStyle And BS_BITMAP) = BS_BITMAP Then dwStyle = dwStyle And Not BS_BITMAP
        If PropPicture.Handle <> 0 Then
            If PropPicture.Type = vbPicTypeIcon Then
                dwStyle = dwStyle Or BS_ICON
                SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
                SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
                SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_ICON, ByVal PropPicture.Handle
            Else
                dwStyle = dwStyle Or BS_BITMAP
                SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
                SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
                SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal PropPicture.Handle
            End If
        Else
            SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
            SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
            SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
        End If
        Me.Refresh
    End If
End If
UserControl.PropertyChanged "Picture"
End Property

Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets a value that determines whether a control may break lines within the text in order to prevent overflow."
WordWrap = PropWordWrap
End Property

Public Property Let WordWrap(ByVal Value As Boolean)
PropWordWrap = Value
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If PropWordWrap = True Then
        If Not (dwStyle And BS_MULTILINE) = BS_MULTILINE Then dwStyle = dwStyle Or BS_MULTILINE
    Else
        If (dwStyle And BS_MULTILINE) = BS_MULTILINE Then dwStyle = dwStyle And Not BS_MULTILINE
    End If
    SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "WordWrap"
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
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If PropAppearance = CCAppearanceFlat Then
        If Not (dwStyle And BS_FLAT) = BS_FLAT Then dwStyle = dwStyle Or BS_FLAT
    Else
        If (dwStyle And BS_FLAT) = BS_FLAT Then dwStyle = dwStyle And Not BS_FLAT
    End If
    SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
End If
Me.Refresh
UserControl.PropertyChanged "Appearance"
End Property

Public Property Get VerticalAlignment() As CCVerticalAlignmentConstants
Attribute VerticalAlignment.VB_Description = "Returns/sets the vertical alignment."
VerticalAlignment = PropVerticalAlignment
End Property

Public Property Let VerticalAlignment(ByVal Value As CCVerticalAlignmentConstants)
Select Case Value
    Case CCVerticalAlignmentTop, CCVerticalAlignmentCenter, CCVerticalAlignmentBottom
        PropVerticalAlignment = Value
    Case Else
        Err.Raise 380
End Select
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If (dwStyle And BS_TOP) = BS_TOP Then dwStyle = dwStyle And Not BS_TOP
    If (dwStyle And BS_VCENTER) = BS_VCENTER Then dwStyle = dwStyle And Not BS_VCENTER
    If (dwStyle And BS_BOTTOM) = BS_BOTTOM Then dwStyle = dwStyle And Not BS_BOTTOM
    Select Case PropVerticalAlignment
        Case CCVerticalAlignmentTop
            dwStyle = dwStyle Or BS_TOP
        Case CCVerticalAlignmentCenter
            dwStyle = dwStyle Or BS_VCENTER
        Case CCVerticalAlignmentBottom
            dwStyle = dwStyle Or BS_BOTTOM
    End Select
    SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "VerticalAlignment"
End Property

Private Sub CreateOptionButton()
If OptionButtonHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or BS_RADIOBUTTON Or BS_TEXT Or BS_NOTIFY
If PropAlignment = CCLeftRightAlignmentRight Then dwStyle = dwStyle Or BS_RIGHTBUTTON
Select Case PropTextAlignment
    Case vbLeftJustify
        dwStyle = dwStyle Or BS_LEFT
    Case vbCenter
        dwStyle = dwStyle Or BS_CENTER
    Case vbRightJustify
        dwStyle = dwStyle Or BS_RIGHT
End Select
If PropPushLike = True Then dwStyle = dwStyle Or BS_PUSHLIKE
If PropWordWrap = True Then dwStyle = dwStyle Or BS_MULTILINE
If PropAppearance = CCAppearanceFlat Then dwStyle = dwStyle Or BS_FLAT
Select Case PropVerticalAlignment
    Case CCVerticalAlignmentTop
        dwStyle = dwStyle Or BS_TOP
    Case CCVerticalAlignmentCenter
        dwStyle = dwStyle Or BS_VCENTER
    Case CCVerticalAlignmentBottom
        dwStyle = dwStyle Or BS_BOTTOM
End Select
If Ambient.RightToLeft = True Then dwExStyle = WS_EX_RTLREADING
OptionButtonHandle = CreateWindowEx(dwExStyle, StrPtr("Button"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If OptionButtonHandle <> 0 Then Call ComCtlsShowAllUIStates(OptionButtonHandle)
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
If PropValue = True Then Me.Value = True
Me.Caption = PropCaption
If Not PropPicture Is Nothing Then Set Me.Picture = PropPicture
If Ambient.UserMode = True Then
    If OptionButtonHandle <> 0 Then Call ComCtlsSetSubclass(OptionButtonHandle, Me, 1)
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
End If
End Sub

Private Sub DestroyOptionButton()
If OptionButtonHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(OptionButtonHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow OptionButtonHandle, SW_HIDE
SetParent OptionButtonHandle, 0
DestroyWindow OptionButtonHandle
OptionButtonHandle = 0
If OptionButtonFontHandle <> 0 Then
    DeleteObject OptionButtonFontHandle
    OptionButtonFontHandle = 0
End If
If OptionButtonTransparentBrush <> 0 Then
    DeleteObject OptionButtonTransparentBrush
    OptionButtonTransparentBrush = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
If OptionButtonTransparentBrush <> 0 Then
    DeleteObject OptionButtonTransparentBrush
    OptionButtonTransparentBrush = 0
End If
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get Pushed() As Boolean
Attribute Pushed.VB_Description = "Returns/sets a value that indicates if the option button is in the pushed state."
Attribute Pushed.VB_MemberFlags = "400"
If OptionButtonHandle <> 0 Then Pushed = CBool((SendMessage(OptionButtonHandle, BM_GETSTATE, 0, ByVal 0&) And BST_PUSHED) = BST_PUSHED)
End Property

Public Property Let Pushed(ByVal Value As Boolean)
If OptionButtonHandle <> 0 Then SendMessage OptionButtonHandle, BM_SETSTATE, IIf(Value = True, 1, 0), ByVal 0&
End Property

Public Property Get Hot() As Boolean
Attribute Hot.VB_Description = "Returns/sets a value that indicates if the option button is hot; that is, the mouse is hovering over it. Requires comctl32.dll version 6.0 or higher."
Attribute Hot.VB_MemberFlags = "400"
If OptionButtonHandle <> 0 And ComCtlsSupportLevel() >= 1 Then Hot = CBool((SendMessage(OptionButtonHandle, BM_GETSTATE, 0, ByVal 0&) And BST_HOT) = BST_HOT)
End Property

Public Property Let Hot(ByVal Value As Boolean)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Private Sub SetImageList(ByVal hImageList As Long)
If OptionButtonHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim BTNIML As BUTTON_IMAGELIST
    With BTNIML
    .hImageList = hImageList
    If .hImageList <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
        If (dwStyle And BS_ICON) = BS_ICON Then dwStyle = dwStyle And Not BS_ICON
        If (dwStyle And BS_BITMAP) = BS_BITMAP Then dwStyle = dwStyle And Not BS_BITMAP
        SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
        SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
        SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
    End If
    With .RCMargin
    Select Case PropImageListAlignment
        Case OptImageListAlignmentLeft
            .Left = PropImageListMargin
        Case OptImageListAlignmentRight
            .Right = PropImageListMargin
        Case OptImageListAlignmentTop
            .Top = PropImageListMargin
        Case OptImageListAlignmentBottom
            .Bottom = PropImageListMargin
    End Select
    End With
    .uAlign = PropImageListAlignment
    SendMessage OptionButtonHandle, BCM_SETIMAGELIST, 0, ByVal VarPtr(BTNIML)
    If .hImageList = BCCL_NOGLYPH Then
        PropImageListName = "(None)"
        Set Me.Picture = PropPicture
    End If
    End With
    Me.Refresh
End If
End Sub

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
    Case WM_KEYDOWN, WM_KEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftState())
            OptionButtonCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_KEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftState())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If OptionButtonCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(OptionButtonCharCodeCache And &HFFFF&)
            OptionButtonCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(wParam And &HFFFF&)
        End If
        RaiseEvent KeyPress(KeyChar)
        wParam = CIntToUInt(KeyChar)
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then WindowProcControl = 1 Else SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_MOUSEACTIVATE
        Static InProc As Boolean
        If ComCtlsRootIsEditor(hWnd) = False And GetFocus() <> OptionButtonHandle Then
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
                            OptionButtonIgnoreClick = True
                            SetFocusAPI UserControl.hWnd
                            OptionButtonIgnoreClick = False
                            WindowProcControl = MA_NOACTIVATE
                        End If
                    Else
                        OptionButtonIgnoreClick = True
                        SetFocusAPI UserControl.hWnd
                        OptionButtonIgnoreClick = False
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

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        If lParam = OptionButtonHandle Then
            Select Case HiWord(wParam)
                Case BN_CLICKED
                    If OptionButtonIgnoreClick = True Then
                        Exit Function
                    ElseIf PropValue = False Then
                        Me.Value = True
                        RaiseEvent Click
                    End If
                Case BN_DOUBLECLICKED
                    RaiseEvent DblClick
            End Select
        End If
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = OptionButtonHandle Then
            Select Case NM.Code
                Case BCN_HOTITEMCHANGE
                    Dim NMBCHI As NMBCHOTITEM
                    CopyMemory NMBCHI, ByVal lParam, LenB(NMBCHI)
                    With NMBCHI
                    If (.dwFlags And HICF_MOUSE) = HICF_MOUSE Then
                        If (.dwFlags And HICF_ENTERING) = HICF_ENTERING Or (.dwFlags And HICF_LEAVING) = HICF_LEAVING Then RaiseEvent HotChanged
                    End If
                    End With
            End Select
        End If
    Case WM_CTLCOLORSTATIC, WM_CTLCOLORBTN
        WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        If PropTransparent = True Then
            SetBkMode wParam, 1
            Dim hDCScreen As Long, hDCBmp As Long
            Dim hBmp As Long, hBmpOld As Long
            With UserControl
            If OptionButtonTransparentBrush = 0 Then
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
                            OptionButtonTransparentBrush = CreatePatternBrush(hBmp)
                            SelectObject hDCBmp, hBmpOld
                            DeleteObject hBmp
                        End If
                        DeleteDC hDCBmp
                    End If
                    ReleaseDC 0, hDCScreen
                End If
            End If
            End With
            If OptionButtonTransparentBrush <> 0 Then WindowProcUserControl = OptionButtonTransparentBrush
        End If
        Exit Function
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS Then SetFocusAPI OptionButtonHandle
End Function
