VERSION 5.00
Begin VB.UserControl ProgressBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   2  'Dot
   HasDC           =   0   'False
   PropertyPages   =   "ProgressBar.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "ProgressBar.ctx":003D
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If False Then
Private PrbOrientationHorizontal, PrbOrientationVertical
Private PrbScrollingStandard, PrbScrollingSmooth
Private PrbStateInProgress, PrbStateError, PrbStatePaused
Private PrbTaskBarStateNone, PrbTaskBarStateMarquee, PrbTaskBarStateInProgress, PrbTaskBarStateError, PrbTaskBarStatePaused
#End If
Public Enum PrbOrientationConstants
PrbOrientationHorizontal = 0
PrbOrientationVertical = 1
End Enum
Public Enum PrbScrollingConstants
PrbScrollingStandard = 0
PrbScrollingSmooth = 1
End Enum
Private Const PBST_NORMAL As Long = 1
Private Const PBST_ERROR As Long = 2
Private Const PBST_PAUSED As Long = 3
Public Enum PrbStateConstants
PrbStateInProgress = PBST_NORMAL
PrbStateError = PBST_ERROR
PrbStatePaused = PBST_PAUSED
End Enum
Private Const TBPF_NOPROGRESS As Long = 0
Private Const TBPF_INDETERMINATE As Long = 1
Private Const TBPF_NORMAL As Long = 2
Private Const TBPF_ERROR As Long = 4
Private Const TBPF_PAUSED As Long = 8
Public Enum PrbTaskBarStateConstants
PrbTaskBarStateNone = TBPF_NOPROGRESS
PrbTaskBarStateMarquee = TBPF_INDETERMINATE
PrbTaskBarStateInProgress = TBPF_NORMAL
PrbTaskBarStateError = TBPF_ERROR
PrbTaskBarStatePaused = TBPF_PAUSED
End Enum
Private Enum VTableIndexITaskBarList3Constants
' Ignore : ITaskBarList3QueryInterface = 1
' Ignore : ITaskBarList3AddRef = 2
' Ignore : ITaskBarList3Release = 3
VTableIndexITaskBarList3HrInit = 4
' Ignore : ITaskBarList3AddTab = 5
' Ignore : ITaskBarList3DeleteTab = 6
' Ignore : ITaskBarList3ActivateTab = 7
' Ignore : ITaskBarList3SetActiveAlt = 8
' Ignore : ITaskBarList3MarkFullscreenWindow = 9
VTableIndexITaskBarList3SetProgressValue = 10
VTableIndexITaskBarList3SetProgressState = 11
' Ignore : ITaskBarList3RegisterTab = 12
' Ignore : ITaskBarList3UnregisterTab = 13
' Ignore : ITaskBarList3SetTabOrder = 14
' Ignore : ITaskBarList3SetTabActive = 15
' Ignore : ITaskBarList3ThumbBarAddButtons = 16
' Ignore : ITaskBarList3ThumbBarUpdateButtons = 17
' Ignore : ITaskBarList3ThumbBarSetImageList = 18
' Ignore : ITaskBarList3SetOverlayIcon = 19
' Ignore : ITaskBarList3SetThumbnailTooltip = 20
' Ignore : ITaskBarList3SetThumbnailClip = 21
End Enum
Private Type TagInitCommonControlsEx
dwSize As Long
dwICC As Long
End Type
Private Type PBRANGE
Min As Long
Max As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
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
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TagInitCommonControlsEx) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, ByRef pCLSID As Any) As Long
Private Declare Function CoCreateInstance Lib "ole32" (ByRef rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByRef riid As Any, ByRef ppv As IUnknown) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Const ICC_PROGRESS_CLASS As Long = &H20
Private Const CLSID_ITaskBarList As String = "{56FDF344-FD6D-11D0-958A-006097C9A090}"
Private Const IID_ITaskBarList3 As String = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"
Private Const CLSCTX_INPROC_SERVER As Long = 1, S_OK As Long = 0
Private Const RDW_UPDATENOW As Long = &H100
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4
Private Const RDW_ALLCHILDREN As Long = &H80
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const SW_HIDE As Long = &H0
Private Const SW_SHOW As Long = &H5
Private Const GA_ROOT As Long = 2
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETBKCOLOR As Long = (CCM_FIRST + 1)
Private Const WM_USER As Long = &H400
Private Const PBM_SETBKCOLOR As Long = CCM_SETBKCOLOR
Private Const PBM_SETRANGE As Long = (WM_USER + 1)
Private Const PBM_SETPOS As Long = (WM_USER + 2)
Private Const PBM_DELTAPOS As Long = (WM_USER + 3)
Private Const PBM_SETSTEP As Long = (WM_USER + 4)
Private Const PBM_STEPIT As Long = (WM_USER + 5)
Private Const PBM_SETRANGE32 As Long = (WM_USER + 6)
Private Const PBM_GETRANGE As Long = (WM_USER + 7)
Private Const PBM_GETPOS As Long = (WM_USER + 8)
Private Const PBM_SETBARCOLOR As Long = (WM_USER + 9)
Private Const PBM_SETMARQUEE As Long = (WM_USER + 10)
Private Const PBM_GETSTEP As Long = (WM_USER + 13)
Private Const PBM_SETSTATE As Long = (WM_USER + 16)
Private Const PBM_GETSTATE As Long = (WM_USER + 17)
Private Const PBS_SMOOTH As Long = &H1
Private Const PBS_VERTICAL As Long = &H4
Private Const PBS_MARQUEE As Long = &H8
Private Const PBS_SMOOTHREVERSE As Long = &H10
Implements ISubclass
Implements OLEGuids.IPerPropertyBrowsingVB
Private ProgressBarHandle As Long
Private ProgressBarITaskBarList3 As IUnknown
Private DispIDMousePointer As Long
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropRange As PBRANGE
Private PropValue As Long
Private PropStep As Integer, PropStepAutoReset As Boolean
Private PropMarquee As Boolean
Private PropMarqueeAnimation As Boolean, PropMarqueeSpeed As Long
Private PropOrientation As PrbOrientationConstants
Private PropScrolling As PrbScrollingConstants
Private PropSmoothReverse As Boolean
Private PropBackColor As OLE_COLOR, PropBarColor As OLE_COLOR
Private PropState As PrbStateConstants

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
.dwICC = ICC_PROGRESS_CLASS
End With
InitCommonControlsEx ICCEX
Call SetVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
DispIDMousePointer = GetDispID(Me, "MousePointer")
End Sub

Private Sub UserControl_InitProperties()
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropRange.Min = 0
PropRange.Max = 100
PropValue = 0
PropStep = 10
PropStepAutoReset = True
PropMarquee = False
PropMarqueeAnimation = False
PropMarqueeSpeed = 80
PropOrientation = PrbOrientationHorizontal
PropScrolling = PrbScrollingStandard
PropSmoothReverse = False
PropBackColor = vbButtonFace
PropBarColor = vbHighlight
PropState = PrbStateInProgress
Call CreateProgressBar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropRange.Min = .ReadProperty("Min", 0)
PropRange.Max = .ReadProperty("Max", 100)
PropValue = .ReadProperty("Value", 0)
PropStep = .ReadProperty("Step", 1)
PropStepAutoReset = .ReadProperty("StepAutoReset", True)
PropMarquee = .ReadProperty("Marquee", False)
PropMarqueeAnimation = .ReadProperty("MarqueeAnimation", False)
PropMarqueeSpeed = .ReadProperty("MarqueeSpeed", 80)
PropOrientation = .ReadProperty("Orientation", PrbOrientationHorizontal)
PropScrolling = .ReadProperty("Scrolling", PrbScrollingStandard)
PropSmoothReverse = .ReadProperty("SmoothReverse", PropSmoothReverse)
PropBackColor = .ReadProperty("BackColor", vbButtonFace)
PropBarColor = .ReadProperty("BarColor", vbHighlight)
PropState = .ReadProperty("State", PrbStateInProgress)
End With
Call CreateProgressBar
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "Min", PropRange.Min, 0
.WriteProperty "Max", PropRange.Max, 100
.WriteProperty "Value", PropValue, 0
.WriteProperty "Step", PropStep, 1
.WriteProperty "StepAutoReset", PropStepAutoReset, True
.WriteProperty "Marquee", PropMarquee, False
.WriteProperty "MarqueeAnimation", PropMarqueeAnimation, False
.WriteProperty "MarqueeSpeed", PropMarqueeSpeed, 80
.WriteProperty "Orientation", PropOrientation, PrbOrientationHorizontal
.WriteProperty "Scrolling", PropScrolling, PrbScrollingStandard
.WriteProperty "SmoothReverse", PropSmoothReverse, False
.WriteProperty "BackColor", PropBackColor, vbButtonFace
.WriteProperty "BarColor", PropBarColor, vbHighlight
.WriteProperty "State", PropState, PrbStateInProgress
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
Static LastHeight As Single, LastWidth As Single, LastAlign As AlignConstants
Static InProc As Boolean
If ProgressBarHandle = 0 Or InProc = True Then Exit Sub
InProc = True
With Extender
Select Case .Align
    Case LastAlign
    Case vbAlignNone
    Case vbAlignTop, vbAlignBottom
        Select Case LastAlign
            Case vbAlignLeft, vbAlignRight
                .Height = LastWidth
        End Select
    Case vbAlignLeft, vbAlignRight
        Select Case LastAlign
            Case vbAlignTop, vbAlignBottom
                .Width = LastHeight
        End Select
End Select
LastHeight = .Height
LastWidth = .Width
LastAlign = .Align
End With
MoveWindow ProgressBarHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyProgressBar
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
hWnd = ProgressBarHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If ProgressBarHandle <> 0 And EnabledVisualStyles() = True Then
    Select Case PropVisualStyles
        Case True
            ActivateVisualStyles ProgressBarHandle
        Case False
            RemoveVisualStyles ProgressBarHandle
    End Select
    Me.Refresh
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
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

Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets the minimum position."
If ProgressBarHandle <> 0 Then
    Min = SendMessageLong(ProgressBarHandle, PBM_GETRANGE, 1, 0)
Else
    Min = PropRange.Min
End If
End Property

Public Property Let Min(ByVal Value As Long)
If Value < PropRange.Max Then
    PropRange.Min = Value
    If PropValue < PropRange.Min Then PropValue = PropRange.Min
Else
    If Ambient.UserMode = False Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If ProgressBarHandle <> 0 Then SendMessageLong ProgressBarHandle, PBM_SETRANGE32, PropRange.Min, PropRange.Max
UserControl.PropertyChanged "Min"
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum position."
If ProgressBarHandle = 0 Then
    Max = SendMessageLong(ProgressBarHandle, PBM_GETRANGE, 0, 0)
Else
    Max = PropRange.Max
End If
End Property

Public Property Let Max(ByVal Value As Long)
If Value > PropRange.Min Then
    PropRange.Max = Value
    If PropValue > PropRange.Max Then PropValue = PropRange.Max
Else
    If Ambient.UserMode = False Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If ProgressBarHandle <> 0 Then SendMessageLong ProgressBarHandle, PBM_SETRANGE32, PropRange.Min, PropRange.Max
UserControl.PropertyChanged "Max"
End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the current position."
Attribute Value.VB_UserMemId = 0
If ProgressBarHandle <> 0 Then
    Value = SendMessageLong(ProgressBarHandle, PBM_GETPOS, 0, 0)
Else
    Value = PropValue
End If
End Property

Public Property Let Value(ByVal NewValue As Long)
PropValue = NewValue
Select Case PropValue
    Case PropRange.Min To PropRange.Max
        If ProgressBarHandle <> 0 Then SendMessageLong ProgressBarHandle, PBM_SETPOS, PropValue, 0
    Case Is < PropRange.Min
        Me.Value = PropRange.Min
    Case Is > PropRange.Max
        Me.Value = PropRange.Max
End Select
UserControl.PropertyChanged "Value"
End Property

Public Property Get Step() As Long
Attribute Step.VB_Description = "Returns/sets the step value for the 'StepIt' procedure."
If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Step = SendMessageLong(ProgressBarHandle, PBM_GETSTEP, 0, 0)
Else
    Step = PropStep
End If
End Property

Public Property Let Step(ByVal Value As Long)
PropStep = Value
If ProgressBarHandle <> 0 Then SendMessageLong ProgressBarHandle, PBM_SETSTEP, PropStep, 0
UserControl.PropertyChanged "Step"
End Property

Public Property Get StepAutoReset() As Boolean
Attribute StepAutoReset.VB_Description = "Returns/sets a value that determines whether the position will be automatically reset when the maximum is exceeded or not. Only applicable for the 'StepIt' procedure."
StepAutoReset = PropStepAutoReset
End Property

Public Property Let StepAutoReset(ByVal Value As Boolean)
PropStepAutoReset = Value
UserControl.PropertyChanged "StepAutoReset"
End Property

Public Property Get Marquee() As Boolean
Attribute Marquee.VB_Description = "Returns/sets a value that determines whether the marquee style is enabled or not. Requires comctl32.dll version 6.0 or higher."
Marquee = PropMarquee
End Property

Public Property Let Marquee(ByVal Value As Boolean)
PropMarquee = Value
If ProgressBarHandle <> 0 Then Call ReCreateProgressBar
UserControl.PropertyChanged "Marquee"
End Property

Public Property Get MarqueeAnimation() As Boolean
Attribute MarqueeAnimation.VB_Description = "Returns/sets a value that determines whether the marquee animation is on or off. Requires comctl32.dll version 6.0 or higher."
MarqueeAnimation = PropMarqueeAnimation
End Property

Public Property Let MarqueeAnimation(ByVal Value As Boolean)
PropMarqueeAnimation = Value
If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If PropMarqueeAnimation = True Then
        SendMessageLong ProgressBarHandle, PBM_SETMARQUEE, 1, PropMarqueeSpeed
    Else
        SendMessageLong ProgressBarHandle, PBM_SETMARQUEE, 0, PropMarqueeSpeed
    End If
End If
UserControl.PropertyChanged "Marquee"
End Property

Public Property Get MarqueeSpeed() As Long
Attribute MarqueeSpeed.VB_Description = "Returns/sets the speed of the marquee animation. That means the time, in milliseconds, between marquee animation updates. Requires comctl32.dll version 6.0 or higher."
MarqueeSpeed = PropMarqueeSpeed
End Property

Public Property Let MarqueeSpeed(ByVal Value As Long)
PropMarqueeSpeed = Value
If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If PropMarquee = True Then SendMessageLong ProgressBarHandle, PBM_SETMARQUEE, 1, PropMarqueeSpeed
End If
UserControl.PropertyChanged "MarqueeSpeed"
End Property

Public Property Get Orientation() As PrbOrientationConstants
Attribute Orientation.VB_Description = "Returns/sets the orientation."
Orientation = PropOrientation
End Property

Public Property Let Orientation(ByVal Value As PrbOrientationConstants)
Select Case Value
    Case PrbOrientationHorizontal, PrbOrientationVertical
        If Extender.Align = vbAlignNone Then
            Dim Swap As Single
            If PropOrientation = PrbOrientationHorizontal And Value = PrbOrientationVertical Then
                Swap = UserControl.Width
                UserControl.Width = UserControl.Height
                UserControl.Height = Swap
            ElseIf PropOrientation = PrbOrientationVertical And Value = PrbOrientationHorizontal Then
                Swap = UserControl.Height
                UserControl.Height = UserControl.Width
                UserControl.Width = Swap
            End If
        End If
        PropOrientation = Value
    Case Else
        Err.Raise 380
End Select
If ProgressBarHandle <> 0 Then Call ReCreateProgressBar
UserControl.PropertyChanged "Orientation"
End Property

Public Property Get Scrolling() As PrbScrollingConstants
Attribute Scrolling.VB_Description = "Returns/sets the scrolling."
Scrolling = PropScrolling
End Property

Public Property Let Scrolling(ByVal Value As PrbScrollingConstants)
Select Case Value
    Case PrbScrollingStandard, PrbScrollingSmooth
        PropScrolling = Value
    Case Else
        Err.Raise 380
End Select
If ProgressBarHandle <> 0 Then Call ReCreateProgressBar
UserControl.PropertyChanged "Scrolling"
End Property

Public Property Get SmoothReverse() As Boolean
Attribute SmoothReverse.VB_Description = "Returns/sets a value that determines the animation behavior when moving backward. If this is set, then a smooth transition will occur, otherwise it will jump to the lower value. Requires comctl32.dll version 6.1 or higher."
SmoothReverse = PropSmoothReverse
End Property

Public Property Let SmoothReverse(ByVal Value As Boolean)
PropSmoothReverse = Value
If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 1 Then Call ReCreateProgressBar
UserControl.PropertyChanged "SmoothReverse"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object. Only applicable if visual styles are not used."
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If ProgressBarHandle <> 0 Then
    SendMessageLong ProgressBarHandle, PBM_SETBKCOLOR, 0, WinColor(PropBackColor)
End If
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get BarColor() As OLE_COLOR
Attribute BarColor.VB_Description = "Returns/sets the bar color used to display text and graphics in an object. Only applicable if visual styles are not used."
BarColor = PropBarColor
End Property

Public Property Let BarColor(ByVal Value As OLE_COLOR)
PropBarColor = Value
If ProgressBarHandle <> 0 Then
    SendMessageLong ProgressBarHandle, PBM_SETBARCOLOR, 0, WinColor(PropBarColor)
End If
UserControl.PropertyChanged "BarColor"
End Property

Public Property Get State() As PrbStateConstants
Attribute State.VB_Description = "Returns/sets the state of the progress bar. Requires comctl32.dll version 6.1 or higher."
If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    State = SendMessageLong(ProgressBarHandle, PBM_GETSTATE, 0, 0)
Else
    State = PropState
End If
End Property

Public Property Let State(ByVal Value As PrbStateConstants)
PropState = Value
If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    SendMessageLong ProgressBarHandle, PBM_SETSTATE, PropState, 0
End If
UserControl.PropertyChanged "State"
End Property

Private Sub CreateProgressBar()
If ProgressBarHandle <> 0 Then Exit Sub
Dim dwStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
If PropOrientation = PrbOrientationVertical Then dwStyle = dwStyle Or PBS_VERTICAL
If PropScrolling = PrbScrollingSmooth Then dwStyle = dwStyle Or PBS_SMOOTH
If ComCtlsSupportLevel() >= 1 Then
    If PropMarquee = True Then dwStyle = dwStyle Or PBS_MARQUEE
    If PropSmoothReverse = True Then dwStyle = dwStyle Or PBS_SMOOTHREVERSE
End If
ProgressBarHandle = CreateWindowEx(0, StrPtr("msctls_progress32"), StrPtr("Progress Bar"), dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If ProgressBarHandle <> 0 Then SendMessageLong ProgressBarHandle, PBM_SETRANGE32, PropRange.Min, PropRange.Max
Me.VisualStyles = PropVisualStyles
Me.Value = PropValue
Me.Step = PropStep
Me.MarqueeAnimation = PropMarqueeAnimation
Me.BackColor = PropBackColor
Me.BarColor = PropBarColor
Me.State = PropState
If Ambient.UserMode = True Then
    If ProgressBarHandle <> 0 Then Call ComCtlsSetSubclass(ProgressBarHandle, Me, 0)
End If
End Sub

Private Sub ReCreateProgressBar()
If Ambient.UserMode = True Then
    SendMessageLong UserControl.hWnd, WM_SETREDRAW, 0, 0
    Call DestroyProgressBar
    Call CreateProgressBar
    Call UserControl_Resize
    SendMessageLong UserControl.hWnd, WM_SETREDRAW, 1, 0
    Me.Refresh
Else
    Call DestroyProgressBar
    Call CreateProgressBar
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyProgressBar()
If ProgressBarHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(ProgressBarHandle, Me)
ShowWindow ProgressBarHandle, SW_HIDE
SetParent ProgressBarHandle, 0
DestroyWindow ProgressBarHandle
ProgressBarHandle = 0
End Sub

Public Sub StepIt()
Attribute StepIt.VB_Description = "Advances the current position by the step increment."
If ProgressBarHandle = 0 Then Exit Sub
If PropStepAutoReset = True Then
    SendMessageLong ProgressBarHandle, PBM_STEPIT, 0, 0
    PropValue = Me.Value
Else
    If Me.Value + Me.Step <= Me.Max Then
        SendMessageLong ProgressBarHandle, PBM_STEPIT, 0, 0
        PropValue = Me.Value
    Else
        Me.Value = Me.Max
    End If
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Sub SetTaskBarProgressState(ByVal State As PrbTaskBarStateConstants)
Attribute SetTaskBarProgressState.VB_Description = "Sets the type and state of the progress indicator displayed on a taskbar button. This is only applicable on windows Vista/7 or above."
If ProgressBarHandle <> 0 Then
    If ProgressBarITaskBarList3 Is Nothing Then Set ProgressBarITaskBarList3 = CreateITaskBarList3()
    If Not ProgressBarITaskBarList3 Is Nothing Then
        Dim hWnd As Long
        hWnd = GetAncestor(ProgressBarHandle, GA_ROOT)
        If hWnd <> 0 Then VTableCall vbLong, ProgressBarITaskBarList3, VTableIndexITaskBarList3SetProgressState, hWnd, State
    End If
End If
End Sub

Public Sub SetTaskBarProgressValue(ByVal Completed As Long, ByVal Total As Long)
Attribute SetTaskBarProgressValue.VB_Description = "Displays or updates a progress bar hosted in a taskbar button to show the specific percentage completed of the full operation. This is only applicable on windows Vista/7 or above."
If ProgressBarHandle <> 0 Then
    If ProgressBarITaskBarList3 Is Nothing Then Set ProgressBarITaskBarList3 = CreateITaskBarList3()
    If Not ProgressBarITaskBarList3 Is Nothing Then
        Dim hWnd As Long
        hWnd = GetAncestor(ProgressBarHandle, GA_ROOT)
        If hWnd <> 0 Then VTableCall vbLong, ProgressBarITaskBarList3, VTableIndexITaskBarList3SetProgressValue, hWnd, Completed, 0&, Total, 0&
    End If
End If
End Sub

Private Function CreateITaskBarList3() As IUnknown
Dim CLSID As OLEGuids.OLECLSID, IID As OLEGuids.OLECLSID
On Error Resume Next
CLSIDFromString StrPtr(CLSID_ITaskBarList), CLSID
CLSIDFromString StrPtr(IID_ITaskBarList3), IID
CoCreateInstance CLSID, 0, CLSCTX_INPROC_SERVER, IID, CreateITaskBarList3
If Not CreateITaskBarList3 Is Nothing Then If VTableCall(vbLong, CreateITaskBarList3, VTableIndexITaskBarList3HrInit) <> S_OK Then Set CreateITaskBarList3 = Nothing
End Function

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
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
                Dim P As POINTAPI
                GetCursorPos P
                If WindowFromPoint(P.X, P.Y) = hWnd Then RaiseEvent Click
        End Select
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function
