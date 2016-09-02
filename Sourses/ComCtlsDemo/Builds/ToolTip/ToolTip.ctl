VERSION 5.00
Begin VB.UserControl ToolTip 
   BackColor       =   &H80000018&
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "ToolTip.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "ToolTip.ctx":0036
End
Attribute VB_Name = "ToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If False Then
Private TipDelayTimeReshow, TipDelayTimeShow, TipDelayTimeInitial
Private TipIconNone, TipIconInfo, TipIconWarning, TipIconError
#End If
Private Const TTDT_RESHOW As Long = 1
Private Const TTDT_AUTOPOP As Long = 2
Private Const TTDT_INITIAL As Long = 3
Public Enum TipDelayTimeConstants
TipDelayTimeReshow = TTDT_RESHOW
TipDelayTimeShow = TTDT_AUTOPOP
TipDelayTimeInitial = TTDT_INITIAL
End Enum
Private Const TTI_NONE As Long = 0
Private Const TTI_INFO As Long = 1
Private Const TTI_WARNING As Long = 2
Private Const TTI_ERROR As Long = 3
Public Enum TipIconConstants
TipIconNone = TTI_NONE
TipIconInfo = TTI_INFO
TipIconWarning = TTI_WARNING
TipIconError = TTI_ERROR
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type TOOLINFO
cbSize As Long
uFlags As Long
hWnd As Long
uId As Long
RECT As RECT
hInst As Long
lpszText As Long
lParam As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Type NMTTDISPINFO
hdr As NMHDR
lpszText As Long
szText(0 To ((80 * 2) - 1)) As Byte
hInst As Long
uFlags As Long
lParam As Long
End Type
Public Event Show(ByVal Tool As TipTool)
Attribute Show.VB_Description = "Occurs when a tool tip is about to be displayed."
Public Event Hide(ByVal Tool As TipTool)
Attribute Hide.VB_Description = "Occurs when a tool tip is about to be hidden."
Public Event LinkClick(ByVal Tool As TipTool)
Attribute LinkClick.VB_Description = "Occurs when clicking on a text link inside a tool tip. This will only occur if the tool tip is tracked. Requires comctl32.dll version 6.1 or higher."
Public Event NeedText(ByVal Tool As TipTool, ByRef Text As String)
Attribute NeedText.VB_Description = "Occurs when a tool tip has no text. Use this event to assign a text dynamically to a tool tip."
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const ICC_TAB_CLASSES As Long = &H8
Private Const GWL_STYLE As Long = (-16)
Private Const WS_POPUP As Long = &H80000000
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_RTLREADING As Long = &H2000
Private Const GA_ROOT As Long = 2
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_NOTIFYFORMAT As Long = &H55
Private Const WM_SHOWWINDOW As Long = &H18
Private Const WM_SETFONT As Long = &H30
Private Const WM_USER As Long = &H400
Private Const TTM_ACTIVATE As Long = (WM_USER + 1)
Private Const TTM_SETDELAYTIME As Long = (WM_USER + 3)
Private Const TTM_RELAYEVENT As Long = (WM_USER + 7)
Private Const TTM_GETTOOLCOUNT As Long = (WM_USER + 13)
Private Const TTM_WINDOWFROMPOINT As Long = (WM_USER + 16)
Private Const TTM_TRACKACTIVATE As Long = (WM_USER + 17)
Private Const TTM_TRACKPOSITION As Long = (WM_USER + 18)
Private Const TTM_SETTIPBKCOLOR As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR As Long = (WM_USER + 20)
Private Const TTM_GETDELAYTIME As Long = (WM_USER + 21)
Private Const TTM_GETTIPBKCOLOR As Long = (WM_USER + 22)
Private Const TTM_GETTIPTEXTCOLOR As Long = (WM_USER + 23)
Private Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)
Private Const TTM_GETMAXTIPWIDTH As Long = (WM_USER + 25)
Private Const TTM_SETMARGIN As Long = (WM_USER + 26)
Private Const TTM_GETMARGIN As Long = (WM_USER + 27)
Private Const TTM_POP As Long = (WM_USER + 28)
Private Const TTM_UPDATE As Long = (WM_USER + 29)
Private Const TTM_GETTITLE As Long = (WM_USER + 35)
Private Const TTM_ADDTOOLA As Long = (WM_USER + 4)
Private Const TTM_ADDTOOLW As Long = (WM_USER + 50)
Private Const TTM_ADDTOOL As Long = TTM_ADDTOOLW
Private Const TTM_DELTOOLA As Long = (WM_USER + 5)
Private Const TTM_DELTOOLW As Long = (WM_USER + 51)
Private Const TTM_DELTOOL As Long = TTM_DELTOOLW
Private Const TTM_NEWTOOLRECTA As Long = (WM_USER + 6)
Private Const TTM_NEWTOOLRECTW As Long = (WM_USER + 52)
Private Const TTM_NEWTOOLRECT As Long = TTM_NEWTOOLRECTW
Private Const TTM_GETTOOLINFOA As Long = (WM_USER + 8)
Private Const TTM_GETTOOLINFOW As Long = (WM_USER + 53)
Private Const TTM_GETTOOLINFO As Long = TTM_GETTOOLINFOW
Private Const TTM_SETTOOLINFOA As Long = (WM_USER + 9)
Private Const TTM_SETTOOLINFOW As Long = (WM_USER + 54)
Private Const TTM_SETTOOLINFO As Long = TTM_SETTOOLINFOW
Private Const TTM_GETTEXTA As Long = (WM_USER + 11)
Private Const TTM_GETTEXTW As Long = (WM_USER + 56)
Private Const TTM_GETTEXT As Long = TTM_GETTEXTW
Private Const TTM_UPDATETIPTEXTA As Long = (WM_USER + 12)
Private Const TTM_UPDATETIPTEXTW As Long = (WM_USER + 57)
Private Const TTM_UPDATETIPTEXT As Long = TTM_UPDATETIPTEXTW
Private Const TTM_ENUMTOOLSA As Long = (WM_USER + 14)
Private Const TTM_ENUMTOOLSW As Long = (WM_USER + 58)
Private Const TTM_ENUMTOOLS As Long = TTM_ENUMTOOLSW
Private Const TTM_GETCURRENTTOOLA As Long = (WM_USER + 15)
Private Const TTM_GETCURRENTTOOLW As Long = (WM_USER + 59)
Private Const TTM_GETCURRENTTOOL As Long = TTM_GETCURRENTTOOLW
Private Const TTM_SETTITLEA As Long = (WM_USER + 32)
Private Const TTM_SETTITLEW As Long = (WM_USER + 33)
Private Const TTM_SETTITLE As Long = TTM_SETTITLEW
Private Const LPSTR_TEXTCALLBACK As Long = (-1)
Private Const TTF_IDISHWND As Long = &H1
Private Const TTF_CENTERTIP As Long = &H2
Private Const TTF_RTLREADING As Long = &H4
Private Const TTF_SUBCLASS As Long = &H10
Private Const TTF_TRACK As Long = &H20
Private Const TTF_ABSOLUTE As Long = &H80
Private Const TTF_TRANSPARENT As Long = &H100
Private Const TTF_PARSELINKS As Long = &H1000
Private Const TTF_DI_SETITEM As Long = &H8000&
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_NOPREFIX As Long = &H2
Private Const TTS_NOANIMATE As Long = &H10
Private Const TTS_NOFADE As Long = &H20
Private Const TTS_BALLOON As Long = &H40
Private Const TTS_CLOSE As Long = &H80
Private Const TTS_USEVISUALSTYLE As Long = &H100 ' Unusable
Private Const TTN_FIRST As Long = (-520)
Private Const TTN_GETDISPINFOA As Long = (TTN_FIRST - 0)
Private Const TTN_GETDISPINFOW As Long = (TTN_FIRST - 10)
Private Const TTN_GETDISPINFO As Long = TTN_GETDISPINFOW
Private Const TTN_SHOW As Long = (TTN_FIRST - 1)
Private Const TTN_POP As Long = (TTN_FIRST - 2)
Private Const TTN_LINKCLICK As Long = (TTN_FIRST - 3)
Implements ISubclass
Private ToolTipHandle As Long
Private ToolTipMaxTipLength As Long
Private ToolTipFontHandle As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropTools As TipTools
Private PropVisualStyles As Boolean
Private PropBackColor As OLE_COLOR
Private PropForeColor As OLE_COLOR
Private PropMaxTipWidth As Long
Private PropTitle As String
Private PropIcon As TipIconConstants
Private PropBalloon As Boolean
Private PropCloseButton As Boolean
Private PropFadeAnimation As Boolean
Private PropUseSystemFont As Boolean

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_TAB_CLASSES)
End Sub

Private Sub UserControl_InitProperties()
Set PropFont = Ambient.Font
PropVisualStyles = True
PropBackColor = vbInfoBackground
PropForeColor = vbInfoText
PropMaxTipWidth = -1
PropTitle = vbNullString
PropIcon = TipIconNone
PropBalloon = False
PropCloseButton = False
PropFadeAnimation = True
PropUseSystemFont = True
If Ambient.UserMode = True Then Call CreateToolTip
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
Set PropFont = .ReadProperty("Font", Ambient.Font)
PropVisualStyles = .ReadProperty("VisualStyles", True)
PropBackColor = .ReadProperty("BackColor", vbInfoBackground)
PropForeColor = .ReadProperty("ForeColor", vbInfoText)
Me.Enabled = .ReadProperty("Enabled", True)
PropMaxTipWidth = .ReadProperty("MaxTipWidth", -1)
PropTitle = VarToStr(.ReadProperty("Title", vbNullString))
PropIcon = .ReadProperty("Icon", TipIconNone)
PropBalloon = .ReadProperty("Balloon", False)
PropCloseButton = .ReadProperty("CloseButton", False)
PropFadeAnimation = .ReadProperty("FadeAnimation", True)
PropUseSystemFont = .ReadProperty("UseSystemFont", True)
End With
If Ambient.UserMode = True Then Call CreateToolTip
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", PropFont, Ambient.Font
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", PropBackColor, vbInfoBackground
.WriteProperty "ForeColor", PropForeColor, vbInfoText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "MaxTipWidth", PropMaxTipWidth, -1
.WriteProperty "Title", StrToVar(PropTitle), vbNullString
.WriteProperty "Icon", PropIcon, TipIconNone
.WriteProperty "Balloon", PropBalloon, False
.WriteProperty "CloseButton", PropCloseButton, False
.WriteProperty "FadeAnimation", PropFadeAnimation, True
.WriteProperty "UseSystemFont", PropUseSystemFont, True
End With
End Sub

Private Sub UserControl_Paint()
If Ambient.UserMode = False Then
    With UserControl
    UserControl.Line (.ScaleWidth \ 2, 0)-Step(0, .ScaleHeight), vbBlue
    UserControl.Line (0, .ScaleHeight \ 2)-Step(.ScaleWidth, 0), vbBlue
    UserControl.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vbInfoText, B
    End With
End If
End Sub

Private Sub UserControl_Resize()
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
If Ambient.UserMode = False Then
    With UserControl
    .Size .ScaleX(17, vbPixels, vbTwips), .ScaleY(17, vbPixels, vbTwips)
    End With
End If
InProc = False
End Sub

Private Sub UserControl_Hide()
If Not PropTools Is Nothing Then
    On Error Resume Next
    If UserControl.Parent Is Nothing Then Set PropTools = Nothing
    On Error GoTo 0
End If
End Sub

Private Sub UserControl_Terminate()
Call DestroyToolTip
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

Public Property Get hToolTip() As Long
Attribute hToolTip.VB_Description = "Returns a handle to an tool tip control."
hToolTip = ToolTipHandle
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
If NewFont Is Nothing Then Set NewFont = Ambient.Font
Set PropFont = NewFont
OldFontHandle = ToolTipFontHandle
ToolTipFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ToolTipHandle <> 0 Then
    If PropUseSystemFont = False Then
        SendMessage ToolTipHandle, WM_SETFONT, ToolTipFontHandle, ByVal 1&
    Else
        SendMessage ToolTipHandle, WM_SETFONT, 0, ByVal 1&
    End If
End If
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
If Ambient.UserMode = False Then Set UserControl.Font = PropFont
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = ToolTipFontHandle
ToolTipFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ToolTipHandle <> 0 Then
    If PropUseSystemFont = False Then
        SendMessage ToolTipHandle, WM_SETFONT, ToolTipFontHandle, ByVal 1&
    Else
        SendMessage ToolTipHandle, WM_SETFONT, 0, ByVal 1&
    End If
End If
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
If Ambient.UserMode = False Then Set UserControl.Font = PropFont
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If ToolTipHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles ToolTipHandle
    Else
        RemoveVisualStyles ToolTipHandle
    End If
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object. This property is ignored if the version of comctl32.dll is 6.1 or higher and the visual styles property is set to true."
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If ToolTipHandle <> 0 Then SendMessage ToolTipHandle, TTM_SETTIPBKCOLOR, WinColor(PropBackColor), ByVal 0&
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object. This property is ignored if the version of comctl32.dll is 6.1 or higher and the visual styles property is set to true."
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
If ToolTipHandle <> 0 Then SendMessage ToolTipHandle, TTM_SETTIPTEXTCOLOR, WinColor(PropForeColor), ByVal 0&
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If ToolTipHandle <> 0 Then SendMessage ToolTipHandle, TTM_ACTIVATE, IIf(Value = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get MaxTipWidth() As Single
Attribute MaxTipWidth.VB_Description = "Returns/sets the maximum width for a tool tip window. A value of -1 or 0 indicates that any width is allowed. When set to -1 the text will be always single lined."
Dim LngValue As Long
If ToolTipHandle <> 0 Then
    LngValue = SendMessage(ToolTipHandle, TTM_GETMAXTIPWIDTH, 0, ByVal 0&)
Else
    LngValue = PropMaxTipWidth
End If
If LngValue = -1 Then
    MaxTipWidth = -1
Else
    MaxTipWidth = UserControl.ScaleX(LngValue, vbPixels, vbContainerSize)
End If
End Property

Public Property Let MaxTipWidth(ByVal Value As Single)
If Value < 0 And Not Value = -1 Then
    If Ambient.UserMode = False Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
Dim LngValue As Long, ErrValue As Long
If Value = -1 Then
    LngValue = -1
Else
    On Error Resume Next
    LngValue = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
    ErrValue = Err.Number
    On Error GoTo 0
End If
If (LngValue >= 0 Or LngValue = -1) And ErrValue = 0 Then
    PropMaxTipWidth = LngValue
    If ToolTipHandle <> 0 Then SendMessage ToolTipHandle, TTM_SETMAXTIPWIDTH, 0, ByVal PropMaxTipWidth
Else
    If Ambient.UserMode = False Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
UserControl.PropertyChanged "MaxTipWidth"
End Property

Public Property Get Title() As String
Attribute Title.VB_Description = "Returns/sets the title."
Title = PropTitle
End Property

Public Property Let Title(ByVal Value As String)
PropTitle = Value
If ToolTipHandle <> 0 Then SendMessage ToolTipHandle, TTM_SETTITLE, PropIcon, ByVal StrPtr(PropTitle)
UserControl.PropertyChanged "Title"
End Property

Public Property Get Icon() As TipIconConstants
Attribute Icon.VB_Description = "Returns/sets a value specifying a standard icon to be displayed. Only applicable if the title property is set."
Icon = PropIcon
End Property

Public Property Let Icon(ByVal Value As TipIconConstants)
Select Case Value
    Case TipIconNone, TipIconInfo, TipIconWarning, TipIconError
        PropIcon = Value
    Case Else
        Err.Raise 380
End Select
If ToolTipHandle <> 0 Then SendMessage ToolTipHandle, TTM_SETTITLE, PropIcon, ByVal StrPtr(PropTitle)
UserControl.PropertyChanged "Icon"
End Property

Public Property Get Balloon() As Boolean
Attribute Balloon.VB_Description = "Returns/sets a value that indicates if the tool tip control has the appearance of a cartoon balloon, with rounded corners and a stem pointing to the item."
Balloon = PropBalloon
End Property

Public Property Let Balloon(ByVal Value As Boolean)
PropBalloon = Value
If ToolTipHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ToolTipHandle, GWL_STYLE)
    If PropBalloon = True Then
        If Not (dwStyle And TTS_BALLOON) = TTS_BALLOON Then dwStyle = dwStyle Or TTS_BALLOON
    Else
        If (dwStyle And TTS_BALLOON) = TTS_BALLOON Then dwStyle = dwStyle And Not TTS_BALLOON
    End If
    SetWindowLong ToolTipHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "Balloon"
End Property

Public Property Get CloseButton() As Boolean
Attribute CloseButton.VB_Description = "Returns/sets a value indicating if the tool tip control displays a close button. Only applicable when the tool tip has a title and the balloon property is set to true. Requires comctl32.dll version 6.0 or higher."
CloseButton = PropCloseButton
End Property

Public Property Let CloseButton(ByVal Value As Boolean)
PropCloseButton = Value
If ToolTipHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ToolTipHandle, GWL_STYLE)
    If PropCloseButton = True Then
        If Not (dwStyle And TTS_CLOSE) = TTS_CLOSE Then dwStyle = dwStyle Or TTS_CLOSE
    Else
        If (dwStyle And TTS_CLOSE) = TTS_CLOSE Then dwStyle = dwStyle And Not TTS_CLOSE
    End If
    SetWindowLong ToolTipHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "CloseButton"
End Property

Public Property Get FadeAnimation() As Boolean
Attribute FadeAnimation.VB_Description = "Returns/sets a value that indicates if the fading animation is enabled or not."
FadeAnimation = PropFadeAnimation
End Property

Public Property Let FadeAnimation(ByVal Value As Boolean)
PropFadeAnimation = Value
If ToolTipHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ToolTipHandle, GWL_STYLE)
    If PropFadeAnimation = True Then
        If (dwStyle And TTS_NOFADE) = TTS_NOFADE Then dwStyle = dwStyle And Not TTS_NOFADE
    Else
        If Not (dwStyle And TTS_NOFADE) = TTS_NOFADE Then dwStyle = dwStyle Or TTS_NOFADE
    End If
    SetWindowLong ToolTipHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "FadeAnimation"
End Property

Public Property Get UseSystemFont() As Boolean
Attribute UseSystemFont.VB_Description = "Returns/sets a value that determines whether the control uses the system font or the font specified by the 'Font' property."
UseSystemFont = PropUseSystemFont
End Property

Public Property Let UseSystemFont(ByVal Value As Boolean)
PropUseSystemFont = Value
If ToolTipHandle <> 0 Then
    If PropUseSystemFont = False Then
        SendMessage ToolTipHandle, WM_SETFONT, ToolTipFontHandle, ByVal 1&
    Else
        SendMessage ToolTipHandle, WM_SETFONT, 0, ByVal 1&
    End If
End If
UserControl.PropertyChanged "UseSystemFont"
End Property

Public Property Get Tools() As TipTools
Attribute Tools.VB_Description = "Returns a reference to a collection of tools."
If PropTools Is Nothing Then
    Set PropTools = New TipTools
    PropTools.FInit Me
End If
Set Tools = PropTools
End Property

Friend Sub FToolsAdd(ByRef NewTool As TipTool, ByVal ID As Long, Optional ByVal Text As String, Optional ByVal Centered As Boolean, Optional ByVal Transparent As Boolean)
If ToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    If GetToolInfo(ID, TI) = False Then
        With TI
        .cbSize = LenB(TI)
        .uFlags = TTF_SUBCLASS Or TTF_IDISHWND Or TTF_PARSELINKS
        If Centered = True Then .uFlags = .uFlags Or TTF_CENTERTIP
        If Transparent = True Then .uFlags = .uFlags Or TTF_TRANSPARENT
        .hWnd = UserControl.hWnd
        .uId = ID
        If Text = vbNullString Then
            .lpszText = LPSTR_TEXTCALLBACK
        Else
            .lpszText = StrPtr(Text)
        End If
        .lParam = ObjPtr(NewTool)
        ToolTipMaxTipLength = GetMax(ToolTipMaxTipLength, Len(Text) + 1)
        End With
        SendMessage ToolTipHandle, TTM_ADDTOOL, 0, ByVal VarPtr(TI)
    End If
End If
End Sub

Friend Sub FToolsRemove(ByVal ID As Long)
If ToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    If GetToolInfo(ID, TI) = True Then SendMessage ToolTipHandle, TTM_DELTOOL, 0, ByVal VarPtr(TI)
End If
End Sub

Friend Property Get FToolText(ByVal ID As Long) As String
If ToolTipHandle <> 0 Then
    Dim TI As TOOLINFO, Text As String
    If GetToolInfo(ID, TI, Text) = True Then FToolText = Text
End If
End Property

Friend Property Let FToolText(ByVal ID As Long, ByVal Value As String)
Attribute FToolText.VB_Description = "Returns/Sets the text for a tool tip. To declare a link the text must contain <A> and </A> respectively at the start and end."
If ToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    If GetToolInfo(ID, TI) = True Then
        If Value = vbNullString Then
            TI.lpszText = -1
        Else
            TI.lpszText = StrPtr(Value)
        End If
        ToolTipMaxTipLength = GetMax(ToolTipMaxTipLength, Len(Value) + 1)
        SendMessage ToolTipHandle, TTM_UPDATETIPTEXT, 0, ByVal VarPtr(TI)
    End If
End If
End Property

Friend Property Get FToolCentered(ByVal ID As Long) As Boolean
If ToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    If GetToolInfo(ID, TI) = True Then FToolCentered = CBool((TI.uFlags And TTF_CENTERTIP) = TTF_CENTERTIP)
End If
End Property

Friend Property Let FToolCentered(ByVal ID As Long, ByVal Value As Boolean)
If ToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    If GetToolInfo(ID, TI) = True Then
        If Value = True Then
            If Not (TI.uFlags And TTF_CENTERTIP) = TTF_CENTERTIP Then TI.uFlags = TI.uFlags Or TTF_CENTERTIP
        Else
            If (TI.uFlags And TTF_CENTERTIP) = TTF_CENTERTIP Then TI.uFlags = TI.uFlags And Not TTF_CENTERTIP
        End If
        Call SetToolInfo(ID, TI)
    End If
End If
End Property

Friend Property Get FToolTransparent(ByVal ID As Long) As Boolean
If ToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    If GetToolInfo(ID, TI) = True Then FToolTransparent = CBool((TI.uFlags And TTF_TRANSPARENT) = TTF_TRANSPARENT)
End If
End Property

Friend Property Let FToolTransparent(ByVal ID As Long, ByVal Value As Boolean)
If ToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    If GetToolInfo(ID, TI) = True Then
        If Value = True Then
            If Not (TI.uFlags And TTF_TRANSPARENT) = TTF_TRANSPARENT Then TI.uFlags = TI.uFlags Or TTF_TRANSPARENT
        Else
            If Not (TI.uFlags And TTF_TRANSPARENT) = TTF_TRANSPARENT Then TI.uFlags = TI.uFlags And Not TTF_TRANSPARENT
        End If
        Call SetToolInfo(ID, TI)
    End If
End If
End Property

Friend Sub FToolTrack(ByVal ID As Long, ByVal State As Boolean)
If ToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    If GetToolInfo(ID, TI) = True Then
        SendMessage ToolTipHandle, TTM_TRACKACTIVATE, IIf(State = True, 1, 0), ByVal VarPtr(TI)
    End If
End If
End Sub

Private Sub CreateToolTip()
If ToolTipHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_POPUP Or TTS_ALWAYSTIP Or TTS_NOPREFIX
If PropBalloon = True Then dwStyle = dwStyle Or TTS_BALLOON
If PropCloseButton = True And ComCtlsSupportLevel() >= 1 Then dwStyle = dwStyle Or TTS_CLOSE
If PropFadeAnimation = False Then dwStyle = dwStyle Or TTS_NOFADE
dwExStyle = WS_EX_TOOLWINDOW
If Ambient.RightToLeft = True Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
Dim hWndParent As Long
hWndParent = GetAncestor(UserControl.ContainerHwnd, GA_ROOT)
If hWndParent <> 0 Then ToolTipHandle = CreateWindowEx(dwExStyle, StrPtr("tooltips_class32"), StrPtr("Tool Tip"), dwStyle, 0, 0, 0, 0, hWndParent, 0, App.hInstance, ByVal 0&)
If ToolTipHandle <> 0 Then SendMessage ToolTipHandle, TTM_SETMAXTIPWIDTH, 0, ByVal PropMaxTipWidth
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.BackColor = PropBackColor
Me.ForeColor = PropForeColor
Me.Title = PropTitle
If Ambient.UserMode = True Then Call ComCtlsSetSubclass(UserControl.hWnd, Me, 0)
End Sub

Private Sub DestroyToolTip()
If ToolTipHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(UserControl.hWnd)
SetParent ToolTipHandle, 0
DestroyWindow ToolTipHandle
ToolTipHandle = 0
If ToolTipFontHandle <> 0 Then
    DeleteObject ToolTipFontHandle
    ToolTipFontHandle = 0
End If
End Sub

Public Property Get DelayTime(ByVal TimeType As TipDelayTimeConstants) As Long
Attribute DelayTime.VB_Description = "Returns/sets a custom delay time (in milliseconds) for a specified delay time type."
Attribute DelayTime.VB_MemberFlags = "400"
Select Case Time
    Case TipDelayTimeReshow, TipDelayTimeShow, TipDelayTimeInitial
        If ToolTipHandle <> 0 Then DelayTime = SendMessage(ToolTipHandle, TTM_GETDELAYTIME, Time, ByVal 0&)
    Case Else
        Err.Raise 380
End Select
End Property

Public Property Let DelayTime(ByVal TimeType As TipDelayTimeConstants, ByVal Value As Long)
Select Case TimeType
    Case TipDelayTimeReshow, TipDelayTimeShow, TipDelayTimeInitial
        Select Case Value
            Case 0 To 32767
                If ToolTipHandle <> 0 Then SendMessage ToolTipHandle, TTM_SETDELAYTIME, TimeType, ByVal MakeDWord(Value, 0)
            Case Else
                Err.Raise 380
        End Select
    Case Else
        Err.Raise 380
End Select
End Property

Public Sub RestoreDelayTime()
Attribute RestoreDelayTime.VB_Description = "Restores all delay time types to default."
Const TTDT_AUTOMATIC As Long = 0
If ToolTipHandle <> 0 Then SendMessage ToolTipHandle, TTM_SETDELAYTIME, TTDT_AUTOMATIC, ByVal -1&
End Sub

Public Function HasToolTip(ByVal hWnd As Long) As Boolean
Attribute HasToolTip.VB_Description = "Returns a value that determines if a specified window is linked to an tool tip or not."
Dim TI As TOOLINFO
HasToolTip = GetToolInfo(hWnd, TI)
End Function

Public Sub HideCurrent()
Attribute HideCurrent.VB_Description = "Hides the current tool tip."
If ToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    TI.cbSize = LenB(TI)
    If SendMessage(ToolTipHandle, TTM_GETCURRENTTOOL, 0, ByVal VarPtr(TI)) <> 0 Then
        SendMessage ToolTipHandle, TTM_TRACKACTIVATE, 0, ByVal VarPtr(TI)
        SendMessage ToolTipHandle, TTM_POP, 0, ByVal 0&
    End If
End If
End Sub

Public Property Get LeftMargin() As Single
Attribute LeftMargin.VB_Description = "Returns/sets the widths of the left margin. This has no visible effect if the version of comctl32.dll is 6.1 or higher and the visual styles property is set to true."
Attribute LeftMargin.VB_MemberFlags = "400"
If ToolTipHandle <> 0 Then
    Dim RC As RECT
    SendMessage ToolTipHandle, TTM_GETMARGIN, 0, ByVal VarPtr(RC)
    LeftMargin = UserControl.ScaleX(RC.Left, vbPixels, vbContainerSize)
End If
End Property

Public Property Let LeftMargin(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
Dim LngValue As Long
LngValue = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If ToolTipHandle <> 0 Then
    Dim RC As RECT
    SendMessage ToolTipHandle, TTM_GETMARGIN, 0, ByVal VarPtr(RC)
    RC.Left = LngValue
    SendMessage ToolTipHandle, TTM_SETMARGIN, 0, ByVal VarPtr(RC)
End If
End Property

Public Property Get TopMargin() As Single
Attribute TopMargin.VB_Description = "Returns/sets the widths of the top margin. This has no visible effect if the version of comctl32.dll is 6.1 or higher and the visual styles property is set to true."
Attribute TopMargin.VB_MemberFlags = "400"
If ToolTipHandle <> 0 Then
    Dim RC As RECT
    SendMessage ToolTipHandle, TTM_GETMARGIN, 0, ByVal VarPtr(RC)
    TopMargin = UserControl.ScaleY(RC.Top, vbPixels, vbContainerSize)
End If
End Property

Public Property Let TopMargin(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
Dim LngValue As Long
LngValue = CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels))
If ToolTipHandle <> 0 Then
    Dim RC As RECT
    SendMessage ToolTipHandle, TTM_GETMARGIN, 0, ByVal VarPtr(RC)
    RC.Top = LngValue
    SendMessage ToolTipHandle, TTM_SETMARGIN, 0, ByVal VarPtr(RC)
End If
End Property

Public Property Get RightMargin() As Single
Attribute RightMargin.VB_Description = "Returns/sets the widths of the right margin. This has no visible effect if the version of comctl32.dll is 6.1 or higher and the visual styles property is set to true."
Attribute RightMargin.VB_MemberFlags = "400"
If ToolTipHandle <> 0 Then
    Dim RC As RECT
    SendMessage ToolTipHandle, TTM_GETMARGIN, 0, ByVal VarPtr(RC)
    RightMargin = UserControl.ScaleX(RC.Right, vbPixels, vbContainerSize)
End If
End Property

Public Property Let RightMargin(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
Dim LngValue As Long
LngValue = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If ToolTipHandle <> 0 Then
    Dim RC As RECT
    SendMessage ToolTipHandle, TTM_GETMARGIN, 0, ByVal VarPtr(RC)
    RC.Right = LngValue
    SendMessage ToolTipHandle, TTM_SETMARGIN, 0, ByVal VarPtr(RC)
End If
End Property

Public Property Get BottomMargin() As Single
Attribute BottomMargin.VB_Description = "Returns/sets the widths of the bottom margin. This has no visible effect if the version of comctl32.dll is 6.1 or higher and the visual styles property is set to true."
Attribute BottomMargin.VB_MemberFlags = "400"
If ToolTipHandle <> 0 Then
    Dim RC As RECT
    SendMessage ToolTipHandle, TTM_GETMARGIN, 0, ByVal VarPtr(RC)
    BottomMargin = UserControl.ScaleY(RC.Bottom, vbPixels, vbContainerSize)
End If
End Property

Public Property Let BottomMargin(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
Dim LngValue As Long
LngValue = CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels))
If ToolTipHandle <> 0 Then
    Dim RC As RECT
    SendMessage ToolTipHandle, TTM_GETMARGIN, 0, ByVal VarPtr(RC)
    RC.Bottom = LngValue
    SendMessage ToolTipHandle, TTM_SETMARGIN, 0, ByVal VarPtr(RC)
End If
End Property

Private Function GetToolInfo(ByVal ID As Long, ByRef TI As TOOLINFO, Optional ByRef Text As String) As Boolean
If ToolTipHandle = 0 Then Exit Function
TI.cbSize = LenB(TI)
TI.hWnd = UserControl.hWnd
TI.uId = ID
Dim Buffer As String
Buffer = String(ToolTipMaxTipLength, vbNullChar)
TI.lpszText = StrPtr(Buffer)
If SendMessage(ToolTipHandle, TTM_GETTOOLINFO, 0, ByVal VarPtr(TI)) <> 0 Then
    GetToolInfo = True
    Text = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End If
End Function

Private Sub SetToolInfo(ByVal ID As Long, ByRef TI As TOOLINFO)
If ToolTipHandle <> 0 Then
    SendMessage ToolTipHandle, TTM_SETTOOLINFO, 0, ByVal VarPtr(TI)
    SendMessage ToolTipHandle, TTM_UPDATE, 0, ByVal 0&
End If
End Sub

Private Function GetToolFromID(ByVal ID As Long) As TipTool
Dim TI As TOOLINFO
If GetToolInfo(ID, TI) = True Then
    If TI.lParam <> 0 Then Set GetToolFromID = PtrToObj(TI.lParam)
End If
End Function

Private Function GetMax(ByVal Param1 As Long, ByVal Param2 As Long) As Long
If Param1 > Param2 Then
    GetMax = Param1
Else
    GetMax = Param2
End If
End Function

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = ToolTipHandle Then
            Dim TI As TOOLINFO
            TI.cbSize = LenB(TI)
            Select Case NM.Code
                Case TTN_SHOW
                    RaiseEvent Show(GetToolFromID(NM.IDFrom))
                Case TTN_POP
                    RaiseEvent Hide(GetToolFromID(NM.IDFrom))
                    If ToolTipHandle <> 0 Then
                        If SendMessage(ToolTipHandle, TTM_GETCURRENTTOOL, 0, ByVal VarPtr(TI)) <> 0 Then
                            SendMessage ToolTipHandle, TTM_TRACKACTIVATE, 0, ByVal VarPtr(TI)
                        End If
                    End If
                Case TTN_LINKCLICK
                    SendMessage ToolTipHandle, TTM_GETCURRENTTOOL, 0, ByVal VarPtr(TI)
                    RaiseEvent LinkClick(GetToolFromID(TI.uId))
                Case TTN_GETDISPINFO
                    Dim NMTTDI As NMTTDISPINFO
                    CopyMemory NMTTDI, ByVal lParam, LenB(NMTTDI)
                    With NMTTDI
                    Dim Text As String
                    RaiseEvent NeedText(GetToolFromID(.hdr.IDFrom), Text)
                    If Not Text = vbNullString Then
                        With NMTTDI
                        If Len(Text) <= 80 Then
                            Text = Left$(Text & vbNullChar, 80)
                            CopyMemory .szText(0), ByVal StrPtr(Text), LenB(Text)
                        Else
                            .lpszText = StrPtr(Text)
                        End If
                        .hInst = 0
                        End With
                        CopyMemory ByVal lParam, NMTTDI, LenB(NMTTDI)
                    End If
                    End With
            End Select
        End If
    Case WM_NOTIFYFORMAT
        Const NF_QUERY As Long = 3
        If wParam = ToolTipHandle And lParam = NF_QUERY Then
            Const NFR_ANSI As Long = 1
            Const NFR_UNICODE As Long = 2
            WindowProcUserControl = NFR_UNICODE
            Exit Function
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function
