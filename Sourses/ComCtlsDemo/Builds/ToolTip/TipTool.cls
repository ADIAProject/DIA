VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TipTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowToolTip As ToolTip
Private PropIndex As Long, PropKey As String, PropID As Long
Private PropTag As String

Private Sub Class_Terminate()
ShadowToolTip.FToolsRemove PropID
End Sub

Friend Sub FInit(ByVal Object As ToolTip, ByVal Key As String, ByVal ID As Long)
Set ShadowToolTip = Object
PropKey = Key
PropID = ID
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the tool."
Index = PropIndex
End Property

Friend Property Let Index(ByVal Value As Long)
PropIndex = Value
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key used in the collection."
Key = PropKey
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = PropTag
End Property

Public Property Let Tag(ByVal Value As String)
PropTag = Value
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a tool tip identification. This is the same as the target window handle."
hWnd = PropID
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the tool tip text."
Text = ShadowToolTip.FToolText(PropID)
End Property

Public Property Let Text(ByVal Value As String)
ShadowToolTip.FToolText(PropID) = Value
End Property

Public Property Get Centered() As Boolean
Attribute Centered.VB_Description = "Returns/sets a value that determines whether the tool tip will be displayed centered or not."
Centered = ShadowToolTip.FToolCentered(PropID)
End Property

Public Property Let Centered(ByVal Value As Boolean)
ShadowToolTip.FToolCentered(PropID) = Value
End Property

Public Property Get Transparent() As Boolean
Attribute Transparent.VB_Description = "Returns/sets a value that determines whether the tool tip ignore mouse messages or not."
Centered = ShadowToolTip.FToolTransparent(PropID)
End Property

Public Property Let Transparent(ByVal Value As Boolean)
ShadowToolTip.FToolTransparent(PropID) = Value
End Property

Public Sub Show()
Attribute Show.VB_Description = "Shows the tool tip and activate tracking."
ShadowToolTip.FToolTrack PropID, True
End Sub

Public Sub Hide()
Attribute Hide.VB_Description = "Hides the tool tip and deactivate tracking."
ShadowToolTip.FToolTrack PropID, False
End Sub
