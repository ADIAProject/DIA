VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowListView As ListView
Private PropIndex As Long, PropKey As String, PropID As Long
Private PropTag As String

Friend Sub FInit(ByVal Object As ListView, ByVal Key As String)
Set ShadowListView = Object
PropKey = Key
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the group."
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

Public Property Get ID() As Long
Attribute ID.VB_Description = "Returns an internal identification of this group."
ID = PropID
End Property

Friend Property Let ID(ByVal Value As Long)
PropID = Value
End Property

Public Property Get Header() As String
Attribute Header.VB_Description = "Returns/sets the header text for the group. Requires comctl32.dll version 6.0 or higher."
Attribute Header.VB_UserMemId = 0
Header = ShadowListView.FGroupHeader(PropID)
End Property

Public Property Let Header(ByVal Value As String)
ShadowListView.FGroupHeader(PropID) = Value
End Property

Public Property Get HeaderAlignment() As LvwGroupHeaderAlignmentConstants
Attribute HeaderAlignment.VB_Description = "Returns/sets the alignment of the group header text. Requires comctl32.dll version 6.0 or higher."
HeaderAlignment = ShadowListView.FGroupHeaderAlignment(PropID)
End Property

Public Property Let HeaderAlignment(ByVal Value As LvwGroupHeaderAlignmentConstants)
ShadowListView.FGroupHeaderAlignment(PropID) = Value
End Property

Public Property Get Footer() As String
Attribute Footer.VB_Description = "Returns/sets the footer text for the group. Requires comctl32.dll version 6.1 or higher."
Footer = ShadowListView.FGroupFooter(PropID)
End Property

Public Property Let Footer(ByVal Value As String)
ShadowListView.FGroupFooter(PropID) = Value
End Property

Public Property Get FooterAlignment() As LvwGroupFooterAlignmentConstants
Attribute FooterAlignment.VB_Description = "Returns/sets the alignment of the group footer text. Requires comctl32.dll version 6.1 or higher."
FooterAlignment = ShadowListView.FGroupFooterAlignment(PropID)
End Property

Public Property Let FooterAlignment(ByVal Value As LvwGroupFooterAlignmentConstants)
ShadowListView.FGroupFooterAlignment(PropID) = Value
End Property

Public Property Get Collapsible() As Boolean
Attribute Collapsible.VB_Description = "Returns/sets a value that determines whether the group can be collapsed or not. Requires comctl32.dll version 6.1 or higher."
Collapsible = ShadowListView.FGroupCollapsible(PropID)
End Property

Public Property Let Collapsible(ByVal Value As Boolean)
ShadowListView.FGroupCollapsible(PropID) = Value
End Property

Public Property Get Collapsed() As Boolean
Attribute Collapsed.VB_Description = "Returns/sets a value that determines whether the group is collapsed or not. Requires comctl32.dll version 6.1 or higher."
Collapsed = ShadowListView.FGroupCollapsed(PropID)
End Property

Public Property Let Collapsed(ByVal Value As Boolean)
ShadowListView.FGroupCollapsed(PropID) = Value
End Property

Public Property Get ShowHeader() As Boolean
Attribute ShowHeader.VB_Description = "Returns/sets a value that determines whether the group displays a header or not. Requires comctl32.dll version 6.1 or higher."
ShowHeader = ShadowListView.FGroupShowHeader(PropID)
End Property

Public Property Let ShowHeader(ByVal Value As Boolean)
ShadowListView.FGroupShowHeader(PropID) = Value
End Property

Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets a value that determines whether the group is selected or not. Requires comctl32.dll version 6.1 or higher."
Selected = ShadowListView.FGroupSelected(PropID)
End Property

Public Property Let Selected(ByVal Value As Boolean)
ShadowListView.FGroupSelected(PropID) = Value
End Property

Public Property Get Position() As Long
Attribute Position.VB_Description = "Returns the current position of the group."
Position = ShadowListView.FGroupPosition(PropID)
End Property

Public Property Let Position(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns the distance between the internal left edge of an object and the left edge of its container. Requires comctl32.dll version 6.1 or higher."
Left = ShadowListView.FGroupLeft(PropID)
End Property

Public Property Let Left(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns the distance between the internal top edge of an object and the top edge of its container. Requires comctl32.dll version 6.1 or higher."
Top = ShadowListView.FGroupTop(PropID)
End Property

Public Property Let Top(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns the width of the group. Requires comctl32.dll version 6.1 or higher."
Width = ShadowListView.FGroupWidth(PropID)
End Property

Public Property Let Width(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns the height of the group. Requires comctl32.dll version 6.1 or higher."
Height = ShadowListView.FGroupHeight(PropID)
End Property

Public Property Let Height(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get ListItemIndices() As Collection
Attribute ListItemIndices.VB_Description = "Returns a reference to a collection containing the indexes to the list items referring to this group. Requires comctl32.dll version 6.0 or higher."
Set ListItemIndices = ShadowListView.FGroupListItemIndices(PropID)
End Property
