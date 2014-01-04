VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowListView As ListView
Private PropIndex As Long
Private PropPtr As Long, PropKey As String
Private PropTag As String
Private PropText As String
Private PropBold As Boolean
Private PropForeColor As OLE_COLOR
Private PropToolTipText As String
Private PropIcon As Long
Private PropSmallIcon As Long
Private PropListSubItems As LvwListSubItems

Private Sub Class_Initialize()
PropForeColor = -1
End Sub

Friend Sub FInit(ByVal Object As ListView, ByVal Index As Long, ByVal Key As String, ByVal Ptr As Long, ByVal Text As String, ByVal Icon As Long, ByVal SmallIcon As Long)
Set ShadowListView = Object
PropIndex = Index
PropKey = Key
PropPtr = Ptr
PropText = Text
PropIcon = Icon
PropSmallIcon = SmallIcon
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the list item."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Index = PropIndex
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

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text of the list item."
Attribute Text.VB_UserMemId = 0
Text = PropText
End Property

Public Property Let Text(ByVal Value As String)
PropText = Value
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemText(PropIndex, 0) = Value
End Property

Public Property Get Icon() As Long
Attribute Icon.VB_Description = "Returns/sets the index of an icon in an associated image list control."
Icon = PropIcon
End Property

Public Property Let Icon(ByVal Value As Long)
If Value < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
PropIcon = Value
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemRedraw PropIndex
End Property

Public Property Get SmallIcon() As Long
Attribute SmallIcon.VB_Description = "Returns/sets the index of an small icon in an associated image list control."
SmallIcon = PropSmallIcon
End Property

Public Property Let SmallIcon(ByVal Value As Long)
If Value < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
PropSmallIcon = Value
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemRedraw PropIndex
End Property

Public Property Get Indentation() As Long
Attribute Indentation.VB_Description = "Returns/sets the indentation of the list item in 'report' view in icon widths."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Indentation = ShadowListView.FListItemIndentation(PropIndex)
End Property

Public Property Let Indentation(ByVal Value As Long)
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemIndentation(PropIndex) = Value
End Property

Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets the selected state of the list item."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Selected = ShadowListView.FListItemSelected(PropIndex)
End Property

Public Property Let Selected(ByVal Value As Boolean)
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemSelected(PropIndex) = Value
End Property

Public Property Get Checked() As Boolean
Attribute Checked.VB_Description = "Returns/sets a value that determines whether the list item is checked or not."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Checked = ShadowListView.FListItemChecked(PropIndex)
End Property

Public Property Let Checked(ByVal Value As Boolean)
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemChecked(PropIndex) = Value
End Property

Public Property Get Ghosted() As Boolean
Attribute Ghosted.VB_Description = "Returns/sets the ghosted state of the list item."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Ghosted = ShadowListView.FListItemGhosted(PropIndex)
End Property

Public Property Let Ghosted(ByVal Value As Boolean)
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemGhosted(PropIndex) = Value
End Property

Public Property Get Hot() As Boolean
Attribute Hot.VB_Description = "Returns/sets a value that determines whether the list item is hot or not. This is only meaningful if the hot tracking property is set to true."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Hot = ShadowListView.FListItemHot(PropIndex)
End Property

Public Property Let Hot(ByVal Value As Boolean)
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemHot(PropIndex) = Value
End Property

Public Property Get Bold() As Boolean
Attribute Bold.VB_Description = "Returns/sets a value that determines whether the text is displayed with a bold font or not."
Bold = PropBold
End Property

Public Property Let Bold(ByVal Value As Boolean)
PropBold = Value
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemRedraw PropIndex
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
If PropForeColor = -1 Then
    ForeColor = ShadowListView.ForeColor
Else
    ForeColor = PropForeColor
End If
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemRedraw PropIndex
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the tool tip text of the list item. Requires that the show info tips property is set to true."
ToolTipText = PropToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
PropToolTipText = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns the distance between the internal left edge of an object and the left edge of its container."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Left = ShadowListView.FListItemLeft(PropIndex)
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns the distance between the internal top edge of an object and the top edge of its container."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Top = ShadowListView.FListItemTop(PropIndex)
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns the width of the list item."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Width = ShadowListView.FListItemWidth(PropIndex)
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns the height of the list item."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then Height = ShadowListView.FListItemHeight(PropIndex)
End Property

Public Sub EnsureVisible()
Attribute EnsureVisible.VB_Description = "Ensures a list item is visible, scrolling or expanding the control if necessary."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemEnsureVisible PropIndex
End Sub

Public Property Get TileViewIndices() As Variant
Attribute TileViewIndices.VB_Description = "Returns/sets an array of column header indices specifying the list sub items that are displayed for this list item in 'tile' view. Requires comctl32.dll version 6.0 or higher."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then TileViewIndices = ShadowListView.FListItemTileViewIndices(PropIndex)
End Sub

Public Property Let TileViewIndices(ByVal ArgList As Variant)
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then ShadowListView.FListItemTileViewIndices(PropIndex) = ArgList
End Property

Public Function CreateDragImage(Optional ByRef X As Single, Optional ByRef Y As Single) As Long
Attribute CreateDragImage.VB_Description = "Retrieves the handle to an image list containing a bitmap that can be used to visualize dragging of this list item.Your application is responsible for destroying the image list when it is no longer needed."
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then CreateDragImage = ShadowListView.FListItemCreateDragImage(PropIndex, X, Y)
End Function

Public Property Get ListSubItems() As LvwListSubItems
Attribute ListSubItems.VB_Description = "Returns/sets a collection of list sub items belonging to the list item."
If PropListSubItems Is Nothing Then
    Set PropListSubItems = New LvwListSubItems
    PropListSubItems.FInit ShadowListView, PropPtr, PropIndex
End If
Set ListSubItems = PropListSubItems
End Property

Public Property Get SubItems(ByVal Index As Integer) As String
Attribute SubItems.VB_Description = "Returns/sets an array of strings representing the list sub items text."
If Index < 1 Or Index > (ShadowListView.ColumnHeaders.Count - 1) Then Err.Raise 380
If ShadowListView.FListItemVerify(PropPtr, PropIndex) = True Then SubItems = ShadowListView.FListItemText(PropIndex, Index)
End Property

Public Property Let SubItems(ByVal Index As Integer, ByVal Value As String)
If Index < 1 Or Index > (ShadowListView.ColumnHeaders.Count - 1) Then Err.Raise 380
If Index > Me.ListSubItems.Count Then
    Dim i As Long
    For i = (Me.ListSubItems.Count + 1) To Index Step 1
        If i < Index Then
            Me.ListSubItems.Add
        Else
            Me.ListSubItems.Add , , Value
        End If
    Next i
Else
    Me.ListSubItems(Index).Text = Value
End If
End Property
