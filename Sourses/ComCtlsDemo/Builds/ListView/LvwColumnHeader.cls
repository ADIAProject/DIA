VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwColumnHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowListView As ListView
Private PropIndex As Long, PropKey As String
Private PropTag As String

Friend Sub FInit(ByVal Object As ListView, ByVal Key As String)
Set ShadowListView = Object
PropKey = Key
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the column."
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

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text of the column header."
Text = ShadowListView.FColumnHeaderText(PropIndex)
End Property

Public Property Let Text(ByVal Value As String)
ShadowListView.FColumnHeaderText(PropIndex) = Value
End Property

Public Property Get Icon() As Long
Attribute Icon.VB_Description = "Returns/sets the index of an icon in an associated image list control."
Icon = ShadowListView.FColumnHeaderIcon(PropIndex)
End Property

Public Property Let Icon(ByVal Value As Long)
If Value < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
ShadowListView.FColumnHeaderIcon(PropIndex) = Value
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of the column header."
Width = ShadowListView.FColumnHeaderWidth(PropIndex)
End Property

Public Property Let Width(ByVal Value As Single)
ShadowListView.FColumnHeaderWidth(PropIndex) = Value
End Property

Public Property Get Alignment() As LvwColumnHeaderAlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of the column header."
Alignment = ShadowListView.FColumnHeaderAlignment(PropIndex)
End Property

Public Property Let Alignment(ByVal Value As LvwColumnHeaderAlignmentConstants)
ShadowListView.FColumnHeaderAlignment(PropIndex) = Value
End Property

Public Property Get Position() As Long
Attribute Position.VB_Description = "Returns/sets the current position of the column header."
Position = ShadowListView.FColumnHeaderPosition(PropIndex)
End Property

Public Property Let Position(ByVal Value As Long)
ShadowListView.FColumnHeaderPosition(PropIndex) = Value
End Property

Public Property Get SortArrow() As LvwColumnHeaderSortArrowConstants
Attribute SortArrow.VB_Description = "Returns/sets the sort arrow of the column header. Requires comctl32.dll version 6.0 or higher."
SortArrow = ShadowListView.FColumnHeaderSortArrow(PropIndex)
End Property

Public Property Let SortArrow(ByVal Value As LvwColumnHeaderSortArrowConstants)
ShadowListView.FColumnHeaderSortArrow(PropIndex) = Value
End Property

Public Property Get IconOnRight() As Boolean
Attribute IconOnRight.VB_Description = "Returns/sets a value that determines if the icon appears on the right of the column header."
IconOnRight = ShadowListView.FColumnHeaderIconOnRight(PropIndex)
End Property

Public Property Let IconOnRight(ByVal Value As Boolean)
ShadowListView.FColumnHeaderIconOnRight(PropIndex) = Value
End Property

Public Property Get Resizable() As Boolean
Attribute Resizable.VB_Description = "Returns/sets a value that determines whether or not the column header can be modified by a user action to resize it. Requires comctl32.dll version 6.1 or higher."
Resizable = ShadowListView.FColumnHeaderResizable(PropIndex)
End Property

Public Property Let Resizable(ByVal Value As Boolean)
ShadowListView.FColumnHeaderResizable(PropIndex) = Value
End Property

Public Property Get SplitButton() As Boolean
Attribute SplitButton.VB_Description = "Returns/sets a value that determines whether or not the column header displays a split button. The 'ColumnDropDown' event is fired when the split button is clicked. Requires comctl32.dll version 6.1 or higher."
SplitButton = ShadowListView.FColumnHeaderSplitButton(PropIndex)
End Property

Public Property Let SplitButton(ByVal Value As Boolean)
ShadowListView.FColumnHeaderSplitButton(PropIndex) = Value
End Property

Public Property Get CheckBox() As Boolean
Attribute CheckBox.VB_Description = "Returns/sets a value that determines whether or not the column header displays a checkbox. This is only meaningful if the allow column checkboxes property is set to true. Requires comctl32.dll version 6.1 or higher."
CheckBox = ShadowListView.FColumnHeaderCheckBox(PropIndex)
End Property

Public Property Let CheckBox(ByVal Value As Boolean)
ShadowListView.FColumnHeaderCheckBox(PropIndex) = Value
End Property

Public Property Get Checked() As Boolean
Attribute Checked.VB_Description = "Returns/sets the checkbox state of the column header. Requires comctl32.dll version 6.1 or higher."
Checked = ShadowListView.FColumnHeaderChecked(PropIndex)
End Property

Public Property Let Checked(ByVal Value As Boolean)
ShadowListView.FColumnHeaderChecked(PropIndex) = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
Left = ShadowListView.FColumnHeaderLeft(PropIndex)
End Property

Public Sub AutoSize(ByVal Value As LvwColumnHeaderAutoSizeConstants)
Attribute AutoSize.VB_Description = "Automatically sizes the column header."
ShadowListView.FColumnHeaderAutoSize PropIndex, Value
End Sub

Public Function SubItemIndex() As Long
Attribute SubItemIndex.VB_Description = "Returns the index of the sub item associated with the column header."
SubItemIndex = ShadowListView.FColumnHeaderSubItemIndex(PropIndex)
End Function
