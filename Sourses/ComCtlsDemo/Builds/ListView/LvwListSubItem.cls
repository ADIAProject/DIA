VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwListSubItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowListView As ListView
Private PropIndex As Long
Private PropPtr As Long, PropSubPtr As Long

Friend Sub FInit(ByVal Object As ListView, ByVal Index As Long, ByVal Ptr As Long, ByVal SubPtr As Long)
Set ShadowListView = Object
PropIndex = Index
PropPtr = Ptr
PropSubPtr = SubPtr
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the list sub item."
Index = ShadowListView.FListSubItemSubIndex(PropSubPtr)
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key used in the collection."
Key = ShadowListView.FListSubItemKey(PropSubPtr)
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = ShadowListView.FListSubItemTag(PropSubPtr)
End Property

Public Property Let Tag(ByVal Value As String)
ShadowListView.FListSubItemTag(PropSubPtr) = Value
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text of the list sub item."
Text = ShadowListView.FListSubItemText(PropSubPtr)
End Property

Public Property Let Text(ByVal Value As String)
With ShadowListView
If .FListItemVerify(PropPtr, PropIndex) = True Then
    .FListSubItemText(PropSubPtr) = Value
    .FListItemText(PropIndex, .FListSubItemSubIndex(PropSubPtr)) = Value
End If
End With
End Property

Public Property Get ReportIcon() As Long
Attribute ReportIcon.VB_Description = "Returns/sets the index of an report icon in an associated image list control."
ReportIcon = ShadowListView.FListSubItemReportIcon(PropSubPtr)
End Property

Public Property Let ReportIcon(ByVal Value As Long)
If Value < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
With ShadowListView
.FListSubItemReportIcon(PropSubPtr) = Value
If .FListItemVerify(PropPtr, PropIndex) = True Then .FListItemRedraw PropIndex
End With
End Property

Public Property Get Bold() As Boolean
Attribute Bold.VB_Description = "Returns/sets a value that determines whether the text is displayed with a bold font or not."
Bold = ShadowListView.FListSubItemBold(PropSubPtr)
End Property

Public Property Let Bold(ByVal Value As Boolean)
With ShadowListView
If .FListItemVerify(PropPtr, PropIndex) = True Then
    .FListSubItemBold(PropSubPtr) = Value
    .FListItemRedraw PropIndex
End If
End With
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
With ShadowListView
ForeColor = .FListSubItemForeColor(PropSubPtr)
If ForeColor = -1 Then ForeColor = .ForeColor
End With
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
With ShadowListView
If .FListItemVerify(PropPtr, PropIndex) = True Then
    .FListSubItemForeColor(PropSubPtr) = Value
    .FListItemRedraw PropIndex
End If
End With
End Property
