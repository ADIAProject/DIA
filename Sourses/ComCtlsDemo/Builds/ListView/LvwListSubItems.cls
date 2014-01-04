VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwListSubItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements OLEGuids.IEnumerationVB
Private ShadowListView As ListView
Private PropIndex As Long
Private PropPtr As Long
Private PropListSubItem As Collection

Private Sub Class_Initialize()
Set PropListSubItem = New Collection
End Sub

Private Sub Class_Terminate()
Dim i As Long
With ShadowListView
For i = 1 To PropListSubItem.Count
    .FListSubItemFree PropListSubItem(i)
Next i
End With
Set PropListSubItem = Nothing
End Sub

Private Sub IEnumerationVB_Cleanup(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long)
End Sub

Private Sub IEnumerationVB_GetNextItem(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long, ByRef VNextItem As Variant, ByRef NoMoreItems As Boolean)
Dim ListItem As LvwListItem
Set ListItem = PtrToObj(PropPtr)
ShadowListView.FListSubItemsNextItem ListItem, Index, Control, Data, VNextItem, NoMoreItems
End Sub

Private Sub IEnumerationVB_Reset(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long)
Index = -1
Control = ShadowListView.FListItemsControl()
End Sub

Private Sub IEnumerationVB_Skip(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long, ByVal SkipCount As Long, ByRef SkippedAll As Boolean)
Index = Index + SkipCount
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Dim ThisEnum As New Enumeration
Set NewEnum = ThisEnum.GetEnum(Me, ShadowListView.FListItemsControl())
End Function

Friend Sub FInit(ByVal Object As ListView, ByVal Ptr As Long, ByVal Index As Long)
Set ShadowListView = Object
PropPtr = Ptr
PropIndex = Index
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal ReportIcon As Long) As LvwListSubItem
Attribute Add.VB_Description = "Adds a list sub item to the list sub items collection."
If Index > (PropListSubItem.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
If ReportIcon < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim SubPtr As Long
SubPtr = ShadowListView.FListSubItemAlloc(Key)
If Key = vbNullString Then
    If Index > 0 And Index <= PropListSubItem.Count Then
        PropListSubItem.Add SubPtr, , Index
    Else
        PropListSubItem.Add SubPtr
    End If
Else
    If Index > 0 And Index <= PropListSubItem.Count Then
        PropListSubItem.Add SubPtr, Key, Index
    Else
        PropListSubItem.Add SubPtr, Key
    End If
End If
Dim NewListSubItem As New LvwListSubItem
NewListSubItem.FInit ShadowListView, PropIndex, PropPtr, SubPtr
If Index = 0 Then
    ShadowListView.FListSubItemSubIndex(SubPtr) = PropListSubItem.Count
Else
    Dim i As Long
    With ShadowListView
    For i = Index To PropListSubItem.Count
        .FListSubItemSubIndex(SubPtr) = i
        If i <> .FListSubItemSubIndex(SubPtr) Then Me.Item(i).Text = Me.Item(i).Text
    Next i
    End With
End If
NewListSubItem.Text = Text
If ReportIcon > 0 Then ShadowListView.FListSubItemReportIcon(SubPtr) = ReportIcon
Set Add = NewListSubItem
End Function

Public Property Get Item(ByVal Index As Variant) As LvwListSubItem
Attribute Item.VB_Description = "Returns a list sub item given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = New LvwListSubItem
Item.FInit ShadowListView, PropIndex, PropPtr, PropListSubItem(Index)
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of list sub items in the collection."
Count = PropListSubItem.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all list sub items from the collection."
Dim i As Long
For i = 1 To PropListSubItem.Count
    Me.Remove 1
Next i
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a list sub item from the collection."
With ShadowListView
.FListSubItemFree PropListSubItem(Index)
PropListSubItem.Remove Index
If .FListItemVerify(PropPtr, PropIndex) = True Then
    Dim i As Long, SubPtr As Long
    For i = 1 To (PropListSubItem.Count + 1)
        If i <= PropListSubItem.Count Then
            SubPtr = PropListSubItem(i)
            .FListSubItemSubIndex(SubPtr) = i
            .FListItemText(PropIndex, i) = .FListSubItemText(SubPtr)
        Else
            .FListItemText(PropIndex, i) = vbNullString
        End If
    Next i
    .FListItemRedraw PropIndex
End If
End With
End Sub
