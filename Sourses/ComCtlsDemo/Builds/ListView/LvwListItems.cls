VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LvwListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements OLEGuids.IEnumerationVB
Private ShadowListView As ListView
Private PropListItem As Collection

Private Sub Class_Initialize()
Set PropListItem = New Collection
End Sub

Private Sub Class_Terminate()
Set PropListItem = Nothing
End Sub

Private Sub IEnumerationVB_Cleanup(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long)
End Sub

Private Sub IEnumerationVB_GetNextItem(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long, ByRef VNextItem As Variant, ByRef NoMoreItems As Boolean)
ShadowListView.FListItemsNextItem Index, Control, Data, VNextItem, NoMoreItems
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

Friend Sub FInit(ByVal Object As ListView)
Set ShadowListView = Object
End Sub

Public Function Add(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal Icon As Long, Optional ByVal SmallIcon As Long) As LvwListItem
Attribute Add.VB_Description = "Adds a list item to the list items collection."
If Index > (PropListItem.Count + 1) Or Index < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
If Icon < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
If SmallIcon < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim NewListItem As New LvwListItem
Dim Ptr As Long
Ptr = ObjPtr(NewListItem)
If Index = 0 Then Index = PropListItem.Count + 1
NewListItem.FInit ShadowListView, Index, Key, Ptr, Text, Icon, SmallIcon
ShadowListView.FListItemsAdd Ptr, Index, Text
If Key = vbNullString Then
    PropListItem.Add NewListItem
Else
    PropListItem.Add NewListItem, Key
End If
Set Add = NewListItem
End Function

Public Property Get Item(ByVal Index As Variant) As LvwListItem
Attribute Item.VB_Description = "Returns a list item given its index or key."
Attribute Item.VB_UserMemId = 0
Select Case VarType(Index)
    Case vbLong, vbInteger, vbByte
        Set Item = PtrToObj(ShadowListView.FListItemPtr(Index))
    Case vbString
        Set Item = PropListItem(Index)
    Case Else
        Err.Raise 13
End Select
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of list items in the collection."
Count = PropListItem.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all list items from the collection."
ShadowListView.FListItemsClear
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a list item from the collection."
Select Case VarType(Index)
    Case vbLong, vbInteger, vbByte
        Dim Ptr As Long, Item As Variant, i As Long
        Ptr = ShadowListView.FListItemPtr(Index)
        ShadowListView.FListItemsRemove Index
        For Each Item In PropListItem
            i = i + 1
            If ObjPtr(Item) = Ptr Then
                PropListItem.Remove i
                Exit For
            End If
        Next Item
    Case vbString
        ShadowListView.FListItemsRemove ShadowListView.FListItemIndex(ObjPtr(PropListItem(Index)))
        PropListItem.Remove Index
    Case Else
        Err.Raise 13
End Select
End Sub
