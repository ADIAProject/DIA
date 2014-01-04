VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TipTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowToolTip As ToolTip
Private PropTool As Collection

Private Sub Class_Initialize()
Set PropTool = New Collection
End Sub

Private Sub Class_Terminate()
Set PropTool = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropTool.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As ToolTip)
Set ShadowToolTip = Object
End Sub

Public Function Add(ByVal hWnd As Long, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal Centered As Boolean, Optional ByVal Transparent As Boolean) As TipTool
Attribute Add.VB_Description = "Adds a tool to the tools collection."
ShadowToolTip.FToolsAdd hWnd, Text, Centered, Transparent
Dim NewTool As New TipTool
NewTool.FInit ShadowToolTip, Key, hWnd
If Key = vbNullString Then
    PropTool.Add NewTool
Else
    PropTool.Add NewTool, Key
End If
NewTool.Index = PropTool.Count
Set Add = NewTool
End Function

Public Property Get Item(ByVal Index As Variant) As TipTool
Attribute Item.VB_Description = "Returns a tool given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropTool(Index)
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of tools in the collection."
Count = PropTool.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all tools from the collection."
Dim i As Long
For i = 1 To PropTool.Count
    Me.Remove 1
Next i
ShadowToolTip.FToolsClear
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a tool from the collection."
PropTool.Remove Index
Dim i As Long
For i = 1 To PropTool.Count
    Me.Item(i).Index = i
Next i
End Sub
