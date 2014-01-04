VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enumeration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements OLEGuids.IEnumeration
Private PropControl As Long, PropIndex As Long, PropData As Long
Private Owner As OLEGuids.IEnumerationVB

Private Sub Class_Initialize()
Call SetVTableSubclass(Me, VTableInterfaceEnumeration)
End Sub

Private Sub Class_Terminate()
If Not Owner Is Nothing Then Owner.Cleanup PropIndex, PropControl, PropData
Set Owner = Nothing
Call RemoveVTableSubclass(Me, VTableInterfaceEnumeration)
End Sub

Private Sub IEnumeration_Next(ByVal cElt As Long, rgVar As Variant, ByVal pcEltFetched As Long)
End Sub

Private Sub IEnumeration_Skip(ByVal cElt As Long)
End Sub

Private Sub IEnumeration_Reset()
End Sub

Private Sub IEnumeration_Clone(ByRef ppEnum As IEnumVARIANT)
End Sub

Public Function GetEnum(ByVal This As OLEGuids.IEnumerationVB, ByVal Control As Long, Optional ByVal Data As Long, Optional ByVal Index As Long = -1) As IEnumVARIANT
PropIndex = Index
PropControl = Control
PropData = Data
Set GetEnum = Me
Set Owner = This
End Function

Friend Sub GetNextItem(ByRef VNextItem As Variant, ByRef NoMoreItems As Boolean)
If Not Owner Is Nothing Then Owner.GetNextItem PropIndex, PropControl, PropData, VNextItem, NoMoreItems Else NoMoreItems = True
End Sub

Friend Sub Reset()
If Not Owner Is Nothing Then Owner.Reset PropIndex, PropControl, PropData
End Sub

Friend Sub Skip(ByVal SkipCount As Long, ByRef SkippedAll As Boolean)
If Not Owner Is Nothing Then Owner.Skip PropIndex, PropControl, PropData, SkipCount, SkippedAll
End Sub
