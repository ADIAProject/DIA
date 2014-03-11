Attribute VB_Name = "VTableHandle"
Option Explicit

' Required:

' OLEGuids.tlb (in IDE only)

#If False Then
Private VTableInterfaceControl, VTableInterfaceInPlaceActiveObject, VTableInterfacePerPropertyBrowsing, VTableInterfaceEnumeration
#End If
Public Enum VTableInterfaceConstants
VTableInterfaceControl = 1
VTableInterfaceInPlaceActiveObject = 2
VTableInterfacePerPropertyBrowsing = 3
VTableInterfaceEnumeration = 4
End Enum
Private Enum VTableIndexIPAOConstants
' Ignore : IPAOQueryInterface = 1
' Ignore : IPAOAddRef = 2
' Ignore : IPAORelease = 3
' Ignore : IPAOGetWindow = 4
' Ignore : IPAOContextSensitiveHelp = 5
VTableIndexIPAOTranslateAccelerator = 6
' Ignore : IPAOOnFrameWindowActivate = 7
' Ignore : IPAOOnDocWindowActivate = 8
' Ignore : IPAOResizeBorder = 9
' Ignore : IPAOEnableModeless = 10
End Enum
Private Enum VTableIndexControlConstants
' Ignore : ControlQueryInterface = 1
' Ignore : ControlAddRef = 2
' Ignore : ControlRelease = 3
VTableIndexControlGetControlInfo = 4
VTableIndexControlOnMnemonic = 5
' Ignore : ControlOnAmbientPropertyChange = 6
' Ignore : ControlFreezeEvents = 7
End Enum
Private Enum VTableIndexPPBConstants
' Ignore : PPBQueryInterface = 1
' Ignore : PPBAddRef = 2
' Ignore : PPBRelease = 3
VTableIndexPPBGetDisplayString = 4
' Ignore : PPBMapPropertyToPage = 5
VTAbleIndexPPBGetPredefinedStrings = 6
VTAbleIndexPPBGetPredefinedValue = 7
End Enum
Private Enum VTableIndexEnumerationConstants
' Ignore : EnumerationQueryInterface
' Ignore : EnumerationAddRef
' Ignore : EnumerationRelease
VTableIndexEnumerationNext = 4
VTableIndexEnumerationSkip = 5
VTableIndexEnumerationReset = 6
VTableIndexEnumerationClone = 7
End Enum
Public Const CTRLINFO_EATS_RETURN As Long = 1
Public Const CTRLINFO_EATS_ESCAPE As Long = 2
Private Type SAFEARRAYBOUND
cElements As Long
lLbound As Long
End Type
Private Type SAFEARRAY1D
cDims As Integer
fFeatures As Integer
cbElements As Long
cLocks As Long
pvData As Long
Bounds(0 To 0) As SAFEARRAYBOUND
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef Var() As Any) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As Long
Private Declare Function SysAllocString Lib "oleaut32" (ByVal lpString As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As IUnknown, ByVal oVft As Long, ByVal CallConv As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByRef pvargResult As Variant) As Long
Private Const CC_STDCALL As Long = 4
Private Const E_OUTOFMEMORY As Long = &H8007000E
Private Const E_POINTER As Long = &H80004003
Private Const E_INVALIDARG As Long = &H80070057
Private Const E_NOTIMPL As Long = &H80004001
Private Const S_FALSE As Long = &H1
Private Const S_OK As Long = &H0
Private VTableSubclassIPAO As VTableSubclass
Private VTableSubclassControl As VTableSubclass
Private VTableSubclassPPB As VTableSubclass, StringsOutArray() As String, CookiesOutArray() As Long
Private VTableSubclassEnumeration As VTableSubclass, SAHeader As SAFEARRAY1D, VariantArray() As Variant

Public Sub SetVTableSubclass(ByVal This As Object, ByVal OLEInterface As VTableInterfaceConstants)
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        If VTableSupported(This, VTableInterfaceInPlaceActiveObject) = True Then Call ReplaceIOleIPAO(This)
    Case VTableInterfaceControl
        If VTableSupported(This, VTableInterfaceControl) = True Then Call ReplaceIOleControl(This)
    Case VTableInterfacePerPropertyBrowsing
        If VTableSupported(This, VTableInterfacePerPropertyBrowsing) = True Then Call ReplaceIPPB(This)
    Case VTableInterfaceEnumeration
        If VTableSupported(This, VTableInterfaceEnumeration) = True Then Call ReplaceIEnumeration(This)
End Select
End Sub

Public Sub RemoveVTableSubclass(ByVal This As Object, ByVal OLEInterface As VTableInterfaceConstants)
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        If VTableSupported(This, VTableInterfaceInPlaceActiveObject) = True Then Call RestoreIOleIPAO(This)
    Case VTableInterfaceControl
        If VTableSupported(This, VTableInterfaceControl) = True Then Call RestoreIOleControl(This)
    Case VTableInterfacePerPropertyBrowsing
        If VTableSupported(This, VTableInterfacePerPropertyBrowsing) = True Then Call RestoreIPPB(This)
    Case VTableInterfaceEnumeration
        If VTableSupported(This, VTableInterfaceEnumeration) = True Then Call RestoreIEnumeration(This)
End Select
End Sub

Public Sub RemoveAllVTableSubclass(ByVal OLEInterface As VTableInterfaceConstants)
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        Set VTableSubclassIPAO = Nothing
    Case VTableInterfaceControl
        Set VTableSubclassControl = Nothing
    Case VTableInterfacePerPropertyBrowsing
        Set VTableSubclassPPB = Nothing
    Case VTableInterfaceEnumeration
        Set VTableSubclassEnumeration = Nothing
End Select
End Sub

Private Function VTableSupported(ByRef This As Object, ByVal OLEInterface As VTableInterfaceConstants) As Boolean
On Error GoTo Cancel
Select Case OLEInterface
    Case VTableInterfaceInPlaceActiveObject
        Dim ShadowIOleIPAO As OLEGuids.IOleInPlaceActiveObject
        Dim ShadowIOleInPlaceActiveObjectVB As OLEGuids.IOleInPlaceActiveObjectVB
        Set ShadowIOleIPAO = This
        Set ShadowIOleInPlaceActiveObjectVB = This
        VTableSupported = Not CBool(ShadowIOleIPAO Is Nothing Or ShadowIOleInPlaceActiveObjectVB Is Nothing)
    Case VTableInterfaceControl
        Dim ShadowIOleControl As OLEGuids.IOleControl
        Dim ShadowIOleControlVB As OLEGuids.IOleControlVB
        Set ShadowIOleControl = This
        Set ShadowIOleControlVB = This
        VTableSupported = Not CBool(ShadowIOleControl Is Nothing Or ShadowIOleControlVB Is Nothing)
    Case VTableInterfacePerPropertyBrowsing
        Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
        Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB
        Set ShadowIPPB = This
        Set ShadowIPerPropertyBrowsingVB = This
        VTableSupported = Not CBool(ShadowIPPB Is Nothing Or ShadowIPerPropertyBrowsingVB Is Nothing)
    Case VTableInterfaceEnumeration
        Dim ShadowIEnumeration As OLEGuids.IEnumeration
        Set ShadowIEnumeration = This
        VTableSupported = Not CBool(ShadowIEnumeration Is Nothing)
End Select
Cancel:
End Function

Public Function VTableCall(ByVal RetType As VbVarType, ByVal OLEInstance As IUnknown, ByVal Entry As Long, ParamArray ArgList() As Variant) As Variant
Entry = Entry - 1
Debug.Assert Not (Entry < 0 Or OLEInstance Is Nothing)
Dim VarArgList As Variant, HResult As Long
VarArgList = ArgList
If UBound(VarArgList) > -1 Then
    Dim i As Long, ArrVarType() As Integer, ArrVarPtr() As Long
    ReDim ArrVarType(LBound(VarArgList) To UBound(VarArgList)) As Integer
    ReDim ArrVarPtr(LBound(VarArgList) To UBound(VarArgList)) As Long
    For i = LBound(VarArgList) To UBound(VarArgList)
        ArrVarType(i) = VarType(VarArgList(i))
        ArrVarPtr(i) = VarPtr(VarArgList(i))
    Next i
    HResult = DispCallFunc(OLEInstance, Entry * 4, CC_STDCALL, RetType, i, VarPtr(ArrVarType(0)), VarPtr(ArrVarPtr(0)), VTableCall)
Else
    HResult = DispCallFunc(OLEInstance, Entry * 4, CC_STDCALL, RetType, 0, 0, 0, VTableCall)
End If
Select Case HResult
    Case S_OK
    Case E_INVALIDARG
        Err.Raise Number:=HResult, Description:="One of the arguments was invalid"
    Case E_POINTER
        Err.Raise Number:=HResult, Description:="Function address was null"
    Case Else
        Err.Raise HResult
End Select
End Function

Private Sub ReplaceIOleIPAO(ByVal This As OLEGuids.IOleInPlaceActiveObject)
If VTableSubclassIPAO Is Nothing Then Set VTableSubclassIPAO = New VTableSubclass
If VTableSubclassIPAO.RefCount = 0 Then
    VTableSubclassIPAO.Subclass ObjPtr(This), VTableIndexIPAOTranslateAccelerator, VTableIndexIPAOTranslateAccelerator, _
    AddressOf IOleIPAO_TranslateAccelerator
End If
VTableSubclassIPAO.AddRef
End Sub

Private Sub RestoreIOleIPAO(ByVal This As OLEGuids.IOleInPlaceActiveObject)
If Not VTableSubclassIPAO Is Nothing Then
    VTableSubclassIPAO.Release
    If VTableSubclassIPAO.RefCount = 0 Then VTableSubclassIPAO.UnSubclass
End If
End Sub

Public Sub ActivateIPAO(ByVal This As Object)
On Error GoTo CATCH_EXCEPTION
Dim PropOleObject As OLEGuids.IOleObject
Dim PropOleInPlaceSite As OLEGuids.IOleInPlaceSite
Dim PropOleInPlaceFrame As OLEGuids.IOleInPlaceFrame
Dim PropOleInPlaceUIWindow As OLEGuids.IOleInPlaceUIWindow
Dim PropOleInPlaceActiveObject As OLEGuids.IOleInPlaceActiveObject
Dim PosRect As OLEGuids.OLERECT
Dim ClipRect As OLEGuids.OLERECT
Dim FrameInfo As OLEGuids.OLEINPLACEFRAMEINFO
Set PropOleObject = This
Set PropOleInPlaceActiveObject = This
Set PropOleInPlaceSite = PropOleObject.GetClientSite
PropOleInPlaceSite.GetWindowContext PropOleInPlaceFrame, PropOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
PropOleInPlaceFrame.SetActiveObject PropOleInPlaceActiveObject, vbNullString
If Not PropOleInPlaceUIWindow Is Nothing Then PropOleInPlaceUIWindow.SetActiveObject PropOleInPlaceActiveObject, vbNullString
CATCH_EXCEPTION:
End Sub

Private Function IOleIPAO_TranslateAccelerator(ByVal This As Object, ByRef Msg As OLEGuids.OLEACCELMSG) As Long
Dim ShadowIOleInPlaceActiveObjectVB As OLEGuids.IOleInPlaceActiveObjectVB
Dim Handled As Boolean
On Error GoTo CATCH_EXCEPTION
If VarPtr(Msg) = 0 Then
    IOleIPAO_TranslateAccelerator = E_POINTER
    Exit Function
End If
Set ShadowIOleInPlaceActiveObjectVB = This
IOleIPAO_TranslateAccelerator = S_OK
ShadowIOleInPlaceActiveObjectVB.TranslateAccelerator Handled, IOleIPAO_TranslateAccelerator, Msg.Message, Msg.wParam, Msg.lParam, GetShiftState()
If Handled = False Then IOleIPAO_TranslateAccelerator = Original_IOleIPAO_TranslateAccelerator(This, Msg)
Exit Function
CATCH_EXCEPTION:
IOleIPAO_TranslateAccelerator = Original_IOleIPAO_TranslateAccelerator(This, Msg)
End Function

Private Function Original_IOleIPAO_TranslateAccelerator(ByVal This As OLEGuids.IOleInPlaceActiveObject, ByRef Msg As OLEGuids.OLEACCELMSG) As Long
VTableSubclassIPAO.SubclassEntry(VTableIndexIPAOTranslateAccelerator) = False
Original_IOleIPAO_TranslateAccelerator = This.TranslateAccelerator(ByVal VarPtr(Msg))
VTableSubclassIPAO.SubclassEntry(VTableIndexIPAOTranslateAccelerator) = True
End Function

Private Sub ReplaceIOleControl(ByVal This As OLEGuids.IOleControl)
If VTableSubclassControl Is Nothing Then Set VTableSubclassControl = New VTableSubclass
If VTableSubclassControl.RefCount = 0 Then
    VTableSubclassControl.Subclass ObjPtr(This), VTableIndexControlGetControlInfo, VTableIndexControlOnMnemonic, _
    AddressOf IOleControl_GetControlInfo, _
    AddressOf IOleControl_OnMnemonic
End If
VTableSubclassControl.AddRef
End Sub

Private Sub RestoreIOleControl(ByVal This As OLEGuids.IOleControl)
If Not VTableSubclassControl Is Nothing Then
    VTableSubclassControl.Release
    If VTableSubclassControl.RefCount = 0 Then VTableSubclassControl.UnSubclass
End If
End Sub

Public Sub OnControlInfoChanged(ByVal This As Object, Optional ByVal OnFocus As Boolean)
Dim PropOleObject As OLEGuids.IOleObject
Dim PropClientSite As OLEGuids.IOleClientSite
Dim PropUnknown As IUnknown
Dim PropControlSite As OLEGuids.IOleControlSite
On Error Resume Next
Set PropOleObject = This
Set PropClientSite = PropOleObject.GetClientSite
Set PropUnknown = PropClientSite
Set PropControlSite = PropUnknown
PropControlSite.OnControlInfoChanged
If OnFocus = True Then PropControlSite.OnFocus 1
End Sub

Private Function IOleControl_GetControlInfo(ByVal This As Object, ByRef CI As OLEGuids.OLECONTROLINFO) As Long
Dim ShadowIOleControlVB As OLEGuids.IOleControlVB
Dim Handled As Boolean
On Error GoTo CATCH_EXCEPTION
If VarPtr(CI) = 0 Then
    IOleControl_GetControlInfo = E_POINTER
    Exit Function
End If
Set ShadowIOleControlVB = This
CI.cb = LenB(CI)
ShadowIOleControlVB.GetControlInfo Handled, CI.cAccel, CI.hAccel, CI.dwFlags
If Handled = False Then
    IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(This, CI)
Else
    If CI.cAccel > 0 And CI.hAccel = 0 Then
        IOleControl_GetControlInfo = E_OUTOFMEMORY
    Else
        IOleControl_GetControlInfo = S_OK
    End If
End If
Exit Function
CATCH_EXCEPTION:
IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(This, CI)
End Function

Private Function IOleControl_OnMnemonic(ByVal This As Object, ByRef Msg As OLEGuids.OLEACCELMSG) As Long
Dim ShadowIOleControlVB As OLEGuids.IOleControlVB
Dim Handled As Boolean
On Error GoTo CATCH_EXCEPTION
If VarPtr(Msg) = 0 Then
    IOleControl_OnMnemonic = E_POINTER
    Exit Function
End If
Set ShadowIOleControlVB = This
ShadowIOleControlVB.OnMnemonic Handled, Msg.Message, Msg.wParam, Msg.lParam, GetShiftState()
If Handled = False Then
    IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(This, Msg)
Else
    IOleControl_OnMnemonic = S_OK
End If
Exit Function
CATCH_EXCEPTION:
IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(This, Msg)
End Function

Private Function Original_IOleControl_GetControlInfo(ByVal This As OLEGuids.IOleControl, ByRef CI As OLEGuids.OLECONTROLINFO) As Long
VTableSubclassControl.SubclassEntry(VTableIndexControlGetControlInfo) = False
Original_IOleControl_GetControlInfo = This.GetControlInfo(CI)
VTableSubclassControl.SubclassEntry(VTableIndexControlGetControlInfo) = True
End Function

Private Function Original_IOleControl_OnMnemonic(ByVal This As OLEGuids.IOleControl, ByRef Msg As OLEGuids.OLEACCELMSG) As Long
VTableSubclassControl.SubclassEntry(VTableIndexControlOnMnemonic) = False
Original_IOleControl_OnMnemonic = This.OnMnemonic(Msg)
VTableSubclassControl.SubclassEntry(VTableIndexControlOnMnemonic) = True
End Function

Private Sub ReplaceIPPB(ByVal This As OLEGuids.IPerPropertyBrowsing)
If VTableSubclassPPB Is Nothing Then Set VTableSubclassPPB = New VTableSubclass
If VTableSubclassPPB.RefCount = 0 Then
    VTableSubclassPPB.Subclass ObjPtr(This), VTableIndexPPBGetDisplayString, VTAbleIndexPPBGetPredefinedValue, _
    AddressOf IPPB_GetDisplayString, 0, _
    AddressOf IPPB_GetPredefinedStrings, _
    AddressOf IPPB_GetPredefinedValue
End If
VTableSubclassPPB.AddRef
End Sub

Private Sub RestoreIPPB(ByVal This As OLEGuids.IPerPropertyBrowsing)
If Not VTableSubclassPPB Is Nothing Then
    VTableSubclassPPB.Release
    If VTableSubclassPPB.RefCount = 0 Then VTableSubclassPPB.UnSubclass
End If
End Sub

Public Function GetDispID(ByVal This As Object, ByRef MethodName As String) As Long
Dim IDispatch As OLEGuids.IDispatch
Dim IID_NULL As OLEGuids.OLECLSID
Set IDispatch = This
IDispatch.GetIDsOfNames IID_NULL, StrConv(MethodName, vbUnicode), 1, 0, GetDispID
End Function

Private Function IPPB_GetDisplayString(ByVal This As Object, ByVal DispID As Long, ByVal lpDisplayName As Long) As Long
Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB
Dim Handled As Boolean
On Error GoTo CATCH_EXCEPTION
If VarPtr(lpDisplayName) = 0 Then
    IPPB_GetDisplayString = E_POINTER
    Exit Function
End If
Dim DisplayName As String
Dim lpString As Long
Set ShadowIPerPropertyBrowsingVB = This
ShadowIPerPropertyBrowsingVB.GetDisplayString Handled, DispID, DisplayName
If Handled = False Then
    IPPB_GetDisplayString = E_NOTIMPL
Else
    lpString = SysAllocString(StrPtr(DisplayName))
    CopyMemory ByVal lpDisplayName, lpString, 4
End If
Exit Function
CATCH_EXCEPTION:
IPPB_GetDisplayString = E_NOTIMPL
End Function

Private Function IPPB_GetPredefinedStrings(ByVal This As Object, ByVal DispID As Long, ByRef pCaStringsOut As OLEGuids.OLECALPOLESTR, ByRef pCaCookiesOut As OLEGuids.OLECADWORD) As Long
Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB
Dim Handled As Boolean
On Error GoTo CATCH_EXCEPTION
If VarPtr(pCaStringsOut) = 0 Or VarPtr(pCaCookiesOut) = 0 Then
    IPPB_GetPredefinedStrings = E_POINTER
    Exit Function
End If
Dim cElems As Long, pElems As Long
Dim nElemCount As Integer
Dim lpString As Long
ReDim StringsOutArray(0) As String
ReDim CookiesOutArray(0) As Long
Set ShadowIPerPropertyBrowsingVB = This
ShadowIPerPropertyBrowsingVB.GetPredefinedStrings Handled, DispID, StringsOutArray(), CookiesOutArray()
If Handled = False Or UBound(StringsOutArray()) = 0 Then
    IPPB_GetPredefinedStrings = E_NOTIMPL
Else
    cElems = UBound(StringsOutArray())
    If Not UBound(CookiesOutArray()) = cElems Then ReDim Preserve CookiesOutArray(cElems) As Long
    pElems = CoTaskMemAlloc(cElems * 4)
    pCaStringsOut.cElems = cElems
    pCaStringsOut.pElems = pElems
    For nElemCount = 0 To cElems - 1
        lpString = CoTaskMemAlloc(Len(StringsOutArray(nElemCount)) + 1)
        CopyMemory ByVal lpString, StrPtr(StringsOutArray(nElemCount)), 4
        CopyMemory ByVal UnsignedAdd(pElems, nElemCount * 4), ByVal lpString, 4
    Next nElemCount
    pElems = CoTaskMemAlloc(cElems * 4)
    pCaCookiesOut.cElems = cElems
    pCaCookiesOut.pElems = pElems
    For nElemCount = 0 To cElems - 1
        CopyMemory ByVal UnsignedAdd(pElems, nElemCount * 4), CookiesOutArray(nElemCount), 4
    Next nElemCount
End If
Exit Function
CATCH_EXCEPTION:
IPPB_GetPredefinedStrings = E_NOTIMPL
End Function

Private Function IPPB_GetPredefinedValue(ByVal This As Object, ByVal DispID As Long, ByVal dwCookie As Long, ByRef pVarOut As Variant) As Long
Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB
Dim Handled As Boolean
On Error GoTo CATCH_EXCEPTION
If VarPtr(dwCookie) = 0 Or VarPtr(pVarOut) = 0 Then
    IPPB_GetPredefinedValue = E_POINTER
    Exit Function
End If
Set ShadowIPerPropertyBrowsingVB = This
ShadowIPerPropertyBrowsingVB.GetPredefinedValue Handled, DispID, dwCookie, pVarOut
If Handled = False Then IPPB_GetPredefinedValue = E_NOTIMPL
Exit Function
CATCH_EXCEPTION:
IPPB_GetPredefinedValue = E_NOTIMPL
End Function

Private Sub ReplaceIEnumeration(ByVal This As OLEGuids.IEnumeration)
If VTableSubclassEnumeration Is Nothing Then Set VTableSubclassEnumeration = New VTableSubclass
If VTableSubclassEnumeration.RefCount = 0 Then
    VTableSubclassEnumeration.Subclass ObjPtr(This), VTableIndexEnumerationNext, VTableIndexEnumerationClone, _
    AddressOf IEnumeration_Next, _
    AddressOf IEnumeration_Skip, _
    AddressOf IEnumeration_Reset, _
    AddressOf IEnumeration_Clone
End If
VTableSubclassEnumeration.AddRef
End Sub

Private Sub RestoreIEnumeration(ByVal This As OLEGuids.IEnumeration)
If Not VTableSubclassEnumeration Is Nothing Then
    VTableSubclassEnumeration.Release
    If VTableSubclassEnumeration.RefCount = 0 Then VTableSubclassEnumeration.UnSubclass
End If
End Sub

Private Function IEnumeration_Next(ByVal This As Object, ByVal VntCount As Long, ByRef VntArray As Variant, ByVal pcvFetched As Long) As Long
On Error GoTo CATCH_EXCEPTION
Dim ThisEnum As Enumeration
Dim liFetched As Long, NoMoreItems As Boolean, i As Long
Call InitSafeArray(VarPtr(VntArray), VntCount)
Set ThisEnum = This
For i = 0 To VntCount - 1
    ThisEnum.GetNextItem VariantArray(i), NoMoreItems
    If NoMoreItems = True Then Exit For
    liFetched = liFetched + 1
Next i
If liFetched = VntCount Then
    IEnumeration_Next = S_OK
Else
    IEnumeration_Next = S_FALSE
End If
If pcvFetched <> 0 Then CopyMemory ByVal pcvFetched, liFetched, 4
Call InitSafeArray(0, 0)
Exit Function
CATCH_EXCEPTION:
IEnumeration_Next = MapCOMErr(Err.Number)
For i = i To 0 Step -1
    VariantArray(i) = Empty
Next i
If pcvFetched <> 0 Then CopyMemory ByVal pcvFetched, 0&, 4
End Function

Private Function IEnumeration_Skip(ByVal This As Object, ByVal cV As Long) As Long
Dim ThisEnum As Enumeration
Dim SkippedAll As Boolean
On Error GoTo CATCH_EXCEPTION
Set ThisEnum = This
ThisEnum.Skip cV, SkippedAll
If SkippedAll = True Then IEnumeration_Skip = S_OK Else IEnumeration_Skip = S_FALSE
Exit Function
CATCH_EXCEPTION:
IEnumeration_Skip = MapCOMErr(Err.Number)
End Function

Private Function IEnumeration_Reset(ByVal This As Object) As Long
Dim ThisEnum As Enumeration
On Error GoTo CATCH_EXCEPTION
Set ThisEnum = This
ThisEnum.Reset
IEnumeration_Reset = S_OK
Exit Function
CATCH_EXCEPTION:
IEnumeration_Reset = MapCOMErr(Err.Number)
End Function

Private Function IEnumeration_Clone(ByVal This As Object, ByRef ppEnum As IEnumVARIANT) As Long
IEnumeration_Clone = E_NOTIMPL
End Function

Private Sub InitSafeArray(ByVal Addr As Long, ByVal cElt As Long)
Const FADF_STATIC As Long = &H2
Const FADF_FIXEDSIZE As Long = &H10
Const FADF_VARIANT As Long = &H800
With SAHeader
If .cDims = 0 Then
    .cbElements = 16
    .cDims = 1
    .fFeatures = FADF_STATIC Or FADF_FIXEDSIZE Or FADF_VARIANT
    CopyMemory ByVal ArrPtr(VariantArray), VarPtr(SAHeader), 4
End If
.Bounds(0).cElements = cElt + 1
.pvData = Addr
End With
End Sub

Private Function MapCOMErr(ByVal ErrNumber As Long) As Long
If ErrNumber <> 0 Then
    If (ErrNumber And &H80000000) Or (ErrNumber = 1) Then
        MapCOMErr = ErrNumber
    Else
        MapCOMErr = &H800A0000 Or ErrNumber
    End If
End If
End Function
