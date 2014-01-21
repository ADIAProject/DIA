Attribute VB_Name = "ComCtlsBase"
Option Explicit

#Const ImplementIDEStopProtection = True

#If False Then
Private OLEDropModeNone, OLEDropModeManual
Private CCAppearanceFlat, CCAppearance3D
Private CCBorderStyleNone, CCBorderStyleSingle, CCBorderStyleThin, CCBorderStyleSunken, CCBorderStyleRaised
Private CCBackStyleTransparent, CCBackStyleOpaque
Private CCLeftRightAlignmentLeft, CCLeftRightAlignmentRight
#End If
Public Enum OLEDropModeConstants
OLEDropModeNone = vbOLEDropNone
OLEDropModeManual = vbOLEDropManual
End Enum
Public Enum CCAppearanceConstants
CCAppearanceFlat = 0
CCAppearance3D = 1
End Enum
Public Enum CCBorderStyleConstants
CCBorderStyleNone = 0
CCBorderStyleSingle = 1
CCBorderStyleThin = 2
CCBorderStyleSunken = 3
CCBorderStyleRaised = 4
End Enum
Public Enum CCBackStyleConstants
CCBackStyleTransparent = 0
CCBackStyleOpaque = 1
End Enum
Public Enum CCLeftRightAlignmentConstants
CCLeftRightAlignmentLeft = 0
CCLeftRightAlignmentRight = 1
End Enum
Private Type DLLVERSIONINFO
cbSize As Long
dwMajor As Long
dwMinor As Long
dwBuildNumber As Long
dwPlatformID As Long
End Type
Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformID As Long
szCSDVersion(0 To ((128 * 2) - 1)) As Byte
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (ByRef lpVersionInfo As OSVERSIONINFO) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowSubclass_W2K Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass_W2K Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc_W2K Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByRef lpAddress As Long, ByVal dwSize As Long, ByVal flAllocType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByRef lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RELEASE As Long = &H8000&
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const GWL_WNDPROC As Long = (-4)
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_UAHDESTROYWINDOW As Long = &H90
Private ShellModHandle As Long, ShellModCount As Long

#If ImplementIDEStopProtection = True Then

Private Type IMAGE_DATA_DIRECTORY
VirtualAddress As Long
Size As Long
End Type
Private Type IMAGE_OPTIONAL_HEADER32
Magic As Integer
MajorLinkerVersion As Byte
MinorLinkerVersion As Byte
SizeOfCode As Long
SizeOfInitalizedData As Long
SizeOfUninitalizedData As Long
AddressOfEntryPoint As Long
BaseOfCode As Long
BaseOfData As Long
ImageBase As Long
SectionAlignment As Long
FileAlignment As Long
MajorOperatingSystemVer As Integer
MinorOperatingSystemVer As Integer
MajorImageVersion As Integer
MinorImageVersion As Integer
MajorSubsystemVersion As Integer
MinorSubsystemVersion As Integer
Reserved1 As Long
SizeOfImage As Long
SizeOfHeaders As Long
CheckSum As Long
Subsystem As Integer
DllCharacteristics As Integer
SizeOfStackReserve As Long
SizeOfStackCommit As Long
SizeOfHeapReserve As Long
SizeOfHeapCommit As Long
LoaderFlags As Long
NumberOfRvaAndSizes As Long
DataDirectory(15) As IMAGE_DATA_DIRECTORY
End Type
Private Type IMAGE_DOS_HEADER
e_magic As Integer
e_cblp As Integer
e_cp As Integer
e_crlc As Integer
e_cparhdr As Integer
e_minalloc As Integer
e_maxalloc As Integer
e_ss As Integer
e_sp As Integer
e_csum As Integer
e_ip As Integer
e_cs As Integer
e_lfarlc As Integer
e_onvo As Integer
e_res(0 To 3) As Integer
e_oemid As Integer
e_oeminfo As Integer
e_res2(0 To 9) As Integer
e_lfanew As Long
End Type

#End If

Public Sub ComCtlsLoadShellMod()
If (ShellModHandle Or ShellModCount) = 0 Then ShellModHandle = LoadLibrary(StrPtr("Shell32.dll"))
ShellModCount = ShellModCount + 1
End Sub

Public Sub ComCtlsReleaseShellMod()
ShellModCount = ShellModCount - 1
If ShellModCount = 0 And ShellModHandle <> 0 Then
    FreeLibrary ShellModHandle
    ShellModHandle = 0
End If
End Sub

Public Sub ComCtlsShowAllUIStates(ByVal hWnd As Long)
Const WM_CHANGEUISTATE As Long = &H127
Const UIS_CLEAR As Long = 2, UISF_HIDEFOCUS As Long = &H1, UISF_HIDEACCEL As Long = &H2
SendMessage hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEFOCUS Or UISF_HIDEACCEL), ByVal 0&
End Sub

Public Sub ComCtlsChangeBorderStyle(ByVal hWnd As Long, ByVal Value As CCBorderStyleConstants)
Const WS_BORDER As Long = &H800000, WS_DLGFRAME As Long = &H400000
Const WS_EX_CLIENTEDGE As Long = &H200, WS_EX_STATICEDGE As Long = &H20000, WS_EX_WINDOWEDGE As Long = &H100
Dim dwStyle As Long, dwExStyle As Long
dwStyle = GetWindowLong(hWnd, GWL_STYLE)
dwExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
If (dwStyle And WS_BORDER) = WS_BORDER Then dwStyle = dwStyle And Not WS_BORDER
If (dwStyle And WS_DLGFRAME) = WS_DLGFRAME Then dwStyle = dwStyle And Not WS_DLGFRAME
If (dwExStyle And WS_EX_STATICEDGE) = WS_EX_STATICEDGE Then dwExStyle = dwExStyle And Not WS_EX_STATICEDGE
If (dwExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
If (dwExStyle And WS_EX_WINDOWEDGE) = WS_EX_WINDOWEDGE Then dwExStyle = dwExStyle And Not WS_EX_WINDOWEDGE
Select Case Value
    Case CCBorderStyleSingle
        dwStyle = dwStyle Or WS_BORDER
    Case CCBorderStyleThin
        dwExStyle = dwExStyle Or WS_EX_STATICEDGE
    Case CCBorderStyleSunken
        dwExStyle = dwExStyle Or WS_EX_CLIENTEDGE
    Case CCBorderStyleRaised
        dwExStyle = dwExStyle Or WS_EX_WINDOWEDGE
        dwStyle = dwStyle Or WS_DLGFRAME
End Select
SetWindowLong hWnd, GWL_STYLE, dwStyle
SetWindowLong hWnd, GWL_EXSTYLE, dwExStyle
Const SWP_FRAMECHANGED As Long = &H20, SWP_NOMOVE As Long = &H2, SWP_NOOWNERZORDER As Long = &H200, SWP_NOSIZE As Long = &H1, SWP_NOZORDER As Long = &H4
SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Public Sub ComCtlsMousePointerSetDisplayString(ByVal MousePointer As Integer, ByRef DisplayName As String)
Select Case MousePointer
    Case 0: DisplayName = "0 - Default"
    Case 1: DisplayName = "1 - Arrow"
    Case 2: DisplayName = "2 - Cross"
    Case 3: DisplayName = "3 - I-Beam"
    Case 4: DisplayName = "4 - Hand"
    Case 5: DisplayName = "5 - Size"
    Case 6: DisplayName = "6 - Size NE SW"
    Case 7: DisplayName = "7 - Size N S"
    Case 8: DisplayName = "8 - Size NW SE"
    Case 9: DisplayName = "9 - Size W E"
    Case 10: DisplayName = "10 - Up Arrow"
    Case 11: DisplayName = "11 - Hourglass"
    Case 12: DisplayName = "12 - No Drop"
    Case 13: DisplayName = "13 - Arrow and Hourglass"
    Case 14: DisplayName = "14 - Arrow and Question"
    Case 15: DisplayName = "15 - Size All"
    Case 16: DisplayName = "16 - Arrow and CD"
    Case 99: DisplayName = "99 - Custom"
End Select
End Sub

Public Sub ComCtlsMousePointerSetPredefinedStrings(ByRef StringsOut() As String, ByRef CookiesOut() As Long)
ReDim StringsOut(0 To (17 + 1)) As String
ReDim CookiesOut(0 To (17 + 1)) As Long
StringsOut(0) = "0 - Default": CookiesOut(0) = 0
StringsOut(1) = "1 - Arrow": CookiesOut(1) = 1
StringsOut(2) = "2 - Cross": CookiesOut(2) = 2
StringsOut(3) = "3 - I-Beam": CookiesOut(3) = 3
StringsOut(4) = "4 - Hand": CookiesOut(4) = 4
StringsOut(5) = "5 - Size": CookiesOut(5) = 5
StringsOut(6) = "6 - Size NE SW": CookiesOut(6) = 6
StringsOut(7) = "7 - Size N S": CookiesOut(7) = 7
StringsOut(8) = "8 - Size NW SE": CookiesOut(8) = 8
StringsOut(9) = "9 - Size W E": CookiesOut(9) = 9
StringsOut(10) = "10 - Up Arrow": CookiesOut(10) = 10
StringsOut(11) = "11 - Hourglass": CookiesOut(11) = 11
StringsOut(12) = "12 - No Drop": CookiesOut(12) = 12
StringsOut(13) = "13 - Arrow and Hourglass": CookiesOut(13) = 13
StringsOut(14) = "14 - Arrow and Question": CookiesOut(14) = 14
StringsOut(15) = "15 - Size All": CookiesOut(15) = 15
StringsOut(16) = "16 - Arrow and CD": CookiesOut(16) = 16
StringsOut(17) = "99 - Custom": CookiesOut(17) = 99
End Sub

Public Sub ComCtlsMousePointerPPInitCombo(ByRef ComboMousePointer As VB.ComboBox)
With ComboMousePointer
.AddItem "0 - Default"
.ItemData(.NewIndex) = 0
.AddItem "1 - Arrow"
.ItemData(.NewIndex) = 1
.AddItem "2 - Cross"
.ItemData(.NewIndex) = 2
.AddItem "3 - I-Beam"
.ItemData(.NewIndex) = 3
.AddItem "4 - Hand"
.ItemData(.NewIndex) = 4
.AddItem "5 - Size"
.ItemData(.NewIndex) = 5
.AddItem "6 - Size NE SW"
.ItemData(.NewIndex) = 6
.AddItem "7 - Size N S"
.ItemData(.NewIndex) = 7
.AddItem "8 - Size NW SE"
.ItemData(.NewIndex) = 8
.AddItem "9 - Size W E"
.ItemData(.NewIndex) = 9
.AddItem "10 - Up Arrow"
.ItemData(.NewIndex) = 10
.AddItem "11 - Hourglass"
.ItemData(.NewIndex) = 11
.AddItem "12 - No Drop"
.ItemData(.NewIndex) = 12
.AddItem "13 - Arrow and Hourglass"
.ItemData(.NewIndex) = 13
.AddItem "14 - Arrow and Question"
.ItemData(.NewIndex) = 14
.AddItem "15 - Size All"
.ItemData(.NewIndex) = 15
.AddItem "16 - Arrow and CD"
.ItemData(.NewIndex) = 16
.AddItem "99 - Custom"
.ItemData(.NewIndex) = 99
End With
End Sub

Public Function ComCtlsSupportLevel() As Byte
Static Done As Boolean, Value As Byte
If Done = False Then
    Dim Version As DLLVERSIONINFO
    On Error Resume Next
    Version.cbSize = LenB(Version)
    Const S_OK As Long = &H0
    If DllGetVersion(Version) = S_OK Then
        If Version.dwMajor = 6 And Version.dwMinor = 0 Then
            Value = 1
        ElseIf Version.dwMajor > 6 Or (Version.dwMajor = 6 And Version.dwMinor > 0) Then
            Value = 2
        End If
    End If
    Done = True
End If
ComCtlsSupportLevel = Value
End Function

Public Function ComCtlsW2KCompatibility() As Boolean
Static Done As Boolean, Value As Boolean
If Done = False Then
    Dim Version As OSVERSIONINFO
    On Error Resume Next
    Version.dwOSVersionInfoSize = LenB(Version)
    If GetVersionEx(Version) <> 0 Then
        With Version
        Const VER_PLATFORM_WIN32_NT As Long = 2
        If .dwPlatformID = VER_PLATFORM_WIN32_NT Then
            If .dwMajorVersion = 5 And .dwMinorVersion = 0 Then Value = True
        End If
        End With
    End If
    Done = True
End If
ComCtlsW2KCompatibility = Value
End Function

Public Sub ComCtlsSetSubclass(ByVal hWnd As Long, ByVal This As ISubclass, ByVal dwRefData As Long, Optional ByVal Name As String)
If hWnd = 0 Then Exit Sub
If Name = vbNullString Then Name = "ComCtl"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 0 Then
    If ComCtlsW2KCompatibility() = False Then
        SetWindowSubclass hWnd, AddressOf ComCtlsSubclassProc, ObjPtr(This), dwRefData
    Else
        SetWindowSubclass_W2K hWnd, AddressOf ComCtlsSubclassProc, ObjPtr(This), dwRefData
    End If
    SetProp hWnd, StrPtr(Name & "SubclassID"), ObjPtr(This)
    SetProp hWnd, StrPtr(Name & "SubclassInit"), 1
End If
End Sub

Public Function ComCtlsDefaultProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If ComCtlsW2KCompatibility() = False Then
    ComCtlsDefaultProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
Else
    ComCtlsDefaultProc = DefSubclassProc_W2K(hWnd, wMsg, wParam, lParam)
End If
End Function

Public Sub ComCtlsRemoveSubclass(ByVal hWnd As Long, Optional ByVal Name As String)
If hWnd = 0 Then Exit Sub
If Name = vbNullString Then Name = "ComCtl"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 1 Then
    If ComCtlsW2KCompatibility() = False Then
        RemoveWindowSubclass hWnd, AddressOf ComCtlsSubclassProc, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    Else
        RemoveWindowSubclass_W2K hWnd, AddressOf ComCtlsSubclassProc, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    End If
    RemoveProp hWnd, StrPtr(Name & "SubclassID")
    RemoveProp hWnd, StrPtr(Name & "SubclassInit")
End If
End Sub

Public Function ComCtlsSubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Select Case wMsg
    Case WM_DESTROY
        ComCtlsSubclassProc = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    Case WM_NCDESTROY, WM_UAHDESTROYWINDOW
        ComCtlsSubclassProc = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        If ComCtlsW2KCompatibility() = False Then
            RemoveWindowSubclass hWnd, AddressOf ComCtlsBase.ComCtlsSubclassProc, uIdSubclass
        Else
            RemoveWindowSubclass_W2K hWnd, AddressOf ComCtlsBase.ComCtlsSubclassProc, uIdSubclass
        End If
        Exit Function
End Select
On Error Resume Next
Dim This As ISubclass
Set This = PtrToObj(uIdSubclass)
If Err.Number = 0 Then
    ComCtlsSubclassProc = This.Message(hWnd, wMsg, wParam, lParam, dwRefData)
Else
    ComCtlsSubclassProc = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End If
End Function

Public Sub ComCtlsSetDesignModeSubclass(ByVal hWnd As Long, ByVal This As Object, ByVal Ordinal As Byte, ByRef ASMWrapper As Long, ByRef PrevWndProc As Long)
If ASMWrapper <> 0 Then Exit Sub
ASMWrapper = VirtualAlloc(ByVal 0, 105, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
If ASMWrapper = 0 Then Exit Sub
Dim ASM(0 To 104) As Byte, VirtualFreePointer As Long
VirtualFreePointer = GetProcAddress(GetModuleHandle(StrPtr("kernel32.dll")), "VirtualFree")
ASM(0) = &H90: ASM(1) = &HFF: ASM(2) = &H5: ASM(7) = &H6A: ASM(8) = &H0: ASM(9) = &H54
ASM(10) = &HFF: ASM(11) = &H74: ASM(12) = &H24: ASM(13) = &H18: ASM(14) = &HFF: ASM(15) = &H74
ASM(16) = &H24: ASM(17) = &H18: ASM(18) = &HFF: ASM(19) = &H74: ASM(20) = &H24: ASM(21) = &H18
ASM(22) = &HFF: ASM(23) = &H74: ASM(24) = &H24: ASM(25) = &H18: ASM(26) = &H68: ASM(31) = &HB8
ASM(36) = &HFF: ASM(37) = &HD0: ASM(38) = &HFF: ASM(39) = &HD: ASM(44) = &HA1: ASM(49) = &H85
ASM(50) = &HC0: ASM(51) = &H75: ASM(52) = &H4: ASM(53) = &H58: ASM(54) = &HC2: ASM(55) = &H10
ASM(56) = &H0: ASM(57) = &HA1: ASM(62) = &H85: ASM(63) = &HC0: ASM(64) = &H74: ASM(65) = &H4
ASM(66) = &H58: ASM(67) = &HC2: ASM(68) = &H10: ASM(69) = &H0: ASM(70) = &H58: ASM(71) = &H59
ASM(72) = &H58: ASM(73) = &H58: ASM(74) = &H58: ASM(75) = &H58: ASM(76) = &H68: ASM(77) = &H0
ASM(78) = &H80: ASM(79) = &H0: ASM(80) = &H0: ASM(81) = &H6A: ASM(82) = &H0: ASM(83) = &H68
ASM(88) = &H51: ASM(89) = &HB8: ASM(94) = &HFF: ASM(95) = &HE0: ASM(96) = &H0: ASM(97) = &H0
ASM(98) = &H0: ASM(99) = &H0: ASM(100) = &H0: ASM(101) = &H0: ASM(102) = &H0: ASM(103) = &H0
CopyMemory ASM(3), UnsignedAdd(ASMWrapper, 96), 4
CopyMemory ASM(40), UnsignedAdd(ASMWrapper, 96), 4
CopyMemory ASM(58), UnsignedAdd(ASMWrapper, 96), 4
CopyMemory ASM(45), UnsignedAdd(ASMWrapper, 100), 4
CopyMemory ASM(84), ASMWrapper, 4
CopyMemory ASM(27), ObjPtr(This), 4
CopyMemory ASM(32), SelfAddressOf(This, Ordinal), 4
CopyMemory ASM(90), VirtualFreePointer, 4
CopyMemory ByVal ASMWrapper, ASM(0), 105
PrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, ASMWrapper)
End Sub

Public Sub ComCtlsRemoveDesignModeSubclass(ByVal hWnd As Long, ByRef ASMWrapper As Long, ByRef PrevWndProc As Long)
If ASMWrapper = 0 Or PrevWndProc = 0 Or hWnd = 0 Then Exit Sub
SetWindowLong hWnd, GWL_WNDPROC, PrevWndProc
PrevWndProc = 0
Dim Counter As Long
CopyMemory ByVal VarPtr(Counter), ByVal UnsignedAdd(ASMWrapper, 96), 4
If Counter = 0 Then
    VirtualFree ByVal ASMWrapper, 0, MEM_RELEASE
Else
    CopyMemory ByVal UnsignedAdd(ASMWrapper, 100), 1&, 4
End If
ASMWrapper = 0
End Sub

Public Function LvwSortingFunctionBinary(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
LvwSortingFunctionBinary = This.Message(0, 0, lParam1, lParam2, 10)
End Function

Public Function LvwSortingFunctionText(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
LvwSortingFunctionText = This.Message(0, 0, lParam1, lParam2, 11)
End Function

Public Function LvwSortingFunctionNumeric(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
LvwSortingFunctionNumeric = This.Message(0, 0, lParam1, lParam2, 12)
End Function

Public Function LvwSortingFunctionCurrency(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
LvwSortingFunctionCurrency = This.Message(0, 0, lParam1, lParam2, 13)
End Function

Public Function LvwSortingFunctionDate(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
LvwSortingFunctionDate = This.Message(0, 0, lParam1, lParam2, 14)
End Function

Public Function TvwSortingFunctionBinary(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
TvwSortingFunctionBinary = This.Message(0, 0, lParam1, lParam2, 10)
End Function

Public Function TvwSortingFunctionText(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal This As ISubclass) As Long
TvwSortingFunctionText = This.Message(0, 0, lParam1, lParam2, 11)
End Function

Public Function CdlShowFolderCallback(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const BFFM_INITIALIZED As Long = 1
If wMsg = BFFM_INITIALIZED Then
    Const WM_USER As Long = &H400
    Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
    Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
    Const BFFM_SETSELECTION As Long = BFFM_SETSELECTIONW
    If lParam <> 0 Then SendMessage hWnd, BFFM_SETSELECTION, 1, ByVal lParam
End If
CdlShowFolderCallback = 0
End Function

Public Sub ComCtlsInitIDEStopProtection()

#If ImplementIDEStopProtection = True Then

If InIDE() = True Then
    Dim ASMWrapper As Long, RestorePointer As Long, OldAddress As Long
    ASMWrapper = VirtualAlloc(ByVal 0, 20, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    OldAddress = GetProcAddress(GetModuleHandle(StrPtr("vba6.dll")), "EbProjectReset")
    RestorePointer = HookIATEntry("vb6.exe", "vba6.dll", "EbProjectReset", ASMWrapper)
    WriteCall ASMWrapper, AddressOf IDEStopProtectionHandler
    WriteByte ASMWrapper, &HC7 ' MOV
    WriteByte ASMWrapper, &H5
    WriteLong ASMWrapper, RestorePointer ' IAT Entry
    WriteLong ASMWrapper, OldAddress ' Address from EbProjectReset
    WriteJump ASMWrapper, OldAddress
End If

#End If

End Sub

#If ImplementIDEStopProtection = True Then

Private Sub IDEStopProtectionHandler()
On Error Resume Next
Call RemoveAllVTableSubclass(VTableInterfaceInPlaceActiveObject)
Call RemoveAllVTableSubclass(VTableInterfaceControl)
Call RemoveAllVTableSubclass(VTableInterfacePerPropertyBrowsing)
Call RemoveAllVTableSubclass(VTableInterfaceEnumeration)
Dim AppForm As Form, CurrControl As Control
For Each AppForm In Forms
    For Each CurrControl In AppForm.Controls
        Select Case TypeName(CurrControl)
            Case "Animation", "DTPicker", "MonthView", "Slider", "TabStrip", "ListView", "TreeView", "IPAddress", "ToolBar", "UpDown", "SpinBox", "Pager", "OptionButtonW", "CheckBoxW", "CommandButtonW", "TextBoxW", "HotKey", "CoolBar"
                Call ComCtlsRemoveSubclass(CurrControl.hWnd)
                Call ComCtlsRemoveSubclass(CurrControl.hWndUserControl)
            Case "ProgressBar", "FrameW"
                Call ComCtlsRemoveSubclass(CurrControl.hWnd)
            Case "StatusBar"
                Call ComCtlsRemoveSubclass(CurrControl.hWnd)
                Call ComCtlsRemoveSubclass(CurrControl.hWndUserControl)
                Call ComCtlsRemoveSubclass(AppForm.hWnd, ProperControlName(CurrControl))
            Case "ToolTip"
                Call ComCtlsRemoveSubclass(AppForm.hWnd, ProperControlName(CurrControl))
            Case "ImageCombo"
                Call ComCtlsRemoveSubclass(CurrControl.hWnd)
                Call ComCtlsRemoveSubclass(CurrControl.hWndCombo)
                If CurrControl.hWndEdit <> 0 Then Call ComCtlsRemoveSubclass(CurrControl.hWndEdit)
                Call ComCtlsRemoveSubclass(CurrControl.hWndUserControl)
            Case "RichTextBox", "MCIWnd"
                CurrControl.IDEStop ' Hidden
        End Select
    Next CurrControl
Next AppForm
End Sub

Private Function HookIATEntry(ByVal Module As String, ByVal Lib As String, ByVal Fnc As String, ByVal NewAddr As Long) As Long
Dim hMod As Long, OldLibFncAddr As Long
Dim lpIAT As Long, IATLen As Long, IATPos As Long
Dim DOSHdr As IMAGE_DOS_HEADER
Dim PEHdr As IMAGE_OPTIONAL_HEADER32
hMod = GetModuleHandle(StrPtr(Module))
If hMod = 0 Then Exit Function
OldLibFncAddr = GetProcAddress(GetModuleHandle(StrPtr(Lib)), Fnc)
If OldLibFncAddr = 0 Then Exit Function
CopyMemory DOSHdr, ByVal hMod, LenB(DOSHdr)
CopyMemory PEHdr, ByVal UnsignedAdd(hMod, DOSHdr.e_lfanew), LenB(PEHdr)
Const IMAGE_NT_SIGNATURE As Long = &H4550
If PEHdr.Magic = IMAGE_NT_SIGNATURE Then
    lpIAT = PEHdr.DataDirectory(15).VirtualAddress + hMod
    IATLen = PEHdr.DataDirectory(15).Size
    IATPos = lpIAT
    Do Until CLongToULong(IATPos) >= CLongToULong(UnsignedAdd(lpIAT, IATLen))
        If DeRef(IATPos) = OldLibFncAddr Then
            VirtualProtect IATPos, 4, PAGE_EXECUTE_READWRITE, 0
            CopyMemory ByVal IATPos, NewAddr, 4
            HookIATEntry = IATPos
            Exit Do
        End If
        IATPos = UnsignedAdd(IATPos, 4)
    Loop
End If
End Function

Private Function DeRef(ByVal Addr As Long) As Long
CopyMemory DeRef, ByVal Addr, 4
End Function

Private Sub WriteJump(ByRef ASM As Long, ByRef Addr As Long)
WriteByte ASM, &HE9
WriteLong ASM, Addr - ASM - 4
End Sub

Private Sub WriteCall(ByRef ASM As Long, ByRef Addr As Long)
WriteByte ASM, &HE8
WriteLong ASM, Addr - ASM - 4
End Sub

Private Sub WriteLong(ByRef ASM As Long, ByRef Lng As Long)
CopyMemory ByVal ASM, Lng, 4
ASM = ASM + 4
End Sub

Private Sub WriteByte(ByRef ASM As Long, ByRef B As Byte)
CopyMemory ByVal ASM, B, 1
ASM = ASM + 1
End Sub

#End If
