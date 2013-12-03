Attribute VB_Name = "Common"
Option Explicit
Private Type MSGBOXPARAMS
cbSize As Long
hWndOwner As Long
hInstance As Long
lpszText As Long
lpszCaption As Long
dwStyle As Long
lpszIcon As Long
dwContextHelpID As Long
lpfnMsgBoxCallback As Long
dwLanguageId As Long
End Type
Private Type BITMAP
BMType As Long
BMWidth As Long
BMHeight As Long
BMWidthBytes As Long
BMPlanes As Integer
BMBitsPixel As Integer
BMBits As Long
End Type
Private Type ICONINFO
fIcon As Long
XHotspot As Long
YHotspot As Long
hBMMask As Long
hBMColor As Long
End Type
Private Type PICTDESC
cbSizeOfStruct As Long
PicType As Long
hImage As Long
XExt As Long
YExt As Long
End Type
Private Type CLSID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectW" (ByRef lpMsgBoxParams As MSGBOXPARAMS) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, ByVal lpCreationTime As Long, ByVal lpLastAccessTime As Long, ByVal lpLastWriteTime As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (ByVal lpFileTime As Long, ByVal lpLocalFileTime As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (ByVal lpFileTime As Long, ByVal lpSystemTime As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal NumberOfBytesToRead As Long, ByRef NumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal VKey As Long) As Integer
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal hRegion As Long, ByVal RDW_RedrawFlags As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal XLeft As Long, ByVal YTop As Long, ByVal hIcon As Long, ByVal CXWidth As Long, ByVal CYWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, ByRef pIconInfo As ICONINFO) As Long
Private Declare Function CreateIconIndirect Lib "user32" (ByRef pIconInfo As ICONINFO) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal Color As Long, ByVal hPal As Long, ByRef ColorRef As Long) As Long
Private Declare Function OleLoadPicture Lib "oleaut32" (ByVal lpStream As IUnknown, ByVal lSize As Long, ByVal fRunmode As Long, ByRef riid As Any, ByRef lpIPicture As IPicture) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (ByRef lpPictDesc As PICTDESC, ByRef riid As Any, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPictureDisp) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCW" (ByVal lpDriverName As Long, ByRef lpDeviceName As Any, ByRef lpOutput As Any, ByRef lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

' (VB-Overwrite)
Public Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String) As VbMsgBoxResult
Dim MSGBOXP As MSGBOXPARAMS
With MSGBOXP
.cbSize = LenB(MSGBOXP)
If (Buttons And vbSystemModal) = 0 Then
    If Not Screen.ActiveForm Is Nothing Then
        .hWndOwner = Screen.ActiveForm.hWnd
    Else
        .hWndOwner = GetActiveWindow()
    End If
Else
    .hWndOwner = GetForegroundWindow()
End If
.hInstance = App.hInstance
.lpszText = StrPtr(Prompt)
If Title = vbNullString Then Title = App.Title
.lpszCaption = StrPtr(Title)
.dwStyle = Buttons
End With
MsgBox = MessageBoxIndirect(MSGBOXP)
End Function

' (VB-Overwrite)
Public Function GetAttr(ByVal PathName As String) As VbFileAttribute
Const INVALID_FILE_ATTRIBUTES As Long = (-1)
Const FILE_ATTRIBUTE_NORMAL As Long = &H80
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
Dim dwAttributes As Long
dwAttributes = GetFileAttributes(StrPtr("\\?\" & PathName))
If dwAttributes = INVALID_FILE_ATTRIBUTES Then
    Err.Raise 53
ElseIf dwAttributes = FILE_ATTRIBUTE_NORMAL Then
    GetAttr = vbNormal
Else
    GetAttr = dwAttributes
End If
End Function

' (VB-Overwrite)
Public Sub SetAttr(ByVal PathName As String, ByVal Attributes As VbFileAttribute)
Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Dim dwAttributes As Long
If Attributes = vbNormal Then
    dwAttributes = FILE_ATTRIBUTE_NORMAL
Else
    If (Attributes And (vbVolume Or vbDirectory Or vbAlias)) <> 0 Then Err.Raise 5
    dwAttributes = Attributes
End If
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
If SetFileAttributes(StrPtr("\\?\" & PathName), dwAttributes) = 0 Then Err.Raise 53
End Sub

' (VB-Overwrite)
Public Function FileLen(ByVal PathName As String) As Long
Const INVALID_HANDLE_VALUE As Long = (-1)
Const GENERIC_READ As Long = &H80000000, FILE_SHARE_READ As Long = &H1, OPEN_EXISTING As Long = 3, FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Dim hFile As Long, Length As Double
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
hFile = CreateFile(StrPtr("\\?\" & PathName), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
If hFile <> INVALID_HANDLE_VALUE Then
    Length = GetFileSize(hFile, 0)
    CloseHandle hFile
Else
    Err.Raise Number:=53, Description:="File not found: '" & PathName & "'"
End If
End Function

' (VB-Overwrite)
Public Function FileDateTime(ByVal PathName As String) As Date
Const INVALID_HANDLE_VALUE As Long = (-1)
Const GENERIC_READ As Long = &H80000000, FILE_SHARE_READ As Long = &H1, OPEN_EXISTING As Long = 3, FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Dim hFile As Long, Length As Double
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
hFile = CreateFile(StrPtr("\\?\" & PathName), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
If hFile <> INVALID_HANDLE_VALUE Then
    Dim FT(0 To 1) As FILETIME, ST As SYSTEMTIME
    GetFileTime hFile, 0, 0, VarPtr(FT(0))
    FileTimeToLocalFileTime VarPtr(FT(0)), VarPtr(FT(1))
    FileTimeToSystemTime VarPtr(FT(1)), VarPtr(ST)
    FileDateTime = DateSerial(ST.wYear, ST.wMonth, ST.wDay) + TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
    CloseHandle hFile
Else
    Err.Raise Number:=53, Description:="File not found: '" & PathName & "'"
End If
End Function

Public Function FileExists(ByVal PathName As String) As Boolean
On Error Resume Next
Dim Attributes As VbFileAttribute, ErrVal As Long
Attributes = GetAttr(PathName)
ErrVal = Err.Number
On Error GoTo 0
If (Attributes And (vbDirectory Or vbVolume)) = 0 And ErrVal = 0 Then FileExists = True
End Function

Public Function LoadFile(ByVal PathName As String) As Variant
Const FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Const INVALID_HANDLE_VALUE As Long = (-1)
Const GENERIC_READ As Long = &H80000000
Const FILE_SHARE_READ As Long = &H1
Const OPEN_EXISTING As Long = 3
Dim hFile As Long, Length As Double
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
hFile = CreateFile(StrPtr("\\?\" & PathName), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
If hFile <> INVALID_HANDLE_VALUE Then
    Length = GetFileSize(hFile, 0) ' File size >= 2^31 not supported.
    If Length > 0 Then
        Dim B() As Byte
        ReDim B(0 To Length - 1) As Byte
        ReadFile hFile, VarPtr(B(0)), Length, 0, 0
        LoadFile = B()
    End If
    CloseHandle hFile
Else
    Err.Raise Number:=53, Description:="File not found: '" & PathName & "'"
End If
End Function

Public Function ApplicationPath() As String
ApplicationPath = App.Path & IIf(Right$(App.Path, 1) = "\", "", "\")
End Function

Public Function AccelCharCode(ByVal Caption As String) As Integer
If Caption = vbNullString Then Exit Function
Dim Pos As Long, Length As Long
Length = Len(Caption)
Pos = Length
Do
    If Mid$(Caption, Pos, 1) = "&" And Pos < Length Then
        AccelCharCode = Asc(UCase$(Mid$(Caption, Pos + 1, 1)))
        If Pos > 1 Then
            If Mid$(Caption, Pos - 1, 1) = "&" Then AccelCharCode = 0
        Else
            If AccelCharCode = vbKeyUp Then AccelCharCode = 0
        End If
        If AccelCharCode <> 0 Then Exit Do
    End If
    Pos = Pos - 1
Loop Until Pos = 0
End Function

Public Function ProperControlName(ByVal Control As VB.Control) As String
Dim Index As Long
On Error Resume Next
Index = Control.Index
If Err.Number <> 0 Or Index < 0 Then ProperControlName = Control.Name Else ProperControlName = Control.Name & "(" & Index & ")"
On Error GoTo 0
End Function

Public Function MousePointerID(ByVal MousePointer As Integer) As Long
Select Case MousePointer
    Case vbArrow
        Const IDC_ARROW As Long = 32512
        MousePointerID = IDC_ARROW
    Case vbCrosshair
        Const IDC_CROSS As Long = 32515
        MousePointerID = IDC_CROSS
    Case vbIbeam
        Const IDC_IBEAM As Long = 32513
        MousePointerID = IDC_IBEAM
    Case vbIconPointer ' Obselete, replaced Icon with Hand
        Const IDC_HAND As Long = 32649
        MousePointerID = IDC_HAND
    Case vbSizePointer, vbSizeAll
        Const IDC_SIZEALL As Long = 32646
        MousePointerID = IDC_SIZEALL
    Case vbSizeNESW
        Const IDC_SIZENESW As Long = 32643
        MousePointerID = IDC_SIZENESW
    Case vbSizeNS
        Const IDC_SIZENS As Long = 32645
        MousePointerID = IDC_SIZENS
    Case vbSizeNWSE
        Const IDC_SIZENWSE As Long = 32642
        MousePointerID = IDC_SIZENWSE
    Case vbSizeWE
        Const IDC_SIZEWE As Long = 32644
        MousePointerID = IDC_SIZEWE
    Case vbUpArrow
        Const IDC_UPARROW As Long = 32516
        MousePointerID = IDC_UPARROW
    Case vbHourglass
        Const IDC_WAIT As Long = 32514
        MousePointerID = IDC_WAIT
    Case vbNoDrop
        Const IDC_NO As Long = 32648
        MousePointerID = IDC_NO
    Case vbArrowHourglass
        Const IDC_APPSTARTING As Long = 32650
        MousePointerID = IDC_APPSTARTING
    Case vbArrowQuestion
        Const IDC_HELP As Long = 32651
        MousePointerID = IDC_HELP
    Case 16
        Const IDC_WAITCD As Long = 32663 ' Undocumented
        MousePointerID = IDC_WAITCD
End Select
End Function

Public Function GetShiftState() As ShiftConstants
GetShiftState = (-vbShiftMask * KeyPressed(vbKeyShift))
GetShiftState = GetShiftState Or (-vbAltMask * KeyPressed(vbKeyMenu))
GetShiftState = GetShiftState Or (-vbCtrlMask * KeyPressed(vbKeyControl))
End Function

Public Function GetMouseState() As MouseButtonConstants
GetMouseState = (-vbLeftButton * KeyPressed(vbLeftButton))
GetMouseState = GetMouseState Or (-vbMiddleButton * KeyPressed(vbMiddleButton))
GetMouseState = GetMouseState Or (-vbRightButton * KeyPressed(vbRightButton))
End Function

Public Function KeyToggled(ByVal VirtKeyCode As KeyCodeConstants) As Boolean
KeyToggled = CBool(LoByte(GetKeyState(VirtKeyCode)) = 1)
End Function
 
Public Function KeyPressed(ByVal VirtKeyCode As KeyCodeConstants) As Boolean
KeyPressed = CBool((GetAsyncKeyState(VirtKeyCode) And &H8000&) = &H8000&)
End Function

Public Function InIDE(Optional ByRef B As Boolean = True) As Boolean
If B = True Then Debug.Assert Not InIDE(InIDE) Else B = True
End Function

Public Function PtrToObj(ByVal ObjectPointer As Long) As Object
Dim TempObj As Object
CopyMemory TempObj, ObjectPointer, 4
Set PtrToObj = TempObj
CopyMemory TempObj, 0&, 4
End Function

Public Function ProcPtr(ByVal Address As Long) As Long
ProcPtr = Address
End Function

Public Function SelfAddressOf(ByVal This As Object, ByVal Ordinal As Byte) As Long
If This Is Nothing Or Not Ordinal > 0 Then Exit Function
Dim ByteSub As Byte, ByteValue As Byte
Dim Address As Long, i As Long, j As Long
CopyMemory ByVal VarPtr(Address), ByVal ObjPtr(This), 4
If ObjProbe(Address + &H7A4, i, ByteSub) = False Then ' Probe for a UserControl
    If ObjProbe(Address + &H1C, i, ByteSub) = False Then    ' Probe for a Class
        If ObjProbe(Address + &H6F8, i, ByteSub) = False Then ' Probe for a Form
            If ObjProbe(Address + &H710, i, ByteSub) = False Then ' Probe for a PropertyPage
                Exit Function
            End If
        End If
    End If
End If
i = i + 4
j = i + 2048
Do While i < j
    CopyMemory ByVal VarPtr(Address), ByVal i, 4
    If IsBadCodePtr(Address) <> 0 Then
        CopyMemory ByVal VarPtr(SelfAddressOf), ByVal i - (Ordinal * 4), 4
        Exit Do
    End If
    CopyMemory ByVal VarPtr(ByteValue), ByVal Address, 1
    If ByteValue <> ByteSub Then
        CopyMemory ByVal VarPtr(SelfAddressOf), ByVal i - (Ordinal * 4), 4
        Exit Do
    End If
    i = i + 4
Loop
End Function

Private Function ObjProbe(ByVal Start As Long, ByRef Method As Long, ByRef ByteSub As Byte) As Boolean
Dim ByteValue As Byte
Dim Address As Long
Dim Limit As Long, Entry As Long
Address = Start
Limit = Address + 64
Do While Address < Limit
    CopyMemory ByVal VarPtr(Entry), ByVal Address, 4
    If Entry <> 0 Then
        CopyMemory ByVal VarPtr(ByteValue), ByVal Entry, 1
        If ByteValue = &H33 Or ByteValue = &HE9 Then
            Method = Address
            ByteSub = ByteValue
            ObjProbe = True
            Exit Do
        End If
    End If
    Address = Address + 4
Loop
End Function

Public Function LoByte(ByVal Word As Integer) As Byte
LoByte = Word And &HFF
End Function

Public Function HiByte(ByVal Word As Integer) As Byte
HiByte = (Word And &HFF00&) \ &H100
End Function

Public Function MakeWord(ByVal LoByte As Byte, ByVal HiByte As Byte) As Integer
If (HiByte And &H80) <> 0 Then
    MakeWord = ((HiByte * &H100&) Or LoByte) Or &HFFFF0000
Else
    MakeWord = (HiByte * &H100) Or LoByte
End If
End Function

Public Function LoWord(ByVal DWord As Long) As Integer
If DWord And &H8000& Then
    LoWord = DWord Or &HFFFF0000
Else
    LoWord = DWord And &HFFFF&
End If
End Function

Public Function HiWord(ByVal DWord As Long) As Integer
HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
MakeDWord = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function

Public Function Get_X_lParam(ByVal lParam As Long) As Long
Get_X_lParam = lParam And &H7FFF&
If lParam And &H8000& Then Get_X_lParam = Get_X_lParam Or &HFFFF8000
End Function

Public Function Get_Y_lParam(ByVal lParam As Long) As Long
Get_Y_lParam = (lParam And &H7FFF0000) \ &H10000
If lParam And &H80000000 Then Get_Y_lParam = Get_Y_lParam Or &HFFFF8000
End Function

Public Function StrToVar(ByVal Text As String) As Variant
Dim B() As Byte
B() = Text
StrToVar = B()
End Function

Public Function VarToStr(ByVal Bytes As Variant) As String
Dim B() As Byte
B() = Bytes
VarToStr = B()
End Function

Public Function UnsignedAdd(ByVal Start As Long, ByVal Incr As Long) As Long
UnsignedAdd = ((Start Xor &H80000000) + Incr) Xor &H80000000
End Function

Public Function CUIntToInt(ByVal Value As Long) As Integer
Const OFFSET_2 As Long = 65536
Const MAXINT_2 As Integer = 32767
If Value < 0 Or Value >= OFFSET_2 Then Err.Raise 6
If Value <= MAXINT_2 Then
    CUIntToInt = Value
Else
    CUIntToInt = Value - OFFSET_2
End If
End Function

Public Function CIntToUInt(ByVal Value As Integer) As Long
Const OFFSET_2 As Long = 65536
If Value < 0 Then
    CIntToUInt = Value + OFFSET_2
Else
    CIntToUInt = Value
End If
End Function

Public Function CULongToLong(ByVal Value As Double) As Long
Const OFFSET_4 As Double = 4294967296#
Const MAXINT_4 As Long = 2147483647
If Value < 0 Or Value >= OFFSET_4 Then Err.Raise 6
If Value <= MAXINT_4 Then
    CULongToLong = Value
Else
    CULongToLong = Value - OFFSET_4
End If
End Function

Public Function CLongToULong(ByVal Value As Long) As Double
Const OFFSET_4 As Double = 4294967296#
If Value < 0 Then
    CLongToULong = Value + OFFSET_4
Else
    CLongToULong = Value
End If
End Function

Public Function StrDecodeUTF8(ByVal Text As String) As String
Const CP_UTF8 As Long = 65001
Dim UTF8Size As Long
Dim Buffer As String, BufferSize As Long
Dim Result As Long
Dim UTF8() As Byte
If Not Text = vbNullString Then
    On Error GoTo Cancel
    UTF8() = StrConv(Text, vbFromUnicode)
    UTF8Size = UBound(UTF8()) + 1
    On Error GoTo 0
    BufferSize = UTF8Size * 2
    Buffer = String(BufferSize, vbNullChar)
    Result = MultiByteToWideChar(CP_UTF8, 0, UTF8(0), UTF8Size, StrPtr(Buffer), BufferSize)
    If Result <> 0 Then StrDecodeUTF8 = Left(Buffer, Result)
End If
Cancel:
End Function

Public Function StrEncodeUTF8(ByVal Text As String) As String
Const CP_UTF8 As Long = 65001
Dim Length As Long
Dim UTF16 As Long
Length = Len(Text)
If Length = 0 Then Exit Function
Dim BufferSize As Long
Dim Result As Long
Dim UTF8() As Byte
BufferSize = Length * 3 + 1
ReDim UTF8(BufferSize - 1)
Result = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), Length, UTF8(0), BufferSize, vbNullString, 0)
If Result <> 0 Then
    Result = Result - 1
    ReDim Preserve UTF8(Result)
    StrEncodeUTF8 = StrConv(UTF8(), vbUnicode)
End If
End Function

Public Function DPI_X() As Long
Const LOGPIXELSX As Long = 88
Dim hDCScreen As Long
hDCScreen = GetDC(0)
If hDCScreen <> 0 Then
    DPI_X = GetDeviceCaps(hDCScreen, LOGPIXELSX)
    ReleaseDC 0, hDCScreen
End If
End Function

Public Function DPI_Y() As Long
Const LOGPIXELSY As Long = 90
Dim hDCScreen As Long
hDCScreen = GetDC(0)
If hDCScreen <> 0 Then
    DPI_Y = GetDeviceCaps(hDCScreen, LOGPIXELSY)
    ReleaseDC 0, hDCScreen
End If
End Function

Public Function WinColor(ByVal Color As Long, Optional ByVal hPal As Long) As Long
Const S_OK As Long = &H0
If OleTranslateColor(Color, hPal, WinColor) <> S_OK Then WinColor = -1
End Function

Public Function R(ByVal Color As Long) As Byte
CopyMemory R, ByVal VarPtr(WinColor(Color)), 1
End Function

Public Function G(ByVal Color As Long) As Byte
CopyMemory G, ByVal VarPtr(WinColor(Color)) + 1, 1
End Function

Public Function B(ByVal Color As Long) As Byte
CopyMemory B, ByVal VarPtr(WinColor(Color)) + 2, 1
End Function

Public Function GrayColor(ByVal Color As Long) As Long
GrayColor = ((77& * (Color And &HFF&) + 152& * (Color And &HFF00&) \ &H100& + 28& * (Color \ &H10000)) \ 256&) * &H10101
End Function

Public Function LoadResImage(ByVal ResID As Long, ByVal ResType As String) As IPictureDisp
Set LoadResImage = PictureFromByteStream(LoadResData(ResID, ResType))
End Function

Public Function PictureFromByteStream(ByRef ByteStream As Variant) As IPictureDisp
Dim IID As CLSID
Dim B() As Byte, ByteCount As Long
Dim hMem  As Long, lpMem  As Long
Dim Stream As IUnknown
Dim NewPicture As IPicture
With IID
.Data1 = &H7BF80980
.Data2 = &HBF32
.Data3 = &H101A
.Data4(0) = &H8B
.Data4(1) = &HBB
.Data4(2) = &H0
.Data4(3) = &HAA
.Data4(4) = &H0
.Data4(5) = &H30
.Data4(6) = &HC
.Data4(7) = &HAB
End With
If VarType(ByteStream) = (vbArray + vbByte) Then
    B() = ByteStream
    ByteCount = (UBound(B()) - LBound(B())) + 1
    hMem = GlobalAlloc(&H2, ByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            CopyMemory ByVal lpMem, B(LBound(B())), ByteCount
            GlobalUnlock hMem
            If CreateStreamOnHGlobal(hMem, 1, Stream) = 0 Then
                If OleLoadPicture(ByVal Stream, ByteCount, 0, IID, NewPicture) = 0 Then Set PictureFromByteStream = NewPicture
            End If
        End If
    End If
End If
End Function

Public Function ImageHandleToPicture(ByVal hImage As Long, ByVal PicType As PictureTypeConstants) As IPictureDisp
If hImage = 0 Then Exit Function
Dim PICTD As PICTDESC, IID As CLSID
With PICTD
.cbSizeOfStruct = LenB(PICTD)
.PicType = PicType
.hImage = hImage
End With
With IID
.Data1 = &H7BF80980
.Data2 = &HBF32
.Data3 = &H101A
.Data4(0) = &H8B
.Data4(1) = &HBB
.Data4(2) = &H0
.Data4(3) = &HAA
.Data4(4) = &H0
.Data4(5) = &H30
.Data4(6) = &HC
.Data4(7) = &HAB
End With
OleCreatePictureIndirect PICTD, IID, True, ImageHandleToPicture
End Function

Public Function BitmapHandleFromPicture(ByVal Picture As IPictureDisp, Optional ByVal BackColor As OLE_COLOR) As Long
If Picture Is Nothing Then Exit Function
Dim hDCDesktop As Long, hBmp As Long
Dim hDC1 As Long, hBmpOld1 As Long
Dim hDC2 As Long, hBmpOld2 As Long
Dim Bmp As BITMAP, hImage As Long
If Not Picture.Type = vbPicTypeIcon Then
    hImage = Picture.Handle
Else
    Dim ICOI As ICONINFO
    GetIconInfo Picture.Handle, ICOI
    hImage = ICOI.hBMColor
End If
If hImage <> 0 Then
    GetObjectAPI hImage, LenB(Bmp), Bmp
    hDCDesktop = CreateDCAsNull(StrPtr("DISPLAY"), ByVal 0&, ByVal 0&, ByVal 0&)
    If hDCDesktop <> 0 Then
        If Not Picture.Type = vbPicTypeIcon Then
            hDC1 = CreateCompatibleDC(hDCDesktop)
            If hDC1 <> 0 Then
                hBmpOld1 = SelectObject(hDC1, hImage)
                hDC2 = CreateCompatibleDC(hDCDesktop)
                If hDC2 <> 0 Then
                    hBmp = CreateCompatibleBitmap(hDCDesktop, Bmp.BMWidth, Bmp.BMHeight)
                    If hBmp <> 0 Then
                        hBmpOld2 = SelectObject(hDC2, hBmp)
                        BitBlt hDC2, 0, 0, Bmp.BMWidth, Bmp.BMHeight, hDC1, 0, 0, vbSrcCopy
                        SelectObject hDC2, hBmpOld2
                        BitmapHandleFromPicture = hBmp
                    End If
                    DeleteDC hDC2
                End If
                SelectObject hDC1, hBmpOld1
                DeleteDC hDC1
            End If
        Else
            hDC1 = CreateCompatibleDC(hDCDesktop)
            If hDC1 <> 0 Then
                hBmp = CreateCompatibleBitmap(hDCDesktop, Bmp.BMWidth, Bmp.BMHeight)
                If hBmp <> 0 Then
                    hBmpOld1 = SelectObject(hDC1, hBmp)
                    Dim Brush As Long
                    Brush = CreateSolidBrush(WinColor(BackColor))
                    Const DI_NORMAL As Long = &H3
                    DrawIconEx hDC1, 0, 0, Picture.Handle, Bmp.BMWidth, Bmp.BMHeight, 0, Brush, DI_NORMAL
                    DeleteObject Brush
                    BitmapHandleFromPicture = hBmp
                End If
                SelectObject hDC1, hBmpOld1
                DeleteDC hDC1
            End If
        End If
        DeleteDC hDCDesktop
    End If
End If
End Function
