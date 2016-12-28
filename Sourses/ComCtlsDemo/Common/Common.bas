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
Private Const LF_FACESIZE As Long = 32
Private Const FW_NORMAL As Long = 400
Private Const FW_BOLD As Long = 700
Private Const DEFAULT_QUALITY As Long = 0
Private Type LOGFONT
LFHeight As Long
LFWidth As Long
LFEscapement As Long
LFOrientation As Long
LFWeight As Long
LFItalic As Byte
LFUnderline As Byte
LFStrikeOut As Byte
LFCharset As Byte
LFOutPrecision As Byte
LFClipPrecision As Byte
LFQuality As Byte
LFPitchAndFamily As Byte
LFFaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectW" (ByRef lpMsgBoxParams As MSGBOXPARAMS) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, ByVal lpCreationTime As Long, ByVal lpLastAccessTime As Long, ByVal lpLastWriteTime As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (ByVal lpFileTime As Long, ByVal lpLocalFileTime As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (ByVal lpFileTime As Long, ByVal lpSystemTime As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal NumberOfBytesToRead As Long, ByRef NumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
Private Declare Function PathGetArgs Lib "shlwapi" Alias "PathGetArgsW" (ByVal lpszPath As Long) As Long
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pbString As Long, ByVal pszStrPtr As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal VKey As Long) As Integer
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthW" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameW" (ByVal hWnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryW" (ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal XLeft As Long, ByVal YTop As Long, ByVal hIcon As Long, ByVal CXWidth As Long, ByVal CYWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, ByRef pIconInfo As ICONINFO) As Long
Private Declare Function CreateIconIndirect Lib "user32" (ByRef pIconInfo As ICONINFO) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal Color As Long, ByVal hPal As Long, ByRef RGBResult As Long) As Long
Private Declare Function OleLoadPicture Lib "oleaut32" (ByVal pStream As IUnknown, ByVal lSize As Long, ByVal fRunmode As Long, ByRef riid As Any, ByRef pIPicture As IPicture) As Long
Private Declare Function OleLoadPicturePath Lib "oleaut32" (ByVal lpszPath As Long, ByVal pUnkCaller As Long, ByVal dwReserved As Long, ByVal ClrReserved As OLE_COLOR, ByRef riid As CLSID, ByRef pIPicture As IPicture) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (ByRef pPictDesc As PICTDESC, ByRef riid As Any, ByVal fPictureOwnsHandle As Long, ByRef pIPicture As IPicture) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef pStream As IUnknown) As Long

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
Public Sub SendKeys(ByRef Text As String, Optional ByRef Wait As Boolean)
CreateObject("WScript.Shell").SendKeys Text, Wait
End Sub

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
Public Function FileLen(ByVal PathName As String) As Variant
Const INVALID_HANDLE_VALUE As Long = (-1), INVALID_FILE_SIZE As Long = (-1)
Const GENERIC_READ As Long = &H80000000, FILE_SHARE_READ As Long = &H1, OPEN_EXISTING As Long = 3, FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Dim hFile As Long
If Left$(PathName, 2) = "\\" Then PathName = "UNC\" & Mid$(PathName, 3)
hFile = CreateFile(StrPtr("\\?\" & PathName), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
If hFile <> INVALID_HANDLE_VALUE Then
    Dim LoDWord As Long, HiDWord As Long, Value As Variant
    LoDWord = GetFileSize(hFile, HiDWord)
    CloseHandle hFile
    If LoDWord <> INVALID_FILE_SIZE Then
        If (LoDWord And &H80000000) Then
            Value = CDec(LoDWord And &H7FFFFFFF) + CDec(2147483648#)
        Else
            Value = CDec(LoDWord)
        End If
        If (HiDWord And &H80000000) Then
            HiDWord = HiDWord And &H7FFFFFFF
            Value = Value + (CDec(HiDWord) + CDec(2147483648#)) * CDec(4294967296#)
        Else
            Value = Value + CDec(HiDWord) * CDec(4294967296#)
        End If
        FileLen = Value
    Else
        FileLen = Null
    End If
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

' (VB-Overwrite)
Public Function Command$()
If InIDE() = False Then
    SysReAllocString VarPtr(Command$), PathGetArgs(GetCommandLine())
    Command$ = LTrim$(Command$)
Else
    Command$ = VBA.Command$()
End If
End Function

Public Function GetEXEName() As String
If InIDE() = False Then
    Const MAX_PATH As Long = 260
    Dim Buffer As String
    Buffer = String(MAX_PATH, vbNullChar)
    Buffer = Left$(Buffer, GetModuleFileName(0, StrPtr(Buffer), MAX_PATH + 1))
    Buffer = Right$(Buffer, Len(Buffer) - InStrRev(Buffer, "\"))
    GetEXEName = Left$(Buffer, InStrRev(Buffer, ".") - 1)
Else
    GetEXEName = App.EXEName
End If
End Function

Public Function GetAppPath() As String
If InIDE() = False Then
    Const MAX_PATH As Long = 260
    Dim Buffer As String
    Buffer = String(MAX_PATH, vbNullChar)
    Buffer = Left$(Buffer, GetModuleFileName(0, StrPtr(Buffer), MAX_PATH + 1))
    GetAppPath = Left$(Buffer, InStrRev(Buffer, "\"))
Else
    GetAppPath = App.Path & IIf(Right$(App.Path, 1) = "\", "", "\")
End If
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

Public Function CreateGDIFontFromOLEFont(ByVal Font As StdFont) As Long
Dim LF As LOGFONT, FontName As String
With LF
FontName = Left$(Font.Name, LF_FACESIZE)
CopyMemory .LFFaceName(0), ByVal StrPtr(FontName), LenB(FontName)
.LFHeight = -MulDiv(CLng(Font.Size), DPI_Y(), 72)
If Font.Bold = True Then .LFWeight = FW_BOLD Else .LFWeight = FW_NORMAL
If Font.Italic = True Then .LFItalic = 1 Else .LFItalic = 0
If Font.Strikethrough = True Then .LFStrikeOut = 1 Else .LFStrikeOut = 0
If Font.Underline = True Then .LFUnderline = 1 Else .LFUnderline = 0
.LFQuality = DEFAULT_QUALITY
.LFCharset = CByte(Font.Charset And &HFF)
End With
CreateGDIFontFromOLEFont = CreateFontIndirect(LF)
End Function

Public Function CloneOLEFont(ByVal Font As IFont) As StdFont
Font.Clone CloneOLEFont
End Function

Public Function GDIFontFromOLEFont(ByVal Font As IFont) As Long
GDIFontFromOLEFont = Font.hFont
End Function

Public Function GetNumberGroupDigit() As String
GetNumberGroupDigit = Mid$(FormatNumber(1000, 0, , , vbTrue), 2, 1)
If GetNumberGroupDigit = "0" Then GetNumberGroupDigit = vbNullString
End Function

Public Function GetDecimalChar() As String
GetDecimalChar = Mid$(CStr(1.1), 2, 1)
End Function

Public Function IsFormLoaded(ByVal FormName As String) As Boolean
Dim i As Integer
For i = 0 To Forms.Count - 1
    If StrComp(Forms(i).Name, FormName, vbTextCompare) = 0 Then
        IsFormLoaded = True
        Exit For
    End If
Next i
End Function

Public Function GetWindowTitle(ByVal hWnd As Long) As String
Dim Buffer As String
Buffer = String(GetWindowTextLength(hWnd) + 1, vbNullChar)
GetWindowText hWnd, StrPtr(Buffer), Len(Buffer)
GetWindowTitle = Left$(Buffer, Len(Buffer) - 1)
End Function

Public Function GetWindowClassName(ByVal hWnd As Long) As String
Dim Buffer As String, RetVal As Long
Buffer = String(256, vbNullChar)
RetVal = GetClassName(hWnd, StrPtr(Buffer), Len(Buffer))
If RetVal <> 0 Then GetWindowClassName = Left$(Buffer, RetVal)
End Function

Public Function GetTitleBarHeight(ByVal Form As VB.Form) As Single
Const SM_CYCAPTION As Long = 4, SM_CYMENU As Long = 15
Const SM_CYSIZEFRAME As Long = 33, SM_CYFIXEDFRAME As Long = 8
Dim CY As Long
CY = GetSystemMetrics(SM_CYCAPTION)
If GetMenu(Form.hWnd) <> 0 Then CY = CY + GetSystemMetrics(SM_CYMENU)
Select Case Form.BorderStyle
    Case vbSizable, vbSizableToolWindow
        CY = CY + (GetSystemMetrics(SM_CYSIZEFRAME) * 2)
    Case vbFixedSingle, vbFixedDialog, vbFixedToolWindow
        CY = CY + (GetSystemMetrics(SM_CYFIXEDFRAME) * 2)
End Select
If CY > 0 Then GetTitleBarHeight = Form.ScaleY(CY, vbPixels, Form.ScaleMode)
End Function

Public Sub SetWindowRedraw(ByVal hWnd As Long, ByVal Enabled As Boolean)
Const WM_SETREDRAW As Long = &HB
SendMessage hWnd, WM_SETREDRAW, IIf(Enabled = True, 1, 0), ByVal 0&
If Enabled = True Then
    Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
    RedrawWindow hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End If
End Sub

Public Function GetWinPath() As String
Const MAX_PATH As Long = 260
Dim Buffer As String
Buffer = String(MAX_PATH, vbNullChar)
If GetWindowsDirectory(StrPtr(Buffer), MAX_PATH) <> 0 Then
    GetWinPath = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    GetWinPath = GetWinPath & IIf(Right$(GetWinPath, 1) = "\", "", "\")
End If
End Function

Public Function GetSysPath() As String
Const MAX_PATH As Long = 260
Dim Buffer As String
Buffer = String(MAX_PATH, vbNullChar)
If GetSystemDirectory(StrPtr(Buffer), MAX_PATH) <> 0 Then
    GetSysPath = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    GetSysPath = GetSysPath & IIf(Right$(GetSysPath, 1) = "\", "", "\")
End If
End Function

Public Function GetShiftStateFromParam(ByVal wParam As Long) As ShiftConstants
Const MK_SHIFT As Long = &H4, MK_CONTROL As Long = &H8
If (wParam And MK_SHIFT) = MK_SHIFT Then GetShiftStateFromParam = vbShiftMask
If (wParam And MK_CONTROL) = MK_CONTROL Then GetShiftStateFromParam = GetShiftStateFromParam Or vbCtrlMask
If GetKeyState(vbKeyMenu) < 0 Then GetShiftStateFromParam = GetShiftStateFromParam Or vbAltMask
End Function

Public Function GetMouseStateFromParam(ByVal wParam As Long) As MouseButtonConstants
Const MK_LBUTTON As Long = &H1, MK_RBUTTON As Long = &H2, MK_MBUTTON As Long = &H10
If (wParam And MK_LBUTTON) = MK_LBUTTON Then GetMouseStateFromParam = vbLeftButton
If (wParam And MK_RBUTTON) = MK_RBUTTON Then GetMouseStateFromParam = GetMouseStateFromParam Or vbRightButton
If (wParam And MK_MBUTTON) = MK_MBUTTON Then GetMouseStateFromParam = GetMouseStateFromParam Or vbMiddleButton
End Function

Public Function GetShiftStateFromMsg() As ShiftConstants
If GetKeyState(vbKeyShift) < 0 Then GetShiftStateFromMsg = vbShiftMask
If GetKeyState(vbKeyControl) < 0 Then GetShiftStateFromMsg = GetShiftStateFromMsg Or vbCtrlMask
If GetKeyState(vbKeyMenu) < 0 Then GetShiftStateFromMsg = GetShiftStateFromMsg Or vbAltMask
End Function

Public Function GetMouseStateFromMsg() As MouseButtonConstants
If GetKeyState(vbLeftButton) < 0 Then GetMouseStateFromMsg = vbLeftButton
If GetKeyState(vbRightButton) < 0 Then GetMouseStateFromMsg = GetMouseStateFromMsg Or vbRightButton
If GetKeyState(vbMiddleButton) < 0 Then GetMouseStateFromMsg = GetMouseStateFromMsg Or vbMiddleButton
End Function

Public Function GetShiftState() As ShiftConstants
GetShiftState = (-vbShiftMask * KeyPressed(vbKeyShift))
GetShiftState = GetShiftState Or (-vbCtrlMask * KeyPressed(vbKeyControl))
GetShiftState = GetShiftState Or (-vbAltMask * KeyPressed(vbKeyMenu))
End Function

Public Function GetMouseState() As MouseButtonConstants
Const SM_SWAPBUTTON As Long = 23
' GetAsyncKeyState requires a mapping of physical mouse buttons to logical mouse buttons.
GetMouseState = (-vbLeftButton * KeyPressed(IIf(GetSystemMetrics(SM_SWAPBUTTON) = 0, vbLeftButton, vbRightButton)))
GetMouseState = GetMouseState Or (-vbRightButton * KeyPressed(IIf(GetSystemMetrics(SM_SWAPBUTTON) = 0, vbRightButton, vbLeftButton)))
GetMouseState = GetMouseState Or (-vbMiddleButton * KeyPressed(vbMiddleButton))
End Function

Public Function KeyToggled(ByVal KeyCode As KeyCodeConstants) As Boolean
KeyToggled = CBool(LoByte(GetKeyState(KeyCode)) = 1)
End Function
 
Public Function KeyPressed(ByVal KeyCode As KeyCodeConstants) As Boolean
KeyPressed = CBool((GetAsyncKeyState(KeyCode) And &H8000&) = &H8000&)
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
If Text = vbNullString Then
    StrToVar = Empty
Else
    Dim B() As Byte
    B() = Text
    StrToVar = B()
End If
End Function

Public Function VarToStr(ByVal Bytes As Variant) As String
If IsEmpty(Bytes) Then
    VarToStr = vbNullString
Else
    Dim B() As Byte
    B() = Bytes
    VarToStr = B()
End If
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

Public Function CULngToLng(ByVal Value As Double) As Long
Const OFFSET_4 As Double = 4294967296#
Const MAXINT_4 As Long = 2147483647
If Value < 0 Or Value >= OFFSET_4 Then Err.Raise 6
If Value <= MAXINT_4 Then
    CULngToLng = Value
Else
    CULngToLng = Value - OFFSET_4
End If
End Function

Public Function CLngToULng(ByVal Value As Long) As Double
Const OFFSET_4 As Double = 4294967296#
If Value < 0 Then
    CLngToULng = Value + OFFSET_4
Else
    CLngToULng = Value
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

Public Function DPICorrectionFactor() As Single
Static Done As Boolean, Value As Single
If Done = False Then
    Value = Screen.TwipsPerPixelX / ((96 / DPI_X()) * 15)
    Done = True
End If
' Returns exactly 1 when no corrections are required.
DPICorrectionFactor = Value
End Function

Public Function WinColor(ByVal Color As Long, Optional ByVal hPal As Long) As Long
If OleTranslateColor(Color, hPal, WinColor) <> 0 Then WinColor = -1
End Function

Public Function PictureFromByteStream(ByRef ByteStream As Variant) As IPictureDisp
Dim IID As CLSID, Stream As IUnknown, NewPicture As IPicture
Dim B() As Byte, ByteCount As Long
Dim hMem As Long, lpMem As Long
With IID
.Data1 = &H7BF80980
.Data2 = &HBF32
.Data3 = &H101A
.Data4(0) = &H8B
.Data4(1) = &HBB
.Data4(3) = &HAA
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
                If OleLoadPicture(Stream, ByteCount, 0, IID, NewPicture) = 0 Then Set PictureFromByteStream = NewPicture
            End If
        End If
    End If
End If
End Function

Public Function PictureFromPath(ByVal PathName As String) As IPictureDisp
Dim IID As CLSID, NewPicture As IPicture
With IID
.Data1 = &H7BF80980
.Data2 = &HBF32
.Data3 = &H101A
.Data4(0) = &H8B
.Data4(1) = &HBB
.Data4(3) = &HAA
.Data4(5) = &H30
.Data4(6) = &HC
.Data4(7) = &HAB
End With
If OleLoadPicturePath(StrPtr(PathName), 0, 0, 0, IID, NewPicture) = 0 Then Set PictureFromPath = NewPicture
End Function

Public Function PictureFromHandle(ByVal Handle As Long, ByVal PicType As VBRUN.PictureTypeConstants) As IPictureDisp
If Handle = 0 Then Exit Function
Dim PICD As PICTDESC, IID As CLSID, NewPicture As IPicture
With PICD
.cbSizeOfStruct = LenB(PICD)
.PicType = PicType
.hImage = Handle
End With
With IID
.Data1 = &H7BF80980
.Data2 = &HBF32
.Data3 = &H101A
.Data4(0) = &H8B
.Data4(1) = &HBB
.Data4(3) = &HAA
.Data4(5) = &H30
.Data4(6) = &HC
.Data4(7) = &HAB
End With
If OleCreatePictureIndirect(PICD, IID, 1, NewPicture) = 0 Then Set PictureFromHandle = NewPicture
End Function

Public Function BitmapHandleFromPicture(ByVal Picture As IPictureDisp, Optional ByVal BackColor As OLE_COLOR) As Long
If Picture Is Nothing Then Exit Function
Dim hDCScreen As Long, hBmp As Long
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
    hDCScreen = GetDC(0)
    If hDCScreen <> 0 Then
        If Not Picture.Type = vbPicTypeIcon Then
            hDC1 = CreateCompatibleDC(hDCScreen)
            If hDC1 <> 0 Then
                hBmpOld1 = SelectObject(hDC1, hImage)
                hDC2 = CreateCompatibleDC(hDCScreen)
                If hDC2 <> 0 Then
                    hBmp = CreateCompatibleBitmap(hDCScreen, Bmp.BMWidth, Bmp.BMHeight)
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
            hDC1 = CreateCompatibleDC(hDCScreen)
            If hDC1 <> 0 Then
                hBmp = CreateCompatibleBitmap(hDCScreen, Bmp.BMWidth, Bmp.BMHeight)
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
        ReleaseDC 0, hDCScreen
    End If
End If
End Function
