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
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal VKey As Long) As Integer
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
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal Color As Long, ByVal hPal As Long, ByRef RGBResult As Long) As Long
Private Declare Function OleLoadPicturePath Lib "oleaut32" (ByVal lpszPath As Long, ByVal pUnkCaller As Long, ByVal dwReserved As Long, ByVal ClrReserved As OLE_COLOR, ByRef riid As CLSID, ByRef pIPicture As IPicture) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (ByRef pPictDesc As PICTDESC, ByRef riid As Any, ByVal fPictureOwnsHandle As Long, ByRef pIPicture As IPicture) As Long

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

Public Function FileExists(ByVal PathName As String) As Boolean
On Error Resume Next
Dim Attributes As VbFileAttribute, ErrVal As Long
Attributes = GetAttr(PathName)
ErrVal = Err.Number
On Error GoTo 0
If (Attributes And (vbDirectory Or vbVolume)) = 0 And ErrVal = 0 Then FileExists = True
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

Public Function CreateFontFromOLEFont(ByVal Font As StdFont) As Long
Dim LF As LOGFONT, FontName As String
With LF
FontName = Left$(Font.Name, LF_FACESIZE)
CopyMemory .LFFaceName(0), ByVal StrPtr(FontName), LenB(FontName)
.LFHeight = -MulDiv(CLng(Font.Size), DPI_Y(), 72)
If Font.Bold = True Then .LFWeight = FW_BOLD Else .LFWeight = FW_NORMAL
.LFItalic = IIf(Font.Italic = True, 1, 0)
.LFStrikeOut = IIf(Font.Strikethrough = True, 1, 0)
.LFUnderline = IIf(Font.Underline = True, 1, 0)
.LFQuality = DEFAULT_QUALITY
.LFCharset = CByte(Font.Charset And &HFF)
End With
CreateFontFromOLEFont = CreateFontIndirect(LF)
End Function

Public Function CloneFont(ByVal Font As IFont) As StdFont
Font.Clone CloneFont
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
If OleTranslateColor(Color, hPal, WinColor) <> 0 Then WinColor = -1
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
