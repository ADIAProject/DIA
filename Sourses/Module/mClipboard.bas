Attribute VB_Name = "mClipboard"
Option Explicit

Private Enum eCBERRORMSG
    CB_OPEN_ERROR = 0
    CB_NO_BITMAP_FORMAT_AVAILABLE = 1
    CB_NO_TEXT_FORMAT_AVAILABLE = 2
    CB_ALREADY_OPEN = 3
End Enum

Public strCBError(3) As String
Public Declare Function GetOpenClipboardWindow Lib "user32.dll" () As Long
Public Declare Function CloseClipboard Lib "user32.dll" () As Long
Public Declare Function SetClipboardViewer Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function EnumClipboardFormats Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardFormatName Lib "user32.dll" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function CountClipboardFormats Lib "user32.dll" () As Long
Private Declare Function GetClipboardOwner Lib "user32.dll" () As Long
Private Declare Function ChangeClipboardChain Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndNext As Long) As Long
Private Declare Function GetClipboardViewer Lib "user32.dll" () As Long
Private Declare Function GetPriorityClipboardFormat Lib "user32.dll" (lpPriorityList As Long, ByVal nCount As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As OLEPIC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As Any) As Long
Private Declare Function CopyImage Lib "user32.dll" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_MOVEABLE = &H2
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2

Private Type OLEPIC
    Size                                As Long
    tType                               As Long
    hBmp                                As Long
    hPal                                As Long
    Reserved                            As Long
End Type

Private Type GUID
    Data1                               As Long
    Data2                               As Integer
    Data3                               As Integer
    Data4(7)                            As Byte
End Type

Private Const IMAGE_BITMAP = 0
Private Const LR_COPYRETURNORG = &H4
Private Const LR_CREATEDIBSECTION = &H2000

Public Const NO_CB_OPEN_ERROR = 0
Public Const NO_CB_OPENED = 0

Private Const NO_CB_FORMAT_AVAILABLE = 0
Private Const NO_CB_VIWER = 0

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CBSetText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sCBText (String)
'!--------------------------------------------------------------------------------
Public Sub CBSetText(ByVal sCBText As String)

    Dim hMem As Long, hPtr As Long, lLenBuffer As Long
    Dim s    As String

    If GetOpenClipboardWindow() = NO_CB_OPENED Then
        If OpenClipboard(frmMain.hWnd) <> NO_CB_OPEN_ERROR Then
            EmptyClipboard
            lLenBuffer = Len(sCBText) + 1
            s = String$(lLenBuffer, 0)
            Mid$(s, 1, lLenBuffer - 1) = sCBText
            hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_DDESHARE, lLenBuffer)
            hPtr = GlobalLock(hMem)
            CopyMemory ByVal hPtr, ByVal s, lLenBuffer
            GlobalUnlock hMem
            SetClipboardData CF_TEXT, hMem
            CloseClipboard
        Else
            MsgError CB_OPEN_ERROR
        End If

    Else
        MsgError CB_ALREADY_OPEN
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub MsgError
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   eErr (eCBERRORMSG)
'!--------------------------------------------------------------------------------
Private Sub MsgError(eErr As eCBERRORMSG)
    MsgBox strCBError(eErr), vbInformation, strAppEXEName
End Sub
