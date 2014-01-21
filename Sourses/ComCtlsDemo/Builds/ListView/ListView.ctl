VERSION 5.00
Begin VB.UserControl ListView 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   HasDC           =   0   'False
   PropertyPages   =   "ListView.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "ListView.ctx":0074
   Begin VB.Timer TimerImageList 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
#If False Then
Private LvwViewIcon, LvwViewSmallIcon, LvwViewList, LvwViewReport, LvwViewTile
Private LvwArrangeNone, LvwArrangeAutoLeft, LvwArrangeAutoTop, LvwArrangeLeft, LvwArrangeTop
Private LvwColumnHeaderAlignmentLeft, LvwColumnHeaderAlignmentRight, LvwColumnHeaderAlignmentCenter
Private LvwColumnHeaderSortArrowNone, LvwColumnHeaderSortArrowDown, LvwColumnHeaderSortArrowUp
Private LvwColumnHeaderAutoSizeToItems, LvwColumnHeaderAutoSizeToHeader
Private LvwLabelEditAutomatic, LvwLabelEditManual, LvwLabelEditDisabled
Private LvwSortOrderAscending, LvwSortOrderDescending
Private LvwSortTypeBinary, LvwSortTypeText, LvwSortTypeNumeric, LvwSortTypeCurrency, LvwSortTypeDate
Private LvwPictureAlignmentTopLeft, LvwPictureAlignmentTopRight, LvwPictureAlignmentBottomLeft, LvwPictureAlignmentBottomRight, LvwPictureAlignmentCenter, LvwPictureAlignmentTile
#End If
Public Enum LvwViewConstants
LvwViewIcon = 0
LvwViewSmallIcon = 1
LvwViewList = 2
LvwViewReport = 3
LvwViewTile = 4
End Enum
Public Enum LvwArrangeConstants
LvwArrangeNone = 0
LvwArrangeAutoLeft = 1
LvwArrangeAutoTop = 2
LvwArrangeLeft = 3
LvwArrangeTop = 4
End Enum
Public Enum LvwColumnHeaderAlignmentConstants
LvwColumnHeaderAlignmentLeft = 0
LvwColumnHeaderAlignmentRight = 1
LvwColumnHeaderAlignmentCenter = 2
End Enum
Public Enum LvwColumnHeaderSortArrowConstants
LvwColumnHeaderSortArrowNone = 0
LvwColumnHeaderSortArrowDown = 1
LvwColumnHeaderSortArrowUp = 2
End Enum
Public Enum LvwColumnHeaderAutoSizeConstants
LvwColumnHeaderAutoSizeToItems = 0
LvwColumnHeaderAutoSizeToHeader = 1
End Enum
Public Enum LvwLabelEditConstants
LvwLabelEditAutomatic = 0
LvwLabelEditManual = 1
LvwLabelEditDisabled = 2
End Enum
Public Enum LvwSortOrderConstants
LvwSortOrderAscending = 0
LvwSortOrderDescending = 1
End Enum
Public Enum LvwSortTypeConstants
LvwSortTypeBinary = 0
LvwSortTypeText = 1
LvwSortTypeNumeric = 2
LvwSortTypeCurrency = 3
LvwSortTypeDate = 4
End Enum
Public Enum LvwPictureAlignmentConstants
LvwPictureAlignmentTopLeft = 0
LvwPictureAlignmentTopRight = 1
LvwPictureAlignmentBottomLeft = 2
LvwPictureAlignmentBottomRight = 3
LvwPictureAlignmentCenter = 4
LvwPictureAlignmentTile = 5
End Enum
Private Type TagInitCommonControlsEx
dwSize As Long
dwICC As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type SIZEAPI
CX As Long
CY As Long
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
Private Type LVITEM
Mask As Long
iItem As Long
iSubItem As Long
State As Long
StateMask As Long
pszText As Long
cchTextMax As Long
iImage As Long
lParam As Long
iIndent As Long
End Type
Private Type LVTILEINFO
cbSize As Long
iItem As Long
cColumns As Long
puColumns As Long
End Type
Private Type LVTILEVIEWINFO
cbSize As Long
dwMask As Long
dwFlags As Long
SizeTile As SIZEAPI
cLines As Long
rcLabelMargin As RECT
End Type
Private Type LVFINDINFO
Flags As Long
psz As Long
lParam As Long
PT As POINTAPI
VKDirection As Long
End Type
Private Type LVCOLUMN
Mask As Long
fmt As Long
CX As Long
pszText As Long
cchTextMax As Long
iSubItem As Long
iImage As Long
iOrder As Long
End Type
Private Type LVHITTESTINFO
PT As POINTAPI
Flags As Long
iItem As Long
iSubItem As Long
End Type
Private Type LVINSERTMARK
cbSize As Long
dwFlags As Long
iItem As Long
dwReserved As Long
End Type
Private Type LVBKIMAGE
ulFlags As Long
hBmp As Long
pszImage As String
cchImageMax As Long
XOffsetPercent As Long
YOffsetPercent As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Const CDDS_PREPAINT As Long = &H1
Private Const CDDS_POSTPAINT As Long = &H2
Private Const CDDS_PREERASE As Long = &H3
Private Const CDDS_POSTERASE As Long = &H4
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_ITEMPREPAINT As Long = (CDDS_ITEM + 1)
Private Const CDDS_ITEMPOSTPAINT As Long = (CDDS_ITEM + 2)
Private Const CDDS_SUBITEM As Long = &H20000
Private Const CDIS_CHECKED As Long = &H8
Private Const CDIS_FOCUS As Long = &H10
Private Const CDIS_HOT As Long = &H40
Private Const CDRF_DODEFAULT As Long = &H0
Private Const CDRF_NEWFONT As Long = &H2
Private Const CDRF_SKIPDEFAULT As Long = &H4
Private Const CDRF_NOTIFYPOSTPAINT As Long = &H10
Private Const CDRF_NOTIFYITEMDRAW As Long = &H20
Private Const CDRF_NOTIFYPOSTERASE As Long = &H40
Private Const CDRF_NOTIFYSUBITEMDRAW As Long = &H20
Private Type NMCUSTOMDRAW
hdr As NMHDR
dwDrawStage As Long
hDC As Long
RC As RECT
dwItemSpec As Long
uItemState As Long
lItemlParam As Long
End Type
Private Type NMLVCUSTOMDRAW
NMCD As NMCUSTOMDRAW
ClrText As Long
ClrTextBk As Long
iSubItem As Long
End Type
Private Type NMLISTVIEW
hdr As NMHDR
iItem As Long
iSubItem As Long
uNewState As Long
uOldState As Long
uChanged As Long
PTAction As POINTAPI
lParam As Long
End Type
Private Type NMITEMACTIVATE
hdr As NMHDR
iItem As Long
iSubItem As Long
uNewState As Long
uOldState As Long
uChanged As Long
PTAction As POINTAPI
lParam As Long
uKeyFlags As Long
End Type
Private Type NMLVGETINFOTIP
hdr As NMHDR
dwFlags As Long
pszText As Long
cchTextMax As Long
iItem As Long
iSubItem As Long
lParam As Long
End Type
Private Type NMLVDISPINFO
hdr As NMHDR
Item As LVITEM
End Type
Private Const L_MAX_URL_LENGTH As Long = 2084
Private Type NMLVEMPTYMARKUP
hdr As NMHDR
dwFlags As Long
szMarkup(0 To ((L_MAX_URL_LENGTH * 2) - 1)) As Byte
End Type
Private Type NMHEADER
hdr As NMHDR
iItem As Long
iButton As Long
lPtrHDItem As Long
End Type
Private Type HDITEM
Mask As Long
CXY As Long
pszText As Long
hBm As Long
cchTextMax As Long
fmt As Long
lParam As Long
iImage As Long
iOrder As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event ContextMenu(ByVal X As Single, ByVal Y As Single)
Attribute ContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
Public Event ItemClick(ByVal Item As LvwListItem, ByVal Button As Integer)
Attribute ItemClick.VB_Description = "Occurs when a list item is clicked."
Public Event ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)
Attribute ItemDblClick.VB_Description = "Occurs when a list item is double clicked."
Public Event ItemFocus(ByVal Item As LvwListItem)
Attribute ItemFocus.VB_Description = "Occurs when a list item is focused."
Public Event ItemActivate(ByVal Item As LvwListItem, ByVal SubItemIndex As Long, ByVal Shift As Integer)
Attribute ItemActivate.VB_Description = "Occurs when a list item is activated."
Public Event ItemSelect(ByVal Item As LvwListItem, ByVal Selected As Boolean)
Attribute ItemSelect.VB_Description = "Occurs when a list item is selected."
Public Event ItemCheck(ByVal Item As LvwListItem, ByVal Checked As Boolean)
Attribute ItemCheck.VB_Description = "Occurs when a list item is checked."
Public Event ItemDrag(ByVal Item As LvwListItem, ByVal Button As Integer)
Attribute ItemDrag.VB_Description = "Occurs when a list item initiate a drag-and-drop operation."
Public Event ItemBkColor(ByVal Item As LvwListItem, ByRef ColorRef As Long)
Attribute ItemBkColor.VB_Description = "Occurs when a list item is about to draw the background in 'report' view. This is a request to provide an alternative back color. The back color is passed in an RGB format."
Public Event BeforeLabelEdit(ByRef Cancel As Boolean)
Attribute BeforeLabelEdit.VB_Description = "Occurs when a user attempts to edit the label of the currently selected list item."
Public Event AfterLabelEdit(ByRef Cancel As Boolean, ByRef NewString As String)
Attribute AfterLabelEdit.VB_Description = "Occurs after a user edits the label of the currently selected list item."
Public Event ColumnClick(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnClick.VB_Description = "Occurs when a column header in a list view is clicked."
Public Event ColumnBeforeSize(ByVal ColumnHeader As LvwColumnHeader, ByRef Cancel As Boolean)
Attribute ColumnBeforeSize.VB_Description = "Occurs when the user has begun dragging a divider one one column header."
Public Event ColumnAfterSize(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnAfterSize.VB_Description = "Occurs when the user has finished dragging a divider on one column header."
Public Event ColumnBeforeDrag(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnBeforeDrag.VB_Description = "Occurs when a drag operation has begun on one column header."
Public Event ColumnAfterDrag(ByVal ColumnHeader As LvwColumnHeader, ByVal NewPosition As Long, ByRef Cancel As Boolean)
Attribute ColumnAfterDrag.VB_Description = "Occurs when a drag operation has ended on one column header."
Public Event ColumnDropDown(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnDropDown.VB_Description = "Occurs when the drop-down arrow on the split button of a column header is clicked. Requires comctl32.dll version 6.1 or higher."
Public Event ColumnCheck(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnCheck.VB_Description = "Occurs when a column header is checked. Requires comctl32.dll version 6.1 or higher."
Public Event GetEmptyMarkup(ByRef Text As String, ByRef Centered As Boolean)
Attribute GetEmptyMarkup.VB_Description = "Occurs when the list view has no list items. This is a request to provide a markup text. Requires comctl32.dll version 6.1 or higher."
Public Event BeginMarqueeSelection(ByRef Cancel As Boolean)
Attribute BeginMarqueeSelection.VB_Description = "Occurs when a bounding box (marquee) selection has begun. Only applicable if the multi select property is set to true."
Public Event PreviewKeyDown(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyDown.VB_Description = "Occurs before the KeyDown event."
Public Event PreviewKeyUp(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyUp.VB_Description = "Occurs before the KeyUp event."
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Public Event KeyPress(KeyChar As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an character key."
Attribute KeyPress.VB_UserMemId = -603
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Public Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcmp Lib "kernel32" Alias "lstrcmpW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function SetWindowTheme Lib "uxtheme" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TagInitCommonControlsEx) As Long
Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function SysAllocString Lib "oleaut32" (ByVal lpString As Long) As Long
Private Declare Function SysFreeString Lib "oleaut32" (ByVal lpString As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageSort Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As ISubclass, ByRef lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Const ICC_LISTVIEW_CLASSES As Long = &H1
Private Const RDW_UPDATENOW As Long = &H100
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4
Private Const RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const CF_UNICODETEXT As Long = 13
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_BORDER As Long = &H800000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_EX_CLIENTEDGE As Long = &H200
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_WINDOWEDGE As Long = &H100
Private Const WS_EX_RTLREADING As Long = &H2000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const WM_VSCROLL As Long = &H115
Private Const WM_HSCROLL As Long = &H114
Private Const SB_LINELEFT As Long = 0, SB_LINERIGHT As Long = 1
Private Const SB_LINEUP As Long = 0, SB_LINEDOWN As Long = 1
Private Const WM_MOUSEACTIVATE As Long = &H21, MA_NOACTIVATE As Long = &H3, MA_NOACTIVATEANDEAT As Long = &H4, HTBORDER As Long = 18
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const SW_HIDE As Long = &H0
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_NOTIFYFORMAT As Long = &H55
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_IME_CHAR As Long = &H286
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_SHOWWINDOW As Long = &H18
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const CLR_NONE As Long = &HFFFFFFFF
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETVERSION As Long = (CCM_FIRST + 7)
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETBKCOLOR As Long = (LVM_FIRST + 0)
Private Const LVM_SETBKCOLOR As Long = (LVM_FIRST + 1)
Private Const LVM_GETIMAGELIST As Long = (LVM_FIRST + 2)
Private Const LVM_SETIMAGELIST As Long = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT As Long = (LVM_FIRST + 4)
Private Const LVM_GETITEMA As Long = (LVM_FIRST + 5)
Private Const LVM_GETITEMW As Long = (LVM_FIRST + 75)
Private Const LVM_GETITEM As Long = LVM_GETITEMW
Private Const LVM_SETITEMA As Long = (LVM_FIRST + 6)
Private Const LVM_SETITEMW As Long = (LVM_FIRST + 76)
Private Const LVM_SETITEM As Long = LVM_SETITEMW
Private Const LVM_INSERTITEMA As Long = (LVM_FIRST + 7)
Private Const LVM_INSERTITEMW As Long = (LVM_FIRST + 77)
Private Const LVM_INSERTITEM As Long = LVM_INSERTITEMW
Private Const LVM_DELETEITEM As Long = (LVM_FIRST + 8)
Private Const LVM_DELETEALLITEMS As Long = (LVM_FIRST + 9)
Private Const LVM_GETCALLBACKMASK As Long = (LVM_FIRST + 10)
Private Const LVM_SETCALLBACKMASK As Long = (LVM_FIRST + 11)
Private Const LVM_GETNEXTITEM As Long = (LVM_FIRST + 12)
Private Const LVM_FINDITEMA As Long = (LVM_FIRST + 13)
Private Const LVM_FINDITEMW As Long = (LVM_FIRST + 83)
Private Const LVM_FINDITEM As Long = LVM_FINDITEMW
Private Const LVM_RESETEMPTYTEXT As Long = (LVM_FIRST + 84) ' Undocumented
Private Const LVM_GETITEMRECT As Long = (LVM_FIRST + 14)
Private Const LVM_SETITEMPOSITION As Long = (LVM_FIRST + 15)
Private Const LVM_GETITEMPOSITION As Long = (LVM_FIRST + 16)
Private Const LVM_GETSTRINGWIDTHA As Long = (LVM_FIRST + 17)
Private Const LVM_GETSTRINGWIDTHW As Long = (LVM_FIRST + 87)
Private Const LVM_GETSTRINGWIDTH As Long = LVM_GETSTRINGWIDTHW
Private Const LVM_HITTEST As Long = (LVM_FIRST + 18)
Private Const LVM_ENSUREVISIBLE As Long = (LVM_FIRST + 19)
Private Const LVM_SCROLL As Long = (LVM_FIRST + 20)
Private Const LVM_REDRAWITEMS As Long = (LVM_FIRST + 21)
Private Const LVM_ARRANGE As Long = (LVM_FIRST + 22)
Private Const LVM_EDITLABELA As Long = (LVM_FIRST + 23)
Private Const LVM_EDITLABELW As Long = (LVM_FIRST + 118)
Private Const LVM_EDITLABEL As Long = LVM_EDITLABELW
Private Const LVM_GETEDITCONTROL As Long = (LVM_FIRST + 24)
Private Const LVM_GETCOLUMNA As Long = (LVM_FIRST + 25)
Private Const LVM_GETCOLUMNW As Long = (LVM_FIRST + 95)
Private Const LVM_GETCOLUMN As Long = LVM_GETCOLUMNW
Private Const LVM_SETCOLUMNA As Long = (LVM_FIRST + 26)
Private Const LVM_SETCOLUMNW As Long = (LVM_FIRST + 96)
Private Const LVM_SETCOLUMN As Long = LVM_SETCOLUMNW
Private Const LVM_INSERTCOLUMNA As Long = (LVM_FIRST + 27)
Private Const LVM_INSERTCOLUMNW As Long = (LVM_FIRST + 97)
Private Const LVM_INSERTCOLUMN As Long = LVM_INSERTCOLUMNW
Private Const LVM_DELETECOLUMN As Long = (LVM_FIRST + 28)
Private Const LVM_GETCOLUMNWIDTH As Long = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVM_GETHEADER As Long = (LVM_FIRST + 31)
Private Const LVM_CREATEDRAGIMAGE As Long = (LVM_FIRST + 33)
Private Const LVM_GETVIEWRECT As Long = (LVM_FIRST + 34)
Private Const LVM_GETTEXTCOLOR As Long = (LVM_FIRST + 35)
Private Const LVM_SETTEXTCOLOR As Long = (LVM_FIRST + 36)
Private Const LVM_GETTEXTBKCOLOR As Long = (LVM_FIRST + 37)
Private Const LVM_SETTEXTBKCOLOR As Long = (LVM_FIRST + 38)
Private Const LVM_GETTOPINDEX As Long = (LVM_FIRST + 39)
Private Const LVM_GETCOUNTPERPAGE As Long = (LVM_FIRST + 40)
Private Const LVM_GETORIGIN As Long = (LVM_FIRST + 41)
Private Const LVM_UPDATE As Long = (LVM_FIRST + 42)
Private Const LVM_SETITEMSTATE As Long = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXTA As Long = (LVM_FIRST + 45)
Private Const LVM_GETITEMTEXTW As Long = (LVM_FIRST + 115)
Private Const LVM_GETITEMTEXT As Long = LVM_GETITEMTEXTW
Private Const LVM_SETITEMTEXTA As Long = (LVM_FIRST + 46)
Private Const LVM_SETITEMTEXTW As Long = (LVM_FIRST + 116)
Private Const LVM_SETITEMTEXT As Long = LVM_SETITEMTEXTW
Private Const LVM_SETITEMCOUNT As Long = (LVM_FIRST + 47)
Private Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)
Private Const LVM_SETITEMPOSITION32 As Long = (LVM_FIRST + 49)
Private Const LVM_GETSELECTEDCOUNT As Long = (LVM_FIRST + 50)
Private Const LVM_GETITEMSPACING As Long = (LVM_FIRST + 51)
Private Const LVM_GETISEARCHSTRINGA As Long = (LVM_FIRST + 52)
Private Const LVM_GETISEARCHSTRINGW As Long = (LVM_FIRST + 117)
Private Const LVM_GETISEARCHSTRING As Long = LVM_GETISEARCHSTRINGW
Private Const LVM_SETICONSPACING As Long = (LVM_FIRST + 53)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)
Private Const LVM_GETSUBITEMRECT As Long = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)
Private Const LVM_SETCOLUMNORDERARRAY As Long = (LVM_FIRST + 58)
Private Const LVM_GETCOLUMNORDERARRAY As Long = (LVM_FIRST + 59)
Private Const LVM_SETHOTITEM As Long = (LVM_FIRST + 60)
Private Const LVM_GETHOTITEM As Long = (LVM_FIRST + 61)
Private Const LVM_SETHOTCURSOR As Long = (LVM_FIRST + 62)
Private Const LVM_GETHOTCURSOR As Long = (LVM_FIRST + 63)
Private Const LVM_APPROXIMATEVIEWRECT As Long = (LVM_FIRST + 64)
Private Const LVM_SETWORKAREAS As Long = (LVM_FIRST + 65)
Private Const LVM_GETSELECTIONMARK As Long = (LVM_FIRST + 66)
Private Const LVM_SETSELECTIONMARK As Long = (LVM_FIRST + 67)
Private Const LVM_SETBKIMAGEA As Long = (LVM_FIRST + 68)
Private Const LVM_SETBKIMAGEW As Long = (LVM_FIRST + 138)
Private Const LVM_SETBKIMAGE As Long = LVM_SETBKIMAGEW
Private Const LVM_GETBKIMAGEA As Long = (LVM_FIRST + 69)
Private Const LVM_GETBKIMAGEW As Long = (LVM_FIRST + 139)
Private Const LVM_GETBKIMAGE As Long = LVM_GETBKIMAGEW
Private Const LVM_GETWORKAREAS As Long = (LVM_FIRST + 70)
Private Const LVM_SETHOVERTIME As Long = (LVM_FIRST + 71)
Private Const LVM_GETHOVERTIME As Long = (LVM_FIRST + 72)
Private Const LVM_GETNUMBEROFWORKAREAS As Long = (LVM_FIRST + 73)
Private Const LVM_SETTOOLTIPS As Long = (LVM_FIRST + 74)
Private Const LVM_GETTOOLTIPS As Long = (LVM_FIRST + 78)
Private Const LVM_SORTITEMSEX As Long = (LVM_FIRST + 81)
Private Const LVM_SETSELECTEDCOLUMN As Long = (LVM_FIRST + 140)
Private Const LVM_SETVIEW As Long = (LVM_FIRST + 142)
Private Const LVM_GETVIEW As Long = (LVM_FIRST + 143)
Private Const LVM_SETTILEVIEWINFO As Long = (LVM_FIRST + 162)
Private Const LVM_GETTILEVIEWINFO As Long = (LVM_FIRST + 163)
Private Const LVM_SETTILEINFO As Long = (LVM_FIRST + 164)
Private Const LVM_GETTILEINFO As Long = (LVM_FIRST + 165)
Private Const LVM_SETINSERTMARK As Long = (LVM_FIRST + 166)
Private Const LVM_GETINSERTMARK As Long = (LVM_FIRST + 167)
Private Const LVM_INSERTMARKHITTEST As Long = (LVM_FIRST + 168)
Private Const LVM_GETINSERTMARKRECT As Long = (LVM_FIRST + 169)
Private Const LVM_SETINSERTMARKCOLOR As Long = (LVM_FIRST + 170)
Private Const LVM_GETINSERTMARKCOLOR As Long = (LVM_FIRST + 171)
Private Const LVM_SETINFOTIP As Long = (LVM_FIRST + 173)
Private Const LVM_GETSELECTEDCOLUMN As Long = (LVM_FIRST + 174)
Private Const LVM_ISITEMVISIBLE As Long = (LVM_FIRST + 182)
Private Const LVN_FIRST As Long = (-100)
Private Const LVN_ITEMCHANGING As Long = (LVN_FIRST - 0)
Private Const LVN_ITEMCHANGED As Long = (LVN_FIRST - 1)
Private Const LVN_INSERTITEM As Long = (LVN_FIRST - 2)
Private Const LVN_DELETEITEM As Long = (LVN_FIRST - 3)
Private Const LVN_DELETEALLITEMS As Long = (LVN_FIRST - 4)
Private Const LVN_BEGINLABELEDITA As Long = (LVN_FIRST - 5)
Private Const LVN_BEGINLABELEDITW As Long = (LVN_FIRST - 75)
Private Const LVN_BEGINLABELEDIT As Long = LVN_BEGINLABELEDITW
Private Const LVN_ENDLABELEDITA As Long = (LVN_FIRST - 6)
Private Const LVN_ENDLABELEDITW As Long = (LVN_FIRST - 76)
Private Const LVN_ENDLABELEDIT As Long = LVN_ENDLABELEDITW
Private Const LVN_COLUMNCLICK As Long = (LVN_FIRST - 8)
Private Const LVN_BEGINDRAG As Long = (LVN_FIRST - 9)
Private Const LVN_BEGINRDRAG As Long = (LVN_FIRST - 11)
Private Const LVN_ODCACHEHINT As Long = (LVN_FIRST - 13)
Private Const LVN_ITEMACTIVATE As Long = (LVN_FIRST - 14)
Private Const LVN_ODSTATECHANGED As Long = (LVN_FIRST - 15)
Private Const LVN_HOTTRACK As Long = (LVN_FIRST - 21)
Private Const LVN_GETDISPINFOA As Long = (LVN_FIRST - 50)
Private Const LVN_GETDISPINFOW As Long = (LVN_FIRST - 77)
Private Const LVN_GETDISPINFO As Long = LVN_GETDISPINFOW
Private Const LVN_SETDISPINFOA As Long = (LVN_FIRST - 51)
Private Const LVN_SETDISPINFOW As Long = (LVN_FIRST - 78)
Private Const LVN_SETDISPINFO As Long = LVN_SETDISPINFOW
Private Const LVN_KEYDOWN As Long = (LVN_FIRST - 55)
Private Const LVN_MARQUEEBEGIN As Long = (LVN_FIRST - 56)
Private Const LVN_GETINFOTIPA As Long = (LVN_FIRST - 57)
Private Const LVN_GETINFOTIPW As Long = (LVN_FIRST - 58)
Private Const LVN_GETINFOTIP As Long = LVN_GETINFOTIPW
Private Const LVN_GETEMPTYMARKUP As Long = (LVN_FIRST - 87)
Private Const LVA_DEFAULT As Long = &H0
Private Const LVA_ALIGNLEFT As Long = &H1
Private Const LVA_ALIGNTOP As Long = &H2
Private Const LVA_SNAPTOGRID As Long = &H5
Private Const LVNI_ALL As Long = &H0
Private Const LVNI_FOCUSED As Long = &H1
Private Const LVNI_SELECTED As Long = &H2
Private Const LVNI_CUT As Long = &H4
Private Const LVNI_DROPHILITED As Long = &H8
Private Const LVNI_ABOVE As Long = &H100
Private Const LVNI_BELOW As Long = &H200
Private Const LVNI_TOLEFT As Long = &H400
Private Const LVNI_TORIGHT As Long = &H800
Private Const LVIF_TEXT As Long = &H1
Private Const LVIF_IMAGE As Long = &H2
Private Const LVIF_PARAM As Long = &H4
Private Const LVIF_STATE As Long = &H8
Private Const LVIF_INDENT As Long = &H10
Private Const LVIF_GROUPID As Long = &H100
Private Const LVIF_COLUMNS As Long = &H200
Private Const LVIF_NORECOMPUTE As Long = &H800
Private Const LVIR_BOUNDS As Long = 0
Private Const LVIR_ICON As Long = 1
Private Const LVIR_LABEL As Long = 2
Private Const LVIR_SELECTBOUNDS As Long = 3
Private Const LVIS_FOCUSED As Long = &H1
Private Const LVIS_SELECTED As Long = &H2
Private Const LVIS_CUT As Long = &H4
Private Const LVIS_DROPHILITED As Long = &H8
Private Const LVIS_ACTIVATING As Long = &H20
Private Const LVIS_OVERLAYMASK As Long = &HF00
Private Const LVIS_STATEIMAGEMASK As Long = &HF000
Private Const LVFI_PARAM As Long = &H1
Private Const LVFI_STRING As Long = &H2
Private Const LVFI_PARTIAL As Long = &H8
Private Const LVFI_WRAP As Long = &H20
Private Const LVFI_NEARESTXY As Long = &H40
Private Const LVKF_ALT As Long = &H1
Private Const LVKF_CONTROL As Long = &H2
Private Const LVKF_SHIFT As Long = &H4
Private Const LVBKIF_SOURCE_NONE As Long = &H0
Private Const LVBKIF_SOURCE_HBITMAP As Long = &H1
Private Const LVBKIF_SOURCE_URL As Long = &H2
Private Const LVBKIF_SOURCE_MASK As Long = &H3
Private Const LVBKIF_STYLE_NORMAL As Long = &H0
Private Const LVBKIF_STYLE_TILE As Long = &H10
Private Const LVBKIF_TYPE_WATERMARK As Long = &H10000000
Private Const LVBKIF_FLAG_TILEOFFSET As Long = &H100
Private Const LVBKIF_FLAG_ALPHABLEND As Long = &H20000000
Private Const LVGIT_UNFOLDED As Long = &H1
Private Const LVSIL_NORMAL As Long = 0
Private Const LVSIL_SMALL As Long = 1
Private Const LVSIL_STATE As Long = 2
Private Const LVHT_NOWHERE As Long = &H1
Private Const LVHT_ONITEMICON As Long = &H2
Private Const LVHT_ONITEMLABEL As Long = &H4
Private Const LVHT_ONITEMSTATEICON As Long = &H8
Private Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)
Private Const LVHT_ABOVE As Long = &H8
Private Const LVHT_BELOW As Long = &H10
Private Const LVHT_TORIGHT As Long = &H20
Private Const LVHT_TOLEFT As Long = &H40
Private Const LV_MAX_WORKAREAS As Long = 16
Private Const EMF_CENTERED As Long = 1
Private Const HDM_FIRST As Long = &H1200
Private Const HDM_SETIMAGELIST As Long = (HDM_FIRST + 8)
Private Const HDM_GETIMAGELIST As Long = (HDM_FIRST + 9)
Private Const HDSIL_NORMAL As Long = 0
Private Const HDSIL_STATE As Long = 0
Private Const HDF_SORTDOWN As Long = &H200
Private Const HDF_SORTUP As Long = &H400
Private Const HDF_BITMAP_ON_RIGHT As Long = &H1000
Private Const HDF_FIXEDWIDTH As Long = &H100
Private Const HDF_SPLITBUTTON As Long = &H1000000
Private Const HDF_CHECKBOX As Long = &H40
Private Const HDF_CHECKED As Long = &H80
Private Const HDS_BUTTONS As Long = &H2
Private Const HDS_HOTTRACK As Long = &H4
Private Const HDS_CHECKBOXES As Long = &H400
Private Const HDS_FULLDRAG As Long = &H80
Private Const HDS_NOSIZING As Long = &H800
Private Const HDN_FIRST As Long = (-300)
Private Const HDN_BEGINTRACKA As Long = (HDN_FIRST - 6)
Private Const HDN_BEGINTRACKW As Long = (HDN_FIRST - 26)
Private Const HDN_BEGINTRACK As Long = HDN_BEGINTRACKW
Private Const HDN_ENDTRACKA As Long = (HDN_FIRST - 7)
Private Const HDN_ENDTRACKW As Long = (HDN_FIRST - 27)
Private Const HDN_ENDTRACK As Long = HDN_ENDTRACKW
Private Const HDN_BEGINDRAG As Long = (HDN_FIRST - 10)
Private Const HDN_ENDDRAG As Long = (HDN_FIRST - 11)
Private Const HDN_ITEMSTATEICONCLICK As Long = (HDN_FIRST - 16)
Private Const HDN_DROPDOWN As Long = (HDN_FIRST - 18)
Private Const LVCF_FMT As Long = &H1
Private Const LVCF_WIDTH As Long = &H2
Private Const LVCF_TEXT As Long = &H4
Private Const LVCF_SUBITEM As Long = &H8
Private Const LVCF_IMAGE As Long = &H10
Private Const LVCF_ORDER As Long = &H20
Private Const LVSCW_AUTOSIZE As Long = (-1)
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = (-2)
Private Const LVIM_AFTER As Long = &H1
Private Const LVCFMT_LEFT As Long = &H0
Private Const LVCFMT_RIGHT As Long = &H1
Private Const LVCFMT_CENTER As Long = &H2
Private Const LVCFMT_JUSTIFYMASK As Long = &H3
Private Const LVCFMT_IMAGE As Long = &H800
Private Const LVCFMT_BITMAP_ON_RIGHT As Long = &H1000
Private Const LVCFMT_COL_HAS_IMAGES As Long = &H8000&
Private Const LVTVIM_TILESIZE As Long = &H1
Private Const LVTVIM_COLUMNS As Long = &H2
Private Const LVTVIM_LABELMARGIN As Long = &H4
Private Const I_IMAGECALLBACK As Long = (-1)
Private Const I_COLUMNSCALLBACK As Long = (-1)
Private Const H_MAX As Long = (&HFFFF + 1)
Private Const NM_FIRST As Long = H_MAX
Private Const NM_CLICK As Long = (NM_FIRST - 2)
Private Const NM_DBLCLK As Long = (NM_FIRST - 3)
Private Const NM_RCLICK As Long = (NM_FIRST - 5)
Private Const NM_RDBLCLK As Long = (NM_FIRST - 6)
Private Const NM_SETFOCUS As Long = (NM_FIRST - 7)
Private Const NM_KILLFOCUS As Long = (NM_FIRST - 8)
Private Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)
Private Const LVS_ICON As Long = &H0
Private Const LVS_REPORT As Long = &H1
Private Const LVS_SMALLICON As Long = &H2
Private Const LVS_LIST As Long = &H3
Private Const LVS_TYPEMASK As Long = &H3
Private Const LVS_SINGLESEL As Long = &H4
Private Const LVS_SHOWSELALWAYS As Long = &H8
Private Const LVS_SORTASCENDING As Long = &H10
Private Const LVS_SORTDESCENDING As Long = &H20
Private Const LVS_SHAREIMAGELISTS As Long = &H40
Private Const LVS_NOLABELWRAP As Long = &H80
Private Const LVS_AUTOARRANGE As Long = &H100
Private Const LVS_EDITLABELS As Long = &H200
Private Const LVS_OWNERDATA As Long = &H1000
Private Const LVS_NOSCROLL As Long = &H2000
Private Const LVS_TYPESTYLEMASK As Long = &HFC00
Private Const LVS_ALIGNTOP As Long = &H0
Private Const LVS_ALIGNLEFT As Long = &H800
Private Const LVS_ALIGNMASK As Long = &HC00
Private Const LVS_OWNERDRAWFIXED As Long = &H400
Private Const LVS_NOCOLUMNHEADER As Long = &H4000
Private Const LVS_NOSORTHEADER As Long = &H8000&
Private Const LVS_EX_GRIDLINES As Long = &H1
Private Const LVS_EX_HEADERDRAGDROP As Long = &H10
Private Const LVS_EX_FLATSB As Long = &H100
Private Const LVS_EX_DOUBLEBUFFER As Long = &H10000
Private Const LVS_EX_SUBITEMIMAGES As Long = &H2
Private Const LVS_EX_FULLROWSELECT As Long = &H20
Private Const LVS_EX_REGIONAL As Long = &H200
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVS_EX_ONECLICKACTIVATE As Long = &H40
Private Const LVS_EX_INFOTIP As Long = &H400
Private Const LVS_EX_LABELTIP As Long = &H4000
Private Const LVS_EX_TRACKSELECT As Long = &H8
Private Const LVS_EX_TWOCLICKACTIVATE As Long = &H80
Private Const LVS_EX_UNDERLINEHOT As Long = &H800
Private Const LVS_EX_BORDERSELECT As Long = &H8000&
Private Const LVS_EX_SNAPTOGRID As Long = &H80000
Private Const LV_VIEW_ICON As Long = &H0
Private Const LV_VIEW_DETAILS As Long = &H1
Private Const LV_VIEW_SMALLICON As Long = &H2
Private Const LV_VIEW_LIST As Long = &H3
Private Const LV_VIEW_TILE As Long = &H4
Implements ISubclass
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private ListViewHandle As Long, ListViewHeaderHandle As Long
Private ListViewFontHandle As Long, ListViewBoldFontHandle As Long, ListViewUnderlineFontHandle As Long, ListViewBoldUnderlineFontHandle As Long
Private ListViewLogFont As LOGFONT, ListViewBoldLogFont As LOGFONT, ListViewUnderlineLogFont As LOGFONT, ListViewBoldUnderlineLogFont As LOGFONT
Private ListViewFocusIndex As Long
Private ListViewLabelInEdit As Boolean
Private ListViewStartLabelEdit As Boolean
Private ListViewButtonDown As Integer
Private ListViewListItemsControl As Long
Private ListViewDragIndexBuffer As Long, ListViewDragIndex As Long
Private ListViewDragOffsetX As Long, ListViewDragOffsetY As Long
Private ListViewMemoryColumnWidth As Integer
Private DispIDMousePointer As Long
Private DispIDIcons As Long, IconsArray() As String
Private DispIDSmallIcons As Long, SmallIconsArray() As String
Private DispIDColumnHeaderIcons As Long, ColumnHeaderIconsArray() As String
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropListItems As LvwListItems
Private PropColumnHeaders As LvwColumnHeaders
Private PropVisualStyles As Boolean
Private PropOLEDragMode As VBRUN.OLEDragConstants
Private PropOLEDragDropScroll As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropIconsName As String, PropIconsControl As Object, PropIconsInit As Boolean
Private PropSmallIconsName As String, PropSmallIconsControl As Object, PropSmallIconsInit As Boolean
Private PropColumnHeaderIconsName As String, PropColumnHeaderIconsControl As Object, PropColumnHeaderIconsInit As Boolean
Private PropBorderStyle As CCBorderStyleConstants
Private PropBackColor As OLE_COLOR
Private PropForeColor As OLE_COLOR
Private PropRedraw As Boolean
Private PropView As LvwViewConstants
Private PropArrange As LvwArrangeConstants
Private PropAllowColumnReorder As Boolean
Private PropAllowColumnCheckboxes As Boolean
Private PropMultiSelect As Boolean
Private PropFullRowSelect As Boolean
Private PropGridLines As Boolean
Private PropLabelEdit As LvwLabelEditConstants
Private PropLabelWrap As Boolean
Private PropSorted As Boolean
Private PropSortKey As Integer
Private PropSortOrder As LvwSortOrderConstants
Private PropSortType As LvwSortTypeConstants
Private PropCheckboxes As Boolean
Private PropHideSelection As Boolean
Private PropHideColumnHeaders As Boolean
Private PropShowInfoTips As Boolean
Private PropShowLabelTips As Boolean
Private PropDoubleBuffer As Boolean
Private PropHoverSelection As Boolean
Private PropHoverSelectionTime As Long
Private PropHotTracking As Boolean
Private PropHighlightHot As Boolean
Private PropUnderlineHot As Boolean
Private PropInsertMarkColor As OLE_COLOR
Private PropTextBackground As CCBackStyleConstants
Private PropClickableColumnHeaders As Boolean
Private PropHighlightColumnHeaders As Boolean
Private PropTrackSizeColumnHeaders As Boolean
Private PropResizableColumnHeaders As Boolean
Private PropPicture As IPictureDisp
Private PropPictureAlignment As LvwPictureAlignmentConstants
Private PropPictureWatermark As Boolean
Private PropTileViewLines As Long
Private PropSnapToGrid As Boolean

Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
    Dim KeyCode As Integer, IsInputKey As Boolean
    KeyCode = wParam And &HFF&
    If wMsg = WM_KEYDOWN Then
        RaiseEvent PreviewKeyDown(KeyCode, IsInputKey)
    ElseIf wMsg = WM_KEYUP Then
        RaiseEvent PreviewKeyUp(KeyCode, IsInputKey)
    End If
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyReturn, vbKeyEscape
            If ListViewHandle <> 0 Then
                If ListViewLabelInEdit = True Then
                    SendMessage Me.hWndLabelEdit, wMsg, wParam, ByVal lParam
                Else
                    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape) And IsInputKey = False Then Exit Sub
                    SendMessage ListViewHandle, wMsg, wParam, ByVal lParam
                End If
                Handled = True
            End If
        Case vbKeyTab, vbKeyReturn, vbKeyEscape
            If IsInputKey = True Then
                If ListViewHandle <> 0 Then
                    SendMessage ListViewHandle, wMsg, wParam, ByVal lParam
                    Handled = True
                End If
            End If
    End Select
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispID As Long, ByRef DisplayName As String)
If DispID = DispIDMousePointer Then
    Call ComCtlsMousePointerSetDisplayString(PropMousePointer, DisplayName)
    Handled = True
ElseIf DispID = DispIDIcons Then
    DisplayName = PropIconsName
    Handled = True
ElseIf DispID = DispIDSmallIcons Then
    DisplayName = PropSmallIconsName
    Handled = True
ElseIf DispID = DispIDColumnHeaderIcons Then
    DisplayName = PropColumnHeaderIconsName
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDMousePointer Then
    Call ComCtlsMousePointerSetPredefinedStrings(StringsOut(), CookiesOut())
    Handled = True
ElseIf DispID = DispIDIcons Or DispID = DispIDSmallIcons Or DispID = DispIDColumnHeaderIcons Then
    Dim ControlEnum As Object
    Dim PropUBound As Long
    On Error GoTo CATCH_EXCEPTION
    PropUBound = UBound(StringsOut())
    ReDim Preserve StringsOut(PropUBound + 1) As String
    ReDim Preserve CookiesOut(PropUBound + 1) As Long
    StringsOut(PropUBound) = "(None)"
    CookiesOut(PropUBound) = PropUBound
    For Each ControlEnum In UserControl.ParentControls
        If TypeName(ControlEnum) = "ImageList" Then
            PropUBound = UBound(StringsOut())
            ReDim Preserve StringsOut(PropUBound + 1) As String
            ReDim Preserve CookiesOut(PropUBound + 1) As Long
            StringsOut(PropUBound) = ProperControlName(ControlEnum)
            CookiesOut(PropUBound) = PropUBound
        End If
    Next ControlEnum
    On Error GoTo 0
    Dim i As Long
    ReDim IconsArray(0 To UBound(StringsOut()))
    ReDim SmallIconsArray(0 To UBound(StringsOut()))
    ReDim ColumnHeaderIconsArray(0 To UBound(StringsOut()))
    For i = 0 To UBound(StringsOut())
        IconsArray(i) = StringsOut(i)
        SmallIconsArray(i) = StringsOut(i)
        ColumnHeaderIconsArray(i) = StringsOut(i)
    Next i
    Handled = True
End If
Exit Sub
CATCH_EXCEPTION:
Handled = False
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispID As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispID = DispIDMousePointer Then
    Value = Cookie
    Handled = True
ElseIf DispID = DispIDIcons Then
    If Cookie < UBound(IconsArray()) Then Value = IconsArray(Cookie)
    Handled = True
ElseIf DispID = DispIDSmallIcons Then
    If Cookie < UBound(SmallIconsArray()) Then Value = SmallIconsArray(Cookie)
    Handled = True
ElseIf DispID = DispIDColumnHeaderIcons Then
    If Cookie < UBound(ColumnHeaderIconsArray()) Then Value = ColumnHeaderIconsArray(Cookie)
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Dim ICCEX As TagInitCommonControlsEx
With ICCEX
.dwSize = LenB(ICCEX)
.dwICC = ICC_LISTVIEW_CLASSES
End With
InitCommonControlsEx ICCEX
Call SetVTableSubclass(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
DispIDMousePointer = GetDispID(Me, "MousePointer")
DispIDIcons = GetDispID(Me, "Icons")
DispIDSmallIcons = GetDispID(Me, "SmallIcons")
DispIDColumnHeaderIcons = GetDispID(Me, "ColumnHeaderIcons")
ReDim IconsArray(0) As String
ReDim SmallIconsArray(0) As String
ReDim ColumnHeaderIconsArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
Set PropFont = Ambient.Font
PropVisualStyles = True
PropOLEDragMode = vbOLEDragManual
PropOLEDragDropScroll = True
Me.OLEDropMode = vbOLEDropNone
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropIconsName = "(None)": Set PropIconsControl = Nothing
PropSmallIconsName = "(None)": Set PropSmallIconsControl = Nothing
PropColumnHeaderIconsName = "(None)": Set PropColumnHeaderIconsControl = Nothing
PropBorderStyle = CCBorderStyleSunken
PropBackColor = vbWindowBackground
PropForeColor = vbWindowText
PropRedraw = True
PropView = LvwViewIcon
PropArrange = LvwArrangeNone
PropAllowColumnReorder = False
PropAllowColumnCheckboxes = False
PropMultiSelect = False
PropFullRowSelect = False
PropGridLines = False
PropLabelEdit = LvwLabelEditAutomatic
PropLabelWrap = True
PropSorted = False
PropSortKey = 0
PropSortOrder = LvwSortOrderAscending
PropSortType = LvwSortTypeBinary
PropCheckboxes = False
PropHideSelection = True
PropHideColumnHeaders = False
PropShowInfoTips = False
PropShowLabelTips = False
PropDoubleBuffer = True
PropHoverSelection = False
PropHoverSelectionTime = -1
PropHotTracking = False
PropHighlightHot = False
PropUnderlineHot = False
PropInsertMarkColor = vbBlack
PropTextBackground = CCBackStyleTransparent
PropClickableColumnHeaders = True
PropHighlightColumnHeaders = False
PropTrackSizeColumnHeaders = True
PropResizableColumnHeaders = True
Set PropPicture = Nothing
PropPictureAlignment = LvwPictureAlignmentTopLeft
PropPictureWatermark = False
PropTileViewLines = 1
PropSnapToGrid = False
Call CreateListView
Me.FListItemsAdd 0, 1, Ambient.DisplayName
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
Set PropFont = .ReadProperty("Font", Ambient.Font)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Enabled = .ReadProperty("Enabled", True)
PropOLEDragMode = .ReadProperty("OLEDragMode", vbOLEDragManual)
PropOLEDragDropScroll = .ReadProperty("OLEDragDropScroll", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropIconsName = VarToStr(.ReadProperty("Icons", "(None)"))
PropSmallIconsName = VarToStr(.ReadProperty("SmallIcons", "(None)"))
PropColumnHeaderIconsName = VarToStr(.ReadProperty("ColumnHeaderIcons", "(None)"))
PropBorderStyle = .ReadProperty("BorderStyle", CCBorderStyleSunken)
PropBackColor = .ReadProperty("BackColor", vbWindowBackground)
PropForeColor = .ReadProperty("ForeColor", vbWindowText)
PropRedraw = .ReadProperty("Redraw", True)
PropView = .ReadProperty("View", LvwViewIcon)
PropArrange = .ReadProperty("Arrange", LvwArrangeNone)
PropAllowColumnReorder = .ReadProperty("AllowColumnReorder", False)
PropAllowColumnCheckboxes = .ReadProperty("AllowColumnCheckboxes", False)
PropMultiSelect = .ReadProperty("MultiSelect", False)
PropFullRowSelect = .ReadProperty("FullRowSelect", False)
PropGridLines = .ReadProperty("GridLines", False)
PropLabelEdit = .ReadProperty("LabelEdit", LvwLabelEditAutomatic)
PropLabelWrap = .ReadProperty("LabelWrap", True)
PropSorted = .ReadProperty("Sorted", False)
PropSortKey = .ReadProperty("SortKey", 0)
PropSortOrder = .ReadProperty("SortOrder", LvwSortOrderAscending)
PropSortType = .ReadProperty("SortType", LvwSortTypeBinary)
PropCheckboxes = .ReadProperty("Checkboxes", False)
PropHideSelection = .ReadProperty("HideSelection", True)
PropHideColumnHeaders = .ReadProperty("HideColumnHeaders", False)
PropShowInfoTips = .ReadProperty("ShowInfoTips", False)
PropShowLabelTips = .ReadProperty("ShowLabelTips", False)
PropDoubleBuffer = .ReadProperty("DoubleBuffer", True)
PropHoverSelection = .ReadProperty("HoverSelection", False)
PropHoverSelectionTime = .ReadProperty("HoverSelectionTime", -1)
PropHotTracking = .ReadProperty("HotTracking", False)
PropHighlightHot = .ReadProperty("HighlightHot", False)
PropUnderlineHot = .ReadProperty("UnderlineHot", False)
PropInsertMarkColor = .ReadProperty("InsertMarkColor", vbBlack)
PropTextBackground = .ReadProperty("TextBackground", CCBackStyleTransparent)
PropClickableColumnHeaders = .ReadProperty("ClickableColumnHeaders", True)
PropHighlightColumnHeaders = .ReadProperty("HighlightColumnHeaders", False)
PropTrackSizeColumnHeaders = .ReadProperty("TrackSizeColumnHeaders", True)
PropResizableColumnHeaders = .ReadProperty("ResizableColumnHeaders", True)
Set PropPicture = .ReadProperty("Picture", Nothing)
PropPictureAlignment = .ReadProperty("PictureAlignment", LvwPictureAlignmentTopLeft)
PropPictureWatermark = .ReadProperty("PictureWatermark", False)
PropTileViewLines = .ReadProperty("TileViewLines", 1)
PropSnapToGrid = .ReadProperty("SnapToGrid", False)
End With
If Ambient.UserMode = True Then
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 3)
    If Not PropIconsName = "(None)" Or Not PropSmallIconsName = "(None)" Or Not PropColumnHeaderIconsName = "(None)" Then TimerImageList.Enabled = True
Else
    Call CreateListView
    Me.FListItemsAdd 0, 2, Ambient.DisplayName
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", PropFont, Ambient.Font
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDragMode", PropOLEDragMode, vbOLEDragManual
.WriteProperty "OLEDragDropScroll", PropOLEDragDropScroll, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "Icons", StrToVar(PropIconsName), "(None)"
.WriteProperty "SmallIcons", StrToVar(PropSmallIconsName), "(None)"
.WriteProperty "ColumnHeaderIcons", StrToVar(PropColumnHeaderIconsName), "(None)"
.WriteProperty "BorderStyle", PropBorderStyle, CCBorderStyleSunken
.WriteProperty "BackColor", PropBackColor, vbWindowBackground
.WriteProperty "ForeColor", PropForeColor, vbWindowText
.WriteProperty "Redraw", PropRedraw, True
.WriteProperty "View", PropView, LvwViewIcon
.WriteProperty "Arrange", PropArrange, LvwArrangeNone
.WriteProperty "AllowColumnReorder", PropAllowColumnReorder, False
.WriteProperty "AllowColumnCheckboxes", PropAllowColumnCheckboxes, False
.WriteProperty "MultiSelect", PropMultiSelect, False
.WriteProperty "FullRowSelect", PropFullRowSelect, False
.WriteProperty "GridLines", PropGridLines, False
.WriteProperty "LabelEdit", PropLabelEdit, LvwLabelEditAutomatic
.WriteProperty "LabelWrap", PropLabelWrap, True
.WriteProperty "Sorted", PropSorted, False
.WriteProperty "SortKey", PropSortKey, 0
.WriteProperty "SortOrder", PropSortOrder, LvwSortOrderAscending
.WriteProperty "SortType", PropSortType, LvwSortTypeBinary
.WriteProperty "Checkboxes", PropCheckboxes, False
.WriteProperty "HideSelection", PropHideSelection, True
.WriteProperty "HideColumnHeaders", PropHideColumnHeaders, False
.WriteProperty "ShowInfoTips", PropShowInfoTips, False
.WriteProperty "ShowLabelTips", PropShowLabelTips, False
.WriteProperty "DoubleBuffer", PropDoubleBuffer, True
.WriteProperty "HoverSelection", PropHoverSelection, False
.WriteProperty "HoverSelectionTime", PropHoverSelectionTime, -1
.WriteProperty "HotTracking", PropHotTracking, False
.WriteProperty "HighlightHot", PropHighlightHot, False
.WriteProperty "UnderlineHot", PropUnderlineHot, False
.WriteProperty "InsertMarkColor", PropInsertMarkColor, vbBlack
.WriteProperty "TextBackground", PropTextBackground, CCBackStyleTransparent
.WriteProperty "ClickableColumnHeaders", PropClickableColumnHeaders, True
.WriteProperty "HighlightColumnHeaders", PropHighlightColumnHeaders, False
.WriteProperty "TrackSizeColumnHeaders", PropTrackSizeColumnHeaders, True
.WriteProperty "ResizableColumnHeaders", PropResizableColumnHeaders, True
.WriteProperty "Picture", PropPicture, Nothing
.WriteProperty "PictureAlignment", PropPictureAlignment, LvwPictureAlignmentTopLeft
.WriteProperty "PictureWatermark", PropPictureWatermark, False
.WriteProperty "TileViewLines", PropTileViewLines, 1
.WriteProperty "SnapToGrid", PropSnapToGrid, False
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
ListViewDragIndex = 0
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
If ListViewHandle <> 0 Then
    If ListViewDragIndex > 0 And Not Effect = vbDropEffectNone Then
        Select Case PropView
            Case LvwViewIcon, LvwViewSmallIcon, LvwViewTile
                Select Case PropArrange
                    Case LvwArrangeNone, LvwArrangeLeft, LvwArrangeTop
                    Case Else
                        Effect = vbDropEffectNone
                End Select
            Case Else
                Effect = vbDropEffectNone
        End Select
    End If
    If State = vbOver And Not Effect = vbDropEffectNone Then
        If PropOLEDragDropScroll = True Then
            Dim RC As RECT
            GetWindowRect ListViewHandle, RC
            Dim dwStyle As Long
            dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
            If (dwStyle And WS_HSCROLL) = WS_HSCROLL Then
                If Abs(X) < 16 Then
                    SendMessage ListViewHandle, WM_HSCROLL, SB_LINELEFT, ByVal 0&
                ElseIf Abs(X - (RC.Right - RC.Left)) < 16 Then
                    SendMessage ListViewHandle, WM_HSCROLL, SB_LINERIGHT, ByVal 0&
                End If
            End If
            If (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
                If Abs(Y) < 16 Then
                    SendMessage ListViewHandle, WM_VSCROLL, SB_LINEUP, ByVal 0&
                ElseIf Abs(Y - (RC.Bottom - RC.Top)) < 16 Then
                    SendMessage ListViewHandle, WM_VSCROLL, SB_LINEDOWN, ByVal 0&
                End If
            End If
        End If
    End If
    If ListViewDragIndex > 0 And Not Effect = vbDropEffectNone Then
        Select Case PropView
            Case LvwViewIcon, LvwViewSmallIcon, LvwViewTile
                Select Case PropArrange
                    Case LvwArrangeNone, LvwArrangeLeft, LvwArrangeTop
                        Dim ViewRect As RECT, P As POINTAPI
                        SendMessage ListViewHandle, LVM_GETVIEWRECT, 0, ByVal VarPtr(ViewRect)
                        P.X = X + (ListViewDragOffsetX - ViewRect.Left)
                        P.Y = Y + (ListViewDragOffsetY - ViewRect.Top)
                        SendMessage ListViewHandle, LVM_SETITEMPOSITION32, ListViewDragIndex - 1, ByVal VarPtr(P)
                End Select
        End Select
    End If
End If
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
If ListViewDragIndex > 0 Then
    If ListViewHandle <> 0 Then
        Dim P(0 To 1) As POINTAPI, RC As RECT
        GetCursorPos P(0)
        ScreenToClient ListViewHandle, P(0)
        SendMessage ListViewHandle, LVM_GETITEMPOSITION, ListViewDragIndex - 1, ByVal VarPtr(P(1))
        SendMessage ListViewHandle, LVM_GETVIEWRECT, 0, ByVal VarPtr(RC)
        ListViewDragOffsetY = (P(1).Y - P(0).Y) + RC.Top
        ListViewDragOffsetX = (P(1).X - P(0).X) + RC.Left
    End If
    If PropOLEDragMode = vbOLEDragAutomatic Then
        Dim Text As String
        Text = Me.FListItemText(ListViewDragIndex, 0)
        Data.SetData StrToVar(Text & vbNullChar), CF_UNICODETEXT
        Data.SetData StrToVar(Text), vbCFText
        AllowedEffects = vbDropEffectMove Or vbDropEffectCopy
    End If
End If
RaiseEvent OLEStartDrag(Data, AllowedEffects)
If AllowedEffects = vbDropEffectNone Then ListViewDragIndex = 0
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
If ListViewDragIndex > 0 Then Exit Sub
If ListViewDragIndexBuffer > 0 Then ListViewDragIndex = ListViewDragIndexBuffer
UserControl.OLEDrag
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
If Ambient.UserMode = False And PropertyName = "DisplayName" Then
    If ListViewHandle <> 0 Then
        If SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&) > 0 Then Me.FListItemText(1, 0) = Ambient.DisplayName
    End If
End If
End Sub

Private Sub UserControl_Resize()
If ListViewHandle = 0 Then Exit Sub
With UserControl
MoveWindow ListViewHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
End Sub

Private Sub UserControl_Hide()
If Not PropListItems Is Nothing Then
    On Error Resume Next
    If UserControl.Parent Is Nothing Then Set PropListItems = Nothing
    On Error GoTo 0
End If
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableSubclass(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyListView
Call ComCtlsRemoveSubclass(UserControl.hWnd)
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerImageList_Timer()
If PropIconsInit = False Then
    If Not PropIconsName = "(None)" Then Me.Icons = PropIconsName
    PropIconsInit = True
End If
If PropSmallIconsInit = False Then
    If Not PropSmallIconsName = "(None)" Then Me.SmallIcons = PropSmallIconsName
    PropSmallIconsInit = True
End If
If PropColumnHeaderIconsInit = False Then
    If Not PropColumnHeaderIconsName = "(None)" Then Me.ColumnHeaderIcons = PropColumnHeaderIconsName
    PropColumnHeaderIconsInit = True
End If
TimerImageList.Enabled = False
End Sub

Public Property Get ControlsEnum() As Object
Attribute ControlsEnum.VB_MemberFlags = "40"
Set ControlsEnum = UserControl.ParentControls
End Property

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name used in code to identify an object."
Name = Ambient.DisplayName
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal Value As String)
Extender.Tag = Value
End Property

Public Property Get Parent() As Object
Attribute Parent.VB_Description = "Returns the object on which this object is located."
Set Parent = UserControl.Parent
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
Left = Extender.Left
End Property

Public Property Let Left(ByVal Value As Single)
Extender.Left = Value
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns/sets the distance between the internal top edge of an object and the top edge of its container."
Top = Extender.Top
End Property

Public Property Let Top(ByVal Value As Single)
Extender.Top = Value
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of an object."
Width = Extender.Width
End Property

Public Property Let Width(ByVal Value As Single)
Extender.Width = Value
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns/sets the height of an object."
Height = Extender.Height
End Property

Public Property Let Height(ByVal Value As Single)
Extender.Height = Value
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
hWnd = ListViewHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get hWndHeader() As Long
Attribute hWndHeader.VB_Description = "Returns a handle to a control."
If ListViewHandle <> 0 Then hWndHeader = SendMessage(ListViewHandle, LVM_GETHEADER, 0, ByVal 0&)
End Property

Public Property Get hWndLabelEdit() As Long
Attribute hWndLabelEdit.VB_Description = "Returns a handle to a control."
If ListViewHandle <> 0 Then hWndLabelEdit = SendMessage(ListViewHandle, LVM_GETEDITCONTROL, 0, ByVal 0&)
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
Set Font = PropFont
End Property

Public Property Let Font(ByVal NewFont As StdFont)
Set Me.Font = NewFont
End Property

Public Property Set Font(ByVal NewFont As StdFont)
Dim OldFontHandle As Long, OldBoldFontHandle As Long, OldUnderlineFontHandle As Long, OldBoldUnderlineFontHandle As Long
Set PropFont = NewFont
Call OLEFontToLogFont(NewFont, ListViewLogFont)
LSet ListViewBoldLogFont = ListViewLogFont
LSet ListViewUnderlineLogFont = ListViewLogFont
LSet ListViewBoldUnderlineLogFont = ListViewLogFont
ListViewBoldLogFont.LFWeight = FW_BOLD
ListViewUnderlineLogFont.LFUnderline = 1
ListViewBoldUnderlineLogFont.LFWeight = FW_BOLD
ListViewBoldUnderlineLogFont.LFUnderline = 1
OldFontHandle = ListViewFontHandle
OldBoldFontHandle = ListViewBoldFontHandle
OldUnderlineFontHandle = ListViewUnderlineFontHandle
OldBoldUnderlineFontHandle = ListViewBoldUnderlineFontHandle
ListViewFontHandle = CreateFontIndirect(ListViewLogFont)
ListViewBoldFontHandle = CreateFontIndirect(ListViewBoldLogFont)
ListViewUnderlineFontHandle = CreateFontIndirect(ListViewUnderlineLogFont)
ListViewBoldUnderlineFontHandle = CreateFontIndirect(ListViewBoldUnderlineLogFont)
If ListViewHandle <> 0 Then SendMessage ListViewHandle, WM_SETFONT, ListViewFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
If OldBoldFontHandle <> 0 Then DeleteObject OldBoldFontHandle
If OldUnderlineFontHandle <> 0 Then DeleteObject OldUnderlineFontHandle
If OldBoldUnderlineFontHandle <> 0 Then DeleteObject OldBoldUnderlineFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long, OldBoldFontHandle As Long, OldUnderlineFontHandle As Long, OldBoldUnderlineFontHandle As Long
Call OLEFontToLogFont(PropFont, ListViewLogFont)
LSet ListViewBoldLogFont = ListViewLogFont
LSet ListViewUnderlineLogFont = ListViewLogFont
LSet ListViewBoldUnderlineLogFont = ListViewLogFont
ListViewBoldLogFont.LFWeight = FW_BOLD
ListViewUnderlineLogFont.LFUnderline = 1
ListViewBoldUnderlineLogFont.LFWeight = FW_BOLD
ListViewBoldUnderlineLogFont.LFUnderline = 1
OldFontHandle = ListViewFontHandle
OldBoldFontHandle = ListViewBoldFontHandle
OldUnderlineFontHandle = ListViewUnderlineFontHandle
OldBoldUnderlineFontHandle = ListViewBoldUnderlineFontHandle
ListViewFontHandle = CreateFontIndirect(ListViewLogFont)
ListViewBoldFontHandle = CreateFontIndirect(ListViewBoldLogFont)
ListViewUnderlineFontHandle = CreateFontIndirect(ListViewUnderlineLogFont)
ListViewBoldUnderlineFontHandle = CreateFontIndirect(ListViewBoldUnderlineLogFont)
If ListViewHandle <> 0 Then SendMessage ListViewHandle, WM_SETFONT, ListViewFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
If OldBoldFontHandle <> 0 Then DeleteObject OldBoldFontHandle
If OldUnderlineFontHandle <> 0 Then DeleteObject OldUnderlineFontHandle
If OldBoldUnderlineFontHandle <> 0 Then DeleteObject OldBoldUnderlineFontHandle
UserControl.PropertyChanged "Font"
End Sub

Private Sub OLEFontToLogFont(ByVal Font As StdFont, ByRef LF As LOGFONT)
Dim FontName As String
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
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If ListViewHandle <> 0 And EnabledVisualStyles() = True Then
    Select Case PropVisualStyles
        Case True
            ActivateVisualStyles ListViewHandle
        Case False
            RemoveVisualStyles ListViewHandle
    End Select
    Call SetVisualStylesHeader
    SendMessage ListViewHandle, LVM_UPDATE, 0, ByVal 0&
    Me.Refresh
    If ComCtlsSupportLevel() >= 2 Then
        If Not PropPicture Is Nothing Then
            If PropPictureAlignment = LvwPictureAlignmentTile Then Set Me.Picture = PropPicture
        End If
    End If
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If ListViewHandle <> 0 Then EnableWindow ListViewHandle, IIf(Value = True, 1, 0)
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDragMode() As VBRUN.OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this control can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
OLEDragMode = PropOLEDragMode
End Property

Public Property Let OLEDragMode(ByVal Value As VBRUN.OLEDragConstants)
Select Case Value
    Case vbOLEDragManual, vbOLEDragAutomatic
        PropOLEDragMode = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "OLEDragMode"
End Property

Public Property Get OLEDragDropScroll() As Boolean
Attribute OLEDragDropScroll.VB_Description = "Returns/Sets whether this object will scroll during an OLE drag/drop operation."
OLEDragDropScroll = PropOLEDragDropScroll
End Property

Public Property Let OLEDragDropScroll(ByVal Value As Boolean)
PropOLEDragDropScroll = Value
UserControl.PropertyChanged "OLEDragDropScroll"
End Property

Public Property Get OLEDropMode() As OLEDropModeConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As OLEDropModeConstants)
Select Case Value
    Case OLEDropModeNone, OLEDropModeManual
        UserControl.OLEDropMode = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "OLEDropMode"
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
MousePointer = PropMousePointer
End Property

Public Property Let MousePointer(ByVal Value As Integer)
Select Case Value
    Case 0 To 16, 99
        PropMousePointer = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Returns/sets a custom mouse icon."
Set MouseIcon = PropMouseIcon
End Property

Public Property Let MouseIcon(ByVal Value As IPictureDisp)
Set Me.MouseIcon = Value
End Property

Public Property Set MouseIcon(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropMouseIcon = Nothing
Else
    If Value.Type = vbPicTypeIcon Or Value.Handle = 0 Then
        Set PropMouseIcon = Value
    Else
        If Ambient.UserMode = False Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
UserControl.PropertyChanged "MouseIcon"
End Property

Public Property Get Icons() As Variant
Attribute Icons.VB_Description = "Returns/sets the image list control to be used for the icons."
If Ambient.UserMode = True Then
    If PropIconsInit = False And PropIconsControl Is Nothing Then
        If Not PropIconsName = "(None)" Then Me.Icons = PropIconsName
        PropIconsInit = True
    End If
    Set Icons = PropIconsControl
Else
    Icons = PropIconsName
End If
End Property

Public Property Set Icons(ByVal Value As Variant)
Me.Icons = Value
End Property

Public Property Let Icons(ByVal Value As Variant)
If Ambient.UserMode = True Then
    If ListViewHandle <> 0 Then
        Dim Success As Boolean, Handle As Long
        On Error Resume Next
        If IsObject(Value) Then
            If TypeName(Value) = "ImageList" Then
                Handle = Value.hImageList
                Success = CBool(Err.Number = 0 And Handle <> 0)
            End If
            If Success = True Then
                SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_NORMAL, ByVal Handle
                PropIconsName = ProperControlName(Value)
                Set PropIconsControl = Value
            End If
        ElseIf VarType(Value) = vbString Then
            Dim ControlEnum As Object, CompareName As String
            For Each ControlEnum In UserControl.ParentControls
                If TypeName(ControlEnum) = "ImageList" Then
                    CompareName = ProperControlName(ControlEnum)
                    If CompareName = Value And Not CompareName = vbNullString Then
                        Err.Clear
                        Handle = ControlEnum.hImageList
                        Success = CBool(Err.Number = 0 And Handle <> 0)
                        If Success = True Then
                            SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_NORMAL, ByVal Handle
                            PropIconsName = Value
                            Set PropIconsControl = ControlEnum
                            Exit For
                        ElseIf Ambient.UserMode = False Then
                            PropIconsName = Value
                            Success = True
                            Exit For
                        End If
                    End If
                End If
            Next ControlEnum
        End If
        On Error GoTo 0
        If Success = False Then
            SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_NORMAL, ByVal 0&
            PropIconsName = "(None)"
            Set PropIconsControl = Nothing
        Else
            SendMessage ListViewHandle, LVM_ARRANGE, LVA_DEFAULT, ByVal 0&
            SendMessage ListViewHandle, LVM_UPDATE, 0, ByVal 0&
        End If
    End If
Else
    PropIconsName = Value
End If
UserControl.PropertyChanged "Icons"
End Property

Public Property Get SmallIcons() As Variant
Attribute SmallIcons.VB_Description = "Returns/sets the image list control to be used for the small icons."
If Ambient.UserMode = True Then
    If PropSmallIconsInit = False And PropSmallIconsControl Is Nothing Then
        If Not PropSmallIconsName = "(None)" Then Me.SmallIcons = PropSmallIconsName
        PropSmallIconsInit = True
    End If
    Set SmallIcons = PropSmallIconsControl
Else
    SmallIcons = PropSmallIconsName
End If
End Property

Public Property Set SmallIcons(ByVal Value As Variant)
Me.SmallIcons = Value
End Property

Public Property Let SmallIcons(ByVal Value As Variant)
If Ambient.UserMode = True Then
    If ListViewHandle <> 0 Then
        Dim Success As Boolean, Handle As Long
        On Error Resume Next
        If IsObject(Value) Then
            If TypeName(Value) = "ImageList" Then
                Handle = Value.hImageList
                Success = CBool(Err.Number = 0 And Handle <> 0)
            End If
            If Success = True Then
                SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal Handle
                PropSmallIconsName = ProperControlName(Value)
                Set PropSmallIconsControl = Value
            End If
        ElseIf VarType(Value) = vbString Then
            Dim ControlEnum As Object, CompareName As String
            For Each ControlEnum In UserControl.ParentControls
                If TypeName(ControlEnum) = "ImageList" Then
                    CompareName = ProperControlName(ControlEnum)
                    If CompareName = Value And Not CompareName = vbNullString Then
                        Err.Clear
                        Handle = ControlEnum.hImageList
                        Success = CBool(Err.Number = 0 And Handle <> 0)
                        If Success = True Then
                            SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal Handle
                            PropSmallIconsName = Value
                            If Ambient.UserMode = True Then Set PropSmallIconsControl = ControlEnum
                            Exit For
                        ElseIf Ambient.UserMode = False Then
                            PropSmallIconsName = Value
                            Success = True
                            Exit For
                        End If
                    End If
                End If
            Next ControlEnum
        End If
        On Error GoTo 0
        If Success = False Then
            SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal 0&
            PropSmallIconsName = "(None)"
            Set PropSmallIconsControl = Nothing
        Else
            SendMessage ListViewHandle, LVM_ARRANGE, LVA_DEFAULT, ByVal 0&
            SendMessage ListViewHandle, LVM_UPDATE, 0, ByVal 0&
        End If
        ' The image list for the column icons need to be reset, because
        ' LVM_SETIMAGELIST with LVSIL_SMALL overrides the image list for the column icons.
        If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
        If ListViewHeaderHandle <> 0 Then
            If Not PropColumnHeaderIconsControl Is Nothing Then
                Dim ImageListHandle As Long
                ImageListHandle = PropColumnHeaderIconsControl.hImageList
                SendMessage ListViewHeaderHandle, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal ImageListHandle
                RedrawWindow ListViewHeaderHandle, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
            Else
                SendMessage ListViewHeaderHandle, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal 0&
            End If
        End If
    End If
Else
    PropSmallIconsName = Value
End If
UserControl.PropertyChanged "SmallIcons"
End Property

Public Property Get ColumnHeaderIcons() As Variant
Attribute ColumnHeaderIcons.VB_Description = "Returns/sets the image list control to be used for the column header icons."
If Ambient.UserMode = True Then
    If PropColumnHeaderIconsInit = False And PropColumnHeaderIconsControl Is Nothing Then
        If Not PropColumnHeaderIconsName = "(None)" Then Me.ColumnHeaderIcons = PropColumnHeaderIconsName
        PropColumnHeaderIconsInit = True
    End If
    Set ColumnHeaderIcons = PropColumnHeaderIconsControl
Else
    ColumnHeaderIcons = PropColumnHeaderIconsName
End If
End Property

Public Property Set ColumnHeaderIcons(ByVal Value As Variant)
Me.ColumnHeaderIcons = Value
End Property

Public Property Let ColumnHeaderIcons(ByVal Value As Variant)
If Ambient.UserMode = True Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHandle <> 0 And ListViewHeaderHandle <> 0 Then
        Dim Success As Boolean, Handle As Long
        On Error Resume Next
        If IsObject(Value) Then
            If TypeName(Value) = "ImageList" Then
                Handle = Value.hImageList
                Success = CBool(Err.Number = 0 And Handle <> 0)
            End If
            If Success = True Then
                SendMessage ListViewHeaderHandle, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal Handle
                PropColumnHeaderIconsName = ProperControlName(Value)
                Set PropColumnHeaderIconsControl = Value
            End If
        ElseIf VarType(Value) = vbString Then
            Dim ControlEnum As Object, CompareName As String
            For Each ControlEnum In UserControl.ParentControls
                If TypeName(ControlEnum) = "ImageList" Then
                    CompareName = ProperControlName(ControlEnum)
                    If CompareName = Value And Not CompareName = vbNullString Then
                        Err.Clear
                        Handle = ControlEnum.hImageList
                        Success = CBool(Err.Number = 0 And Handle <> 0)
                        If Success = True Then
                            SendMessage ListViewHeaderHandle, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal Handle
                            PropColumnHeaderIconsName = Value
                            Set PropColumnHeaderIconsControl = ControlEnum
                            Exit For
                        ElseIf Ambient.UserMode = False Then
                            PropColumnHeaderIconsName = Value
                            Success = True
                            Exit For
                        End If
                    End If
                End If
            Next ControlEnum
        End If
        On Error GoTo 0
        If Success = False Then
            SendMessage ListViewHeaderHandle, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal 0&
            PropColumnHeaderIconsName = "(None)"
            Set PropColumnHeaderIconsControl = Nothing
        End If
        If Me.ColumnHeaders.Count > 0 Then
            Dim i As Long, Icon As Long
            For i = 1 To Me.ColumnHeaders.Count
                Icon = Me.FColumnHeaderIcon(i)
                If Icon > 0 Then Me.FColumnHeaderIcon(i) = Icon
            Next i
        End If
    End If
Else
    PropColumnHeaderIconsName = Value
End If
UserControl.PropertyChanged "ColumnHeaderIcons"
End Property

Public Property Get BorderStyle() As CCBorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style."
Attribute BorderStyle.VB_UserMemId = -504
BorderStyle = PropBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As CCBorderStyleConstants)
Select Case Value
    Case CCBorderStyleNone, CCBorderStyleSingle, CCBorderStyleThin, CCBorderStyleSunken, CCBorderStyleRaised
        PropBorderStyle = Value
    Case Else
        Err.Raise 380
End Select
If ListViewHandle <> 0 Then Call ComCtlsChangeBorderStyle(ListViewHandle, PropBorderStyle)
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If ListViewHandle <> 0 Then
    If Value <> CLR_NONE Then
        SendMessage ListViewHandle, LVM_SETBKCOLOR, 0, ByVal WinColor(PropBackColor)
    Else
        Err.Raise 380
    End If
    Me.Refresh
    If Not PropPicture Is Nothing Then
        If PropPicture.Type = vbPicTypeIcon Then Set Me.Picture = PropPicture
    End If
End If
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, LVM_SETTEXTCOLOR, 0, ByVal WinColor(PropForeColor)
    Me.Refresh
End If
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Returns/sets a value that determines whether or not the list view redraws when changing the list items. You can speed up the creation of large lists by disabling this property before adding the list items."
Redraw = PropRedraw
End Property

Public Property Let Redraw(ByVal Value As Boolean)
PropRedraw = Value
If ListViewHandle <> 0 And Ambient.UserMode = True Then
    SendMessage ListViewHandle, WM_SETREDRAW, IIf(PropRedraw = True, 1, 0), ByVal 0&
    If PropRedraw = True Then Me.Refresh
End If
End Property

Public Property Get View() As LvwViewConstants
Attribute View.VB_Description = "Returns/sets the current view."
View = PropView
End Property

Public Property Let View(ByVal Value As LvwViewConstants)
Select Case Value
    Case LvwViewIcon, LvwViewSmallIcon, LvwViewList, LvwViewReport, LvwViewTile
        PropView = Value
    Case Else
        Err.Raise 380
End Select
If ListViewHandle <> 0 And Ambient.UserMode = True Then
    If ComCtlsSupportLevel() >= 1 Then
        Dim NewView As Long
        Select Case PropView
            Case LvwViewIcon
                NewView = LV_VIEW_ICON
            Case LvwViewSmallIcon
                NewView = LV_VIEW_SMALLICON
            Case LvwViewList
                NewView = LV_VIEW_LIST
            Case LvwViewReport
                NewView = LV_VIEW_DETAILS
            Case LvwViewTile
                NewView = LV_VIEW_TILE
        End Select
        SendMessage ListViewHandle, LVM_SETVIEW, NewView, ByVal 0&
    Else
        If PropView = LvwViewTile Then PropView = LvwViewIcon
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
        dwStyle = dwStyle And Not LVS_TYPEMASK
        Select Case PropView
            Case LvwViewIcon
                dwStyle = dwStyle Or LVS_ICON
            Case LvwViewSmallIcon
                dwStyle = dwStyle Or LVS_SMALLICON
            Case LvwViewList
                dwStyle = dwStyle Or LVS_LIST
            Case LvwViewReport
                dwStyle = dwStyle Or LVS_REPORT
        End Select
        SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
    End If
    If PropView = LvwViewList Then
        If ListViewMemoryColumnWidth <> 0 Then SendMessage ListViewHandle, LVM_SETCOLUMNWIDTH, 0, ByVal CLng(ListViewMemoryColumnWidth)
    ElseIf PropView = LvwViewReport Then
        Call CheckHeaderControl
    End If
    If ComCtlsSupportLevel() >= 2 Then
        If Not PropPicture Is Nothing Then
            If PropPictureAlignment = LvwPictureAlignmentTile Then Set Me.Picture = PropPicture
        End If
    End If
End If
UserControl.PropertyChanged "View"
End Property

Public Property Get Arrange() As LvwArrangeConstants
Attribute Arrange.VB_Description = "Returns/sets a value indicating how the icons in a 'icon', 'small icon' or 'tile' view are arranged."
Arrange = PropArrange
End Property

Public Property Let Arrange(ByVal Value As LvwArrangeConstants)
Select Case Value
    Case LvwArrangeNone, LvwArrangeAutoLeft, LvwArrangeAutoTop, LvwArrangeLeft, LvwArrangeTop
        PropArrange = Value
    Case Else
        Err.Raise 380
End Select
If ListViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    If (dwStyle And LVS_AUTOARRANGE) = LVS_AUTOARRANGE Then dwStyle = dwStyle And Not LVS_AUTOARRANGE
    If (dwStyle And LVS_ALIGNLEFT) = LVS_ALIGNLEFT Then dwStyle = dwStyle And Not LVS_ALIGNLEFT
    If (dwStyle And LVS_ALIGNTOP) = LVS_ALIGNTOP Then dwStyle = dwStyle And Not LVS_ALIGNTOP
    Select Case PropArrange
        Case LvwArrangeAutoLeft
            dwStyle = dwStyle Or LVS_AUTOARRANGE Or LVS_ALIGNLEFT
        Case LvwArrangeAutoTop
            dwStyle = dwStyle Or LVS_AUTOARRANGE Or LVS_ALIGNTOP
        Case LvwArrangeLeft
            dwStyle = dwStyle Or LVS_ALIGNLEFT
        Case LvwArrangeTop
            dwStyle = dwStyle Or LVS_ALIGNTOP
    End Select
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "Arrange"
End Property

Public Property Get AllowColumnReorder() As Boolean
Attribute AllowColumnReorder.VB_Description = "Returns/sets a value that determines whether or not a user can reorder column headers in 'report' view."
AllowColumnReorder = PropAllowColumnReorder
End Property

Public Property Let AllowColumnReorder(ByVal Value As Boolean)
PropAllowColumnReorder = Value
If ListViewHandle <> 0 Then
    If PropAllowColumnReorder = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_HEADERDRAGDROP, ByVal LVS_EX_HEADERDRAGDROP
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_HEADERDRAGDROP, ByVal 0&
    End If
End If
UserControl.PropertyChanged "AllowColumnReorder"
End Property

Public Property Get AllowColumnCheckboxes() As Boolean
Attribute AllowColumnCheckboxes.VB_Description = "Returns/sets a value that determines whether or not the column headers in 'report' view are allowed to place checkboxes. Requires comctl32.dll version 6.1 or higher."
AllowColumnCheckboxes = PropAllowColumnCheckboxes
End Property

Public Property Let AllowColumnCheckboxes(ByVal Value As Boolean)
PropAllowColumnCheckboxes = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropAllowColumnCheckboxes = CBool((dwStyle And HDS_CHECKBOXES) = HDS_CHECKBOXES) Then
            If PropAllowColumnCheckboxes = True Then
                If Not (dwStyle And HDS_CHECKBOXES) = HDS_CHECKBOXES Then dwStyle = dwStyle Or HDS_CHECKBOXES
            Else
                If (dwStyle And HDS_CHECKBOXES) = HDS_CHECKBOXES Then dwStyle = dwStyle And Not HDS_CHECKBOXES
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
        End If
    End If
End If
UserControl.PropertyChanged "AllowColumnCheckboxes"
End Property

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Returns/sets a value indicating whether a user can make multiple selections in the list view and how the multiple selections can be made."
MultiSelect = PropMultiSelect
End Property

Public Property Let MultiSelect(ByVal Value As Boolean)
PropMultiSelect = Value
If ListViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    If PropMultiSelect = True Then
        If (dwStyle And LVS_SINGLESEL) = LVS_SINGLESEL Then dwStyle = dwStyle And Not LVS_SINGLESEL
    Else
        If Not (dwStyle And LVS_SINGLESEL) = LVS_SINGLESEL Then dwStyle = dwStyle Or LVS_SINGLESEL
    End If
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
    If PropMultiSelect = False Then
        Dim ListItem As LvwListItem
        Set ListItem = Me.SelectedItem
        If Not ListItem Is Nothing Then ListItem.Selected = True
    End If
End If
UserControl.PropertyChanged "MultiSelect"
End Property

Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Returns/sets whether selecting a list item highlights the entire row in 'report' view."
FullRowSelect = PropFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal Value As Boolean)
PropFullRowSelect = Value
If ListViewHandle <> 0 Then
    If PropFullRowSelect = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, ByVal LVS_EX_FULLROWSELECT
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, ByVal 0&
    End If
End If
UserControl.PropertyChanged "FullRowSelect"
End Property

Public Property Get GridLines() As Boolean
Attribute GridLines.VB_Description = "Returns/sets whether grid lines appear between rows and columns in 'report' view."
GridLines = PropGridLines
End Property

Public Property Let GridLines(ByVal Value As Boolean)
PropGridLines = Value
If ListViewHandle <> 0 Then
    If PropGridLines = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, ByVal LVS_EX_GRIDLINES
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, ByVal 0&
    End If
End If
UserControl.PropertyChanged "GridLines"
End Property

Public Property Get LabelEdit() As LvwLabelEditConstants
Attribute LabelEdit.VB_Description = "Returns/sets a value that determines if a user can edit the label of a list item."
LabelEdit = PropLabelEdit
End Property

Public Property Let LabelEdit(ByVal Value As LvwLabelEditConstants)
Select Case Value
    Case LvwLabelEditAutomatic, LvwLabelEditManual, LvwLabelEditDisabled
        PropLabelEdit = Value
    Case Else
        Err.Raise 380
End Select
If ListViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    Select Case PropLabelEdit
        Case LvwLabelEditAutomatic, LvwLabelEditManual
            If Not (dwStyle And LVS_EDITLABELS) = LVS_EDITLABELS Then dwStyle = dwStyle Or LVS_EDITLABELS
        Case LvwLabelEditDisabled
            If (dwStyle And LVS_EDITLABELS) = LVS_EDITLABELS Then dwStyle = dwStyle And Not LVS_EDITLABELS
    End Select
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "LabelEdit"
End Property

Public Property Get LabelWrap() As Boolean
Attribute LabelWrap.VB_Description = "Returns/sets a value that determines if labels are wrapped when the list view is in icon view."
LabelWrap = PropLabelWrap
End Property

Public Property Let LabelWrap(ByVal Value As Boolean)
PropLabelWrap = Value
If ListViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    If PropLabelWrap = True Then
        If (dwStyle And LVS_NOLABELWRAP) = LVS_NOLABELWRAP Then dwStyle = dwStyle And Not LVS_NOLABELWRAP
    Else
        If Not (dwStyle And LVS_NOLABELWRAP) = LVS_NOLABELWRAP Then dwStyle = dwStyle Or LVS_NOLABELWRAP
    End If
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
    SendMessage ListViewHandle, LVM_UPDATE, 0, ByVal 0&
End If
UserControl.PropertyChanged "LabelWrap"
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Returns/sets a value indicating if the list items are automatically sorted."
Sorted = PropSorted
End Property

Public Property Let Sorted(ByVal Value As Boolean)
PropSorted = Value
If PropSorted = True And Ambient.UserMode = True Then Call SortListItems
UserControl.PropertyChanged "Sorted"
End Property

Public Property Get SortKey() As Integer
Attribute SortKey.VB_Description = "Returns/sets the current sort key."
SortKey = PropSortKey
End Property

Public Property Let SortKey(ByVal Value As Integer)
If Value < 0 Then
    If Ambient.UserMode = False Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropSortKey = Value
If PropSorted = True And Ambient.UserMode = True Then Call SortListItems
UserControl.PropertyChanged "SortKey"
End Property

Public Property Get SortOrder() As LvwSortOrderConstants
Attribute SortOrder.VB_Description = "Returns/sets a value that determines whether the list items will be sorted in ascending or descending order."
SortOrder = PropSortOrder
End Property

Public Property Let SortOrder(ByVal Value As LvwSortOrderConstants)
Select Case Value
    Case LvwSortOrderAscending, LvwSortOrderDescending
        PropSortOrder = Value
    Case Else
        Err.Raise 380
End Select
If PropSorted = True And Ambient.UserMode = True Then Call SortListItems
UserControl.PropertyChanged "SortOrder"
End Property

Public Property Get SortType() As LvwSortTypeConstants
Attribute SortType.VB_Description = "Returns/sets the sort type."
SortType = PropSortType
End Property

Public Property Let SortType(ByVal Value As LvwSortTypeConstants)
Select Case Value
    Case LvwSortTypeBinary, LvwSortTypeText, LvwSortTypeNumeric, LvwSortTypeCurrency, LvwSortTypeDate
        PropSortType = Value
    Case Else
        Err.Raise 380
End Select
If PropSorted = True And Ambient.UserMode = True Then Call SortListItems
UserControl.PropertyChanged "SortType"
End Property

Public Property Get Checkboxes() As Boolean
Attribute Checkboxes.VB_Description = "Returns/sets a value that determines whether or not a checkbox is displayed next to each list item."
Checkboxes = PropCheckboxes
End Property

Public Property Let Checkboxes(ByVal Value As Boolean)
PropCheckboxes = Value
If ListViewHandle <> 0 Then
    If PropCheckboxes = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_CHECKBOXES, ByVal LVS_EX_CHECKBOXES
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_CHECKBOXES, ByVal 0&
    End If
End If
UserControl.PropertyChanged "Checkboxes"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value that determines whether the selected item will display as selected when the list view loses focus or not."
HideSelection = PropHideSelection
End Property

Public Property Let HideSelection(ByVal Value As Boolean)
PropHideSelection = Value
If ListViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    If PropHideSelection = True Then
        If (dwStyle And LVS_SHOWSELALWAYS) = LVS_SHOWSELALWAYS Then dwStyle = dwStyle And Not LVS_SHOWSELALWAYS
    Else
        If Not (dwStyle And LVS_SHOWSELALWAYS) = LVS_SHOWSELALWAYS Then dwStyle = dwStyle Or LVS_SHOWSELALWAYS
    End If
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "HideSelection"
End Property

Public Property Get HideColumnHeaders() As Boolean
Attribute HideColumnHeaders.VB_Description = "Returns/sets a value that determines whether or not the column headers are hidden in 'report' view."
HideColumnHeaders = PropHideColumnHeaders
End Property

Public Property Let HideColumnHeaders(ByVal Value As Boolean)
PropHideColumnHeaders = Value
If ListViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    If PropHideColumnHeaders = True Then
        If Not (dwStyle And LVS_NOCOLUMNHEADER) = LVS_NOCOLUMNHEADER Then dwStyle = dwStyle Or LVS_NOCOLUMNHEADER
    Else
        If (dwStyle And LVS_NOCOLUMNHEADER) = LVS_NOCOLUMNHEADER Then dwStyle = dwStyle And Not LVS_NOCOLUMNHEADER
    End If
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "HideColumnHeaders"
End Property

Public Property Get ShowInfoTips() As Boolean
Attribute ShowInfoTips.VB_Description = "Returns/sets a value that determines whether the tool tip text properties will be displayed or not."
ShowInfoTips = PropShowInfoTips
End Property

Public Property Let ShowInfoTips(ByVal Value As Boolean)
PropShowInfoTips = Value
If ListViewHandle <> 0 Then
    If PropShowInfoTips = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_INFOTIP, ByVal LVS_EX_INFOTIP
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_INFOTIP, ByVal 0&
    End If
End If
UserControl.PropertyChanged "ShowInfoTips"
End Property

Public Property Get ShowLabelTips() As Boolean
Attribute ShowLabelTips.VB_Description = "Returns/sets a value indicating that if a partially hidden label in any list view mode lacks tool tip text, the list view will unfold the label or not. Unfolding partially hidden labels for the 'icon' view are always done."
ShowLabelTips = PropShowLabelTips
End Property

Public Property Let ShowLabelTips(ByVal Value As Boolean)
PropShowLabelTips = Value
If ListViewHandle <> 0 Then
    If PropShowLabelTips = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_LABELTIP, ByVal LVS_EX_LABELTIP
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_LABELTIP, ByVal 0&
    End If
End If
UserControl.PropertyChanged "ShowLabelTips"
End Property

Public Property Get DoubleBuffer() As Boolean
Attribute DoubleBuffer.VB_Description = "Returns/sets a value that determines whether the control paints via double-buffering, which reduces flicker. Requires comctl32.dll version 6.0 or higher."
DoubleBuffer = PropDoubleBuffer
End Property

Public Property Let DoubleBuffer(ByVal Value As Boolean)
PropDoubleBuffer = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If PropDoubleBuffer = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_DOUBLEBUFFER, ByVal LVS_EX_DOUBLEBUFFER
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_DOUBLEBUFFER, ByVal 0&
    End If
End If
UserControl.PropertyChanged "DoubleBuffer"
End Property

Public Property Get HoverSelection() As Boolean
Attribute HoverSelection.VB_Description = "Returns/sets a value that determines whether or not an list item is automatically selected when the cursor remains over the list item for a certain period of time."
HoverSelection = PropHoverSelection
End Property

Public Property Let HoverSelection(ByVal Value As Boolean)
If PropHotTracking = False Then PropHoverSelection = Value
If ListViewHandle <> 0 Then
    If PropHoverSelection = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT, ByVal LVS_EX_TRACKSELECT
    Else
        If PropHotTracking = False Then SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT, ByVal 0&
    End If
End If
UserControl.PropertyChanged "HoverSelection"
End Property

Public Property Get HoverSelectionTime() As Long
Attribute HoverSelectionTime.VB_Description = "Returns/sets the hover selection time in milliseconds. A value of -1 indicates that the default time is used."
If ListViewHandle <> 0 Then
    HoverSelectionTime = SendMessage(ListViewHandle, LVM_GETHOVERTIME, 0, ByVal 0&)
Else
    HoverSelectionTime = PropHoverSelectionTime
End If
End Property

Public Property Let HoverSelectionTime(ByVal Value As Long)
If Value < -1 Or Value = 0 Then
    If Ambient.UserMode = False Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropHoverSelectionTime = Value
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_SETHOVERTIME, 0, ByVal Value
UserControl.PropertyChanged "HoverSelectionTime"
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets whether hot tracking is enabled."
HotTracking = PropHotTracking
End Property

Public Property Let HotTracking(ByVal Value As Boolean)
If Value = True Then PropHoverSelection = True
PropHotTracking = Value
If ListViewHandle <> 0 Then
    If PropHotTracking = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE, ByVal LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE
    Else
        If PropHoverSelection = True Then
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_ONECLICKACTIVATE, ByVal 0&
        Else
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE, ByVal 0&
        End If
    End If
End If
UserControl.PropertyChanged "HotTracking"
End Property

Public Property Get HighlightHot() As Boolean
Attribute HighlightHot.VB_Description = "Returns/sets a value that determines whether hot items that may be activated to be displayed with highlighted text. Only applicable if the hot tracking property is set to true."
HighlightHot = PropHighlightHot
End Property

Public Property Let HighlightHot(ByVal Value As Boolean)
PropHighlightHot = Value
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_REDRAWITEMS, 0, ByVal SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&)
UserControl.PropertyChanged "HighlightHot"
End Property

Public Property Get UnderlineHot() As Boolean
Attribute UnderlineHot.VB_Description = "Returns/sets a value that determines whether hot items that may be activated to be displayed with underlined text. Only applicable if the hot tracking property is set to true."
UnderlineHot = PropUnderlineHot
End Property

Public Property Let UnderlineHot(ByVal Value As Boolean)
PropUnderlineHot = Value
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_REDRAWITEMS, 0, ByVal SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&)
UserControl.PropertyChanged "UnderlineHot"
End Property

Public Property Get InsertMarkColor() As OLE_COLOR
Attribute InsertMarkColor.VB_Description = "Returns/sets the color of the insertion mark. Requires comctl32.dll version 6.1 or higher."
InsertMarkColor = PropInsertMarkColor
End Property

Public Property Let InsertMarkColor(ByVal Value As OLE_COLOR)
PropInsertMarkColor = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then SendMessage ListViewHandle, LVM_SETINSERTMARKCOLOR, 0, ByVal WinColor(PropInsertMarkColor)
UserControl.PropertyChanged "InsertMarkColor"
End Property

Public Property Get TextBackground() As CCBackStyleConstants
Attribute TextBackground.VB_Description = "Returns/sets a value that determines if the text background is transparent or uses the background color of the list view."
TextBackground = PropTextBackground
End Property

Public Property Let TextBackground(ByVal Value As CCBackStyleConstants)
Select Case Value
    Case CCBackStyleTransparent, CCBackStyleOpaque
        PropTextBackground = Value
    Case Else
        Err.Raise 380
End Select
If ListViewHandle <> 0 Then
    If PropTextBackground = CCBackStyleTransparent Then
        SendMessage ListViewHandle, LVM_SETTEXTBKCOLOR, 0, ByVal CLR_NONE
    Else
        SendMessage ListViewHandle, LVM_SETTEXTBKCOLOR, 0, ByVal WinColor(PropBackColor)
    End If
End If
UserControl.PropertyChanged "TextBackground"
End Property

Public Property Get ClickableColumnHeaders() As Boolean
Attribute ClickableColumnHeaders.VB_Description = "Returns/sets a value that determines whether or not the column headers act like buttons and are clickable in 'report' view."
ClickableColumnHeaders = PropClickableColumnHeaders
End Property

Public Property Let ClickableColumnHeaders(ByVal Value As Boolean)
PropClickableColumnHeaders = Value
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropClickableColumnHeaders = CBool((dwStyle And HDS_BUTTONS) = HDS_BUTTONS) Then
            If PropClickableColumnHeaders = True Then
                If Not (dwStyle And HDS_BUTTONS) = HDS_BUTTONS Then dwStyle = dwStyle Or HDS_BUTTONS
            Else
                If (dwStyle And HDS_BUTTONS) = HDS_BUTTONS Then dwStyle = dwStyle And Not HDS_BUTTONS
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
        End If
    End If
End If
UserControl.PropertyChanged "ClickableColumnHeaders"
End Property

Public Property Get HighlightColumnHeaders() As Boolean
Attribute HighlightColumnHeaders.VB_Description = "Returns/sets a value that determines whether or not the control highlights the column headers as the pointer passes over them. This flag is ignored on Windows XP (or above) when the desktop theme overrides it."
HighlightColumnHeaders = PropHighlightColumnHeaders
End Property

Public Property Let HighlightColumnHeaders(ByVal Value As Boolean)
PropHighlightColumnHeaders = Value
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropHighlightColumnHeaders = CBool((dwStyle And HDS_HOTTRACK) = HDS_HOTTRACK) Then
            If PropHighlightColumnHeaders = True Then
                If Not (dwStyle And HDS_HOTTRACK) = HDS_HOTTRACK Then dwStyle = dwStyle Or HDS_HOTTRACK
            Else
                If (dwStyle And HDS_HOTTRACK) = HDS_HOTTRACK Then dwStyle = dwStyle And Not HDS_HOTTRACK
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
        End If
    End If
End If
UserControl.PropertyChanged "HighlightColumnHeaders"
End Property

Public Property Get TrackSizeColumnHeaders() As Boolean
Attribute TrackSizeColumnHeaders.VB_Description = "Returns/sets a value that determines whether or not the control display column header contents even while the user resizes them."
TrackSizeColumnHeaders = PropTrackSizeColumnHeaders
End Property

Public Property Let TrackSizeColumnHeaders(ByVal Value As Boolean)
PropTrackSizeColumnHeaders = Value
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropTrackSizeColumnHeaders = CBool((dwStyle And HDS_FULLDRAG) = HDS_FULLDRAG) Then
            If PropTrackSizeColumnHeaders = True Then
                If Not (dwStyle And HDS_FULLDRAG) = HDS_FULLDRAG Then dwStyle = dwStyle Or HDS_FULLDRAG
            Else
                If (dwStyle And HDS_FULLDRAG) = HDS_FULLDRAG Then dwStyle = dwStyle And Not HDS_FULLDRAG
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
        End If
    End If
End If
UserControl.PropertyChanged "TrackSizeColumnHeaders"
End Property

Public Property Get ResizableColumnHeaders() As Boolean
Attribute ResizableColumnHeaders.VB_Description = "Returns/sets a value that determines whether or not the user can drag the divider on the column header to resize them. Requires comctl32.dll version 6.1 or higher."
If Ambient.UserMode = True And ComCtlsSupportLevel() <= 1 Then
    ResizableColumnHeaders = True
Else
    ResizableColumnHeaders = PropResizableColumnHeaders
End If
End Property

Public Property Let ResizableColumnHeaders(ByVal Value As Boolean)
PropResizableColumnHeaders = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropResizableColumnHeaders = Not CBool((dwStyle And HDS_NOSIZING) = HDS_NOSIZING) Then
            If PropResizableColumnHeaders = True Then
                If (dwStyle And HDS_NOSIZING) = HDS_NOSIZING Then dwStyle = dwStyle And Not HDS_NOSIZING
            Else
                If Not (dwStyle And HDS_NOSIZING) = HDS_NOSIZING Then dwStyle = dwStyle Or HDS_NOSIZING
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
        End If
    End If
End If
UserControl.PropertyChanged "ResizableColumnHeaders"
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns/sets the background picture. Requires comctl32.dll version 6.0 or higher."
Set Picture = PropPicture
End Property

Public Property Let Picture(ByVal Value As IPictureDisp)
Set Me.Picture = Value
End Property

Public Property Set Picture(ByVal Value As IPictureDisp)
Dim LVBKI As LVBKIMAGE
With LVBKI
If Value Is Nothing Then
    Set PropPicture = Nothing
    If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
        .hBmp = 0
        .ulFlags = LVBKIF_SOURCE_NONE
        SendMessage ListViewHandle, LVM_SETBKIMAGE, 0, ByVal VarPtr(LVBKI)
        .ulFlags = LVBKIF_TYPE_WATERMARK
        SendMessage ListViewHandle, LVM_SETBKIMAGE, 0, ByVal VarPtr(LVBKI)
    End If
Else
    Set PropPicture = Value
    If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
        .hBmp = 0
        .ulFlags = LVBKIF_SOURCE_NONE
        SendMessage ListViewHandle, LVM_SETBKIMAGE, 0, ByVal VarPtr(LVBKI)
        .ulFlags = LVBKIF_TYPE_WATERMARK
        SendMessage ListViewHandle, LVM_SETBKIMAGE, 0, ByVal VarPtr(LVBKI)
        .ulFlags = LVBKIF_STYLE_NORMAL
        If Value.Handle <> 0 Then
            .hBmp = BitmapHandleFromPicture(PropPicture, PropBackColor)
            If PropPictureWatermark = False Then
                ' There is a much better result without LVS_EX_DOUBLEBUFFER
                ' when loading picture by 'hBmp'. (Weighing the pros and cons)
                If PropDoubleBuffer = True Then Me.DoubleBuffer = False
                Select Case PropPictureAlignment
                    Case LvwPictureAlignmentTopLeft
                        .XOffsetPercent = 0
                        .YOffsetPercent = 0
                    Case LvwPictureAlignmentTopRight
                        .XOffsetPercent = 100
                        .YOffsetPercent = 0
                    Case LvwPictureAlignmentBottomLeft
                        .XOffsetPercent = 0
                        .YOffsetPercent = 100
                    Case LvwPictureAlignmentBottomRight
                        .XOffsetPercent = 100
                        .YOffsetPercent = 100
                    Case LvwPictureAlignmentCenter
                        .XOffsetPercent = 50
                        .YOffsetPercent = 50
                    Case LvwPictureAlignmentTile
                        ' There is a better result when no column is selected.
                        Set Me.SelectedColumn = Nothing
                        .ulFlags = .ulFlags Or LVBKIF_STYLE_TILE
                        If ComCtlsSupportLevel() >= 2 And PropView = LvwViewReport Then
                            Dim HeaderHandle As Long
                            If HeaderHandle = 0 Then HeaderHandle = Me.hWndHeader
                            If HeaderHandle <> 0 Then
                                .ulFlags = .ulFlags Or LVBKIF_FLAG_TILEOFFSET
                                Dim RC As RECT
                                GetWindowRect ListViewHeaderHandle, RC
                                .YOffsetPercent = -(RC.Bottom - RC.Top)
                            End If
                        End If
                End Select
                .ulFlags = .ulFlags Or LVBKIF_SOURCE_HBITMAP
            Else
                ' Here it does not matter whether LVS_EX_DOUBLEBUFFER is set or not.
                ' Though it is better to set it as it reduces flicker, especially
                ' when a watermark is in place.
                If PropDoubleBuffer = False Then Me.DoubleBuffer = True
                .ulFlags = .ulFlags Or LVBKIF_TYPE_WATERMARK
            End If
            SendMessage ListViewHandle, LVM_SETBKIMAGE, 0, ByVal VarPtr(LVBKI)
        End If
    End If
End If
End With
UserControl.PropertyChanged "Picture"
End Property

Public Property Get PictureAlignment() As LvwPictureAlignmentConstants
Attribute PictureAlignment.VB_Description = "Returns/sets the picture alignment. Requires comctl32.dll version 6.0 or higher."
PictureAlignment = PropPictureAlignment
End Property

Public Property Let PictureAlignment(ByVal Value As LvwPictureAlignmentConstants)
Select Case Value
    Case LvwPictureAlignmentTopLeft, LvwPictureAlignmentTopRight, LvwPictureAlignmentBottomLeft, LvwPictureAlignmentBottomRight, LvwPictureAlignmentCenter, LvwPictureAlignmentTile
        PropPictureAlignment = Value
    Case Else
        Err.Raise 380
End Select
Set Me.Picture = PropPicture
UserControl.PropertyChanged "PictureAlignment"
End Property

Public Property Get PictureWatermark() As Boolean
Attribute PictureWatermark.VB_Description = "Returns/sets a value that determines whether a watermark background bitmap is supplied in the picture property. That means the picture will always be displayed in the lower right corner. Requires comctl32.dll version 6.0 or higher."
PictureWatermark = PropPictureWatermark
End Property

Public Property Let PictureWatermark(ByVal Value As Boolean)
PropPictureWatermark = Value
Set Me.Picture = PropPicture
UserControl.PropertyChanged "PictureWatermark"
End Property

Public Property Get TileViewLines() As Long
Attribute TileViewLines.VB_Description = "Returns/sets the maximum number of text lines (not counting the title) in each list item in 'tile' view. Requires comctl32.dll version 6.0 or higher."
TileViewLines = PropTileViewLines
End Property

Public Property Let TileViewLines(ByVal Value As Long)
Select Case Value
    Case 0 To 20
        PropTileViewLines = Value
    Case Else
        If Ambient.UserMode = False Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
End Select
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim LVTVI As LVTILEVIEWINFO
    With LVTVI
    .cbSize = LenB(LVTVI)
    .dwMask = LVTVIM_COLUMNS
    SendMessage ListViewHandle, LVM_GETTILEVIEWINFO, 0, ByVal VarPtr(LVTVI)
    .cLines = Value
    SendMessage ListViewHandle, LVM_SETTILEVIEWINFO, 0, ByVal VarPtr(LVTVI)
    End With
End If
UserControl.PropertyChanged "TileViewLines"
End Property

Public Property Get SnapToGrid() As Boolean
Attribute SnapToGrid.VB_Description = "Returns/sets a value that determines whether or not the list items automatically snaps into a grid in 'icon', 'small icon' or 'tile' view. Requires comctl32.dll version 6.0 or higher."
SnapToGrid = PropSnapToGrid
End Property

Public Property Let SnapToGrid(ByVal Value As Boolean)
PropSnapToGrid = Value
If ListViewHandle <> 0 Then
    If PropSnapToGrid = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SNAPTOGRID, ByVal LVS_EX_SNAPTOGRID
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SNAPTOGRID, ByVal 0&
    End If
End If
UserControl.PropertyChanged "SnapToGrid"
End Property

Public Property Get ListItems() As LvwListItems
Attribute ListItems.VB_Description = "Returns a reference to a collection of the list item objects."
If PropListItems Is Nothing Then
    Set PropListItems = New LvwListItems
    PropListItems.FInit Me
End If
Set ListItems = PropListItems
End Property

Friend Sub FListItemsAdd(ByVal Ptr As Long, ByVal Index As Long, Optional ByVal Text As String)
Dim LVI As LVITEM
With LVI
.Mask = LVIF_TEXT Or LVIF_IMAGE Or LVIF_PARAM Or LVIF_INDENT
.iItem = Index - 1
.pszText = StrPtr(Text)
.cchTextMax = Len(Text) + 1
.iImage = I_IMAGECALLBACK
.lParam = Ptr
.iIndent = 0
End With
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_INSERTITEM, 0, ByVal VarPtr(LVI)
If PropSorted = True Then If PropSortKey = 0 Then Call SortListItems
End Sub

Friend Sub FListItemsRemove(ByVal Index As Long)
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, LVM_DELETEITEM, Index - 1, ByVal 0&
    If (Me.ListItems.Count - 1) = 0 Then
        Call CheckItemFocus(0)
    ElseIf ListViewFocusIndex > Index Then
        ListViewFocusIndex = ListViewFocusIndex - 1
    End If
End If
End Sub

Friend Sub FListItemsClear()
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_DELETEALLITEMS, 0, ByVal 0&
Set PropListItems = Nothing
Call CheckItemFocus(0)
End Sub

Friend Function FListItemsControl() As Long
FListItemsControl = ListViewListItemsControl
End Function

Friend Sub FListItemsNextItem(ByRef Index As Long, ByRef Control As Long, ByRef Data As Long, ByRef VNextItem As Variant, ByRef NoMoreItems As Boolean)
If Control <> ListViewListItemsControl Then Err.Raise Number:=1, Description:="Collection has changed during enumeration"
Index = Index + 1
NoMoreItems = CBool(Index < 0 Or Index >= Me.ListItems.Count)
If NoMoreItems = False Then Set VNextItem = Me.ListItems(Index + 1)
End Sub

Friend Sub FListSubItemsNextItem(ByVal ListItem As LvwListItem, ByRef Index As Long, ByRef Control As Long, ByRef Data As Long, ByRef VNextItem As Variant, ByRef NoMoreItems As Boolean)
If Control <> ListViewListItemsControl Then Err.Raise Number:=1, Description:="Collection has changed during enumeration"
Index = Index + 1
NoMoreItems = CBool(Index < 0 Or Index >= ListItem.ListSubItems.Count)
If NoMoreItems = False Then Set VNextItem = ListItem.ListSubItems(Index + 1)
End Sub

Friend Function FListItemPtr(ByVal Index As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    .Mask = LVIF_PARAM
    .iItem = Index - 1
    SendMessage ListViewHandle, LVM_GETITEM, 0, ByVal VarPtr(LVI)
    FListItemPtr = .lParam
    End With
End If
End Function

Friend Function FListItemVerify(ByVal Ptr As Long, ByRef Index As Long) As Boolean
If Ptr = Me.FListItemPtr(Index) Then
    FListItemVerify = True
Else
    Index = Me.FListItemIndex(Ptr)
    FListItemVerify = CBool(Index <> 0)
End If
End Function

Friend Function FListItemIndex(ByVal Ptr As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVFI As LVFINDINFO
    With LVFI
    .Flags = LVFI_PARAM
    .lParam = Ptr
    End With
    FListItemIndex = SendMessage(ListViewHandle, LVM_FINDITEM, -1, ByVal VarPtr(LVFI)) + 1
End If
End Function

Friend Sub FListItemRedraw(ByVal Index As Long)
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, LVM_REDRAWITEMS, Index - 1, ByVal (Index - 1)
    UpdateWindow ListViewHandle
End If
End Sub

Friend Property Get FListItemText(ByVal Index As Long, ByVal SubItemIndex As Long) As String
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    Dim Buffer As String
    Buffer = String(260, vbNullChar)
    .pszText = StrPtr(Buffer)
    .cchTextMax = 260
    .iSubItem = SubItemIndex
    End With
    SendMessage ListViewHandle, LVM_GETITEMTEXT, Index - 1, ByVal VarPtr(LVI)
    FListItemText = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End If
End Property

Friend Property Let FListItemText(ByVal Index As Long, ByVal SubItemIndex As Long, ByVal Value As String)
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    .pszText = StrPtr(Value)
    .cchTextMax = Len(Value) + 1
    .iSubItem = SubItemIndex
    End With
    SendMessage ListViewHandle, LVM_SETITEMTEXT, Index - 1, ByVal VarPtr(LVI)
End If
End Property

Friend Property Get FListItemIndentation(ByVal Index As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    LVI.Mask = LVIF_INDENT
    SendMessage ListViewHandle, LVM_GETITEM, 0, ByVal VarPtr(LVI)
    FListItemIndentation = LVI.iIndent
End If
End Property

Friend Property Let FListItemIndentation(ByVal Index As Long, ByVal Value As Long)
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    LVI.Mask = LVIF_INDENT
    LVI.iIndent = Value
    SendMessage ListViewHandle, LVM_SETITEM, 0, ByVal VarPtr(LVI)
End If
End Property

Friend Property Get FListItemSelected(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 Then FListItemSelected = CBool((SendMessage(ListViewHandle, LVM_GETITEMSTATE, Index - 1, ByVal LVIS_SELECTED) And LVIS_SELECTED) = LVIS_SELECTED)
End Property

Friend Property Let FListItemSelected(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    If Value = True Then
        .StateMask = LVIS_SELECTED Or LVIS_FOCUSED
        .State = LVIS_SELECTED Or LVIS_FOCUSED
    Else
        .StateMask = LVIS_SELECTED
        .State = Not LVIS_SELECTED
    End If
    End With
    SendMessage ListViewHandle, LVM_SETITEMSTATE, Index - 1, ByVal VarPtr(LVI)
End If
End Property

Friend Property Get FListItemChecked(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 Then FListItemChecked = CBool((SendMessage(ListViewHandle, LVM_GETITEMSTATE, Index - 1, ByVal LVIS_STATEIMAGEMASK) And &H2000) = &H2000)
End Property

Friend Property Let FListItemChecked(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    .StateMask = &H3000
    If Value = True Then
        .State = &H2000
    Else
        .State = &H1000
    End If
    End With
    SendMessage ListViewHandle, LVM_SETITEMSTATE, Index - 1, ByVal VarPtr(LVI)
End If
End Property

Friend Property Get FListItemGhosted(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 Then FListItemGhosted = CBool((SendMessage(ListViewHandle, LVM_GETITEMSTATE, Index - 1, ByVal LVIS_CUT) And LVIS_CUT) = LVIS_CUT)
End Property

Friend Property Let FListItemGhosted(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    .StateMask = LVIS_CUT
    If Value = True Then
        .State = LVIS_CUT
    Else
        .State = Not LVIS_CUT
    End If
    End With
    SendMessage ListViewHandle, LVM_SETITEMSTATE, Index - 1, ByVal VarPtr(LVI)
End If
End Property

Friend Property Get FListItemHot(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETHOTITEM, 0, ByVal 0&)
    If iItem > -1 Then FListItemHot = CBool(Index = (iItem + 1))
End If
End Property

Friend Property Let FListItemHot(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 Then
    If Value = True Then
        SendMessage ListViewHandle, LVM_SETHOTITEM, Index - 1, ByVal 0&
    Else
        SendMessage ListViewHandle, LVM_SETHOTITEM, -1, ByVal 0&
    End If
End If
End Property

Friend Property Get FListItemLeft(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = LVIR_SELECTBOUNDS
    SendMessage ListViewHandle, LVM_GETITEMRECT, Index - 1, ByVal VarPtr(RC)
    FListItemLeft = UserControl.ScaleX(RC.Left, vbPixels, vbContainerPosition)
End If
End Property

Friend Property Get FListItemTop(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = LVIR_SELECTBOUNDS
    SendMessage ListViewHandle, LVM_GETITEMRECT, Index - 1, ByVal VarPtr(RC)
    FListItemTop = UserControl.ScaleY(RC.Top, vbPixels, vbContainerPosition)
End If
End Property

Friend Property Get FListItemWidth(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = LVIR_SELECTBOUNDS
    SendMessage ListViewHandle, LVM_GETITEMRECT, Index - 1, ByVal VarPtr(RC)
    FListItemWidth = UserControl.ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize)
End If
End Property

Friend Property Get FListItemHeight(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = LVIR_SELECTBOUNDS
    SendMessage ListViewHandle, LVM_GETITEMRECT, Index - 1, ByVal VarPtr(RC)
    FListItemHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
End If
End Property

Friend Sub FListItemEnsureVisible(ByVal Index As Long)
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_ENSUREVISIBLE, Index - 1, ByVal 0&
End Sub

Friend Property Get FListItemTileViewIndices(ByVal Index As Long) As Variant
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim Buffer(0 To 19) As Long
    Dim LVTI As LVTILEINFO
    With LVTI
    .cbSize = LenB(LVTI)
    .iItem = Index - 1
    .cColumns = 20
    .puColumns = VarPtr(Buffer(0))
    SendMessage ListViewHandle, LVM_GETTILEINFO, 0, ByVal VarPtr(LVTI)
    If .cColumns > 0 Then
        Dim ArgList() As Long, i As Long
        ReDim ArgList(0 To (.cColumns - 1)) As Long
        For i = 0 To (.cColumns - 1)
            ArgList(i) = Buffer(i)
        Next i
        FListItemTileViewIndices = ArgList()
    Else
        FListItemTileViewIndices = Empty
    End If
    End With
End If
End Property

Friend Property Let FListItemTileViewIndices(ByVal Index As Long, ByVal ArgList As Variant)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim LVTI As LVTILEINFO
    With LVTI
    .cbSize = LenB(LVTI)
    .iItem = Index - 1
    If IsArray(ArgList) Then
        Dim Ptr As Long
        CopyMemory Ptr, ByVal UnsignedAdd(VarPtr(ArgList), 8), 4
        If Ptr <> 0 Then
            Dim RetVal As Long
            CopyMemory ByVal VarPtr(RetVal), Ptr, 4
            If RetVal <> 0 Then
                Dim DimensionCount As Integer
                CopyMemory DimensionCount, ByVal Ptr, 2
                If DimensionCount = 1 Then
                    Dim Arr() As Long, Count As Long, i As Long
                    For i = LBound(ArgList) To UBound(ArgList)
                        Select Case VarType(ArgList(i))
                            Case vbLong, vbInteger, vbByte
                                If ArgList(i) > 0 Then
                                    ReDim Preserve Arr(0 To Count) As Long
                                    Arr(Count) = ArgList(i)
                                    Count = Count + 1
                                End If
                        End Select
                    Next i
                    If Count > 0 Then
                        .cColumns = Count
                        .puColumns = VarPtr(Arr(0))
                    Else
                        .cColumns = 0
                        .puColumns = 0
                    End If
                Else
                    Err.Raise Number:=5, Description:="Array must be single dimensioned"
                End If
            Else
                Err.Raise Number:=91, Description:="Array is not allocated"
            End If
        Else
            Err.Raise 5
        End If
    ElseIf IsEmpty(ArgList) Then
        .cColumns = 0
        .puColumns = 0
    Else
        Err.Raise 380
    End If
    SendMessage ListViewHandle, LVM_SETTILEINFO, 0, ByVal VarPtr(LVTI)
    End With
End If
End Property

Friend Function FListItemCreateDragImage(ByVal Index As Long, ByRef X As Single, ByRef Y As Single) As Long
If ListViewHandle <> 0 Then
    Dim P As POINTAPI
    FListItemCreateDragImage = SendMessage(ListViewHandle, LVM_CREATEDRAGIMAGE, Index - 1, ByVal VarPtr(P))
    X = UserControl.ScaleX(P.X, vbPixels, vbContainerPosition)
    Y = UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition)
End If
End Function

Friend Function FListSubItemAlloc(ByVal Key As String) As Long
FListSubItemAlloc = HeapAlloc(GetProcessHeap(), -HEAP_ZERO_MEMORY, 26)
If FListSubItemAlloc <> 0 Then
    CopyMemory ByVal UnsignedAdd(FListSubItemAlloc, 4), SysAllocString(StrPtr(Key)), 4
    CopyMemory ByVal UnsignedAdd(FListSubItemAlloc, 22), -1&, 4
End If
End Function

Friend Sub FListSubItemFree(ByVal SubPtr As Long)
Dim lpText As Long, lpKey As Long, lpTag As Long
CopyMemory lpText, ByVal SubPtr, 4
CopyMemory lpKey, ByVal UnsignedAdd(SubPtr, 4), 4
CopyMemory lpTag, ByVal UnsignedAdd(SubPtr, 8), 4
If lpText <> 0 Then SysFreeString lpText
If lpKey <> 0 Then SysFreeString lpKey
If lpTag <> 0 Then SysFreeString lpTag
HeapFree GetProcessHeap(), 0, SubPtr
End Sub

Friend Property Get FListSubItemText(ByRef SubPtr As Long) As String
Dim lpText As Long
CopyMemory lpText, ByVal SubPtr, 4
FListSubItemText = vbNullString
If lpText <> 0 Then
    Dim lpString As Long
    lpString = SysAllocString(lpText)
    If lpString = 0 Then Err.Raise 7
    CopyMemory ByVal VarPtr(FListSubItemText), lpString, 4
End If
End Property

Friend Property Let FListSubItemText(ByRef SubPtr As Long, ByVal Value As String)
Dim lpText As Long
CopyMemory lpText, ByVal SubPtr, 4
If lpText <> 0 Then
    SysFreeString lpText
    lpText = 0
End If
If LenB(Value) <> 0 Then lpText = SysAllocString(StrPtr(Value))
CopyMemory ByVal SubPtr, lpText, 4
End Property

Friend Property Get FListSubItemKey(ByVal SubPtr As Long) As String
Dim lpKey As Long
CopyMemory lpKey, ByVal UnsignedAdd(SubPtr, 4), 4
FListSubItemKey = vbNullString
If lpKey <> 0 Then
    Dim lpString As Long
    lpString = SysAllocString(lpKey)
    If lpString = 0 Then Err.Raise 7
    CopyMemory ByVal VarPtr(FListSubItemKey), lpString, 4
End If
End Property

Friend Property Get FListSubItemTag(ByRef SubPtr As Long) As String
Dim lpTag As Long
CopyMemory lpTag, ByVal UnsignedAdd(SubPtr, 8), 4
FListSubItemTag = vbNullString
If lpTag <> 0 Then
    Dim lpString As Long
    lpString = SysAllocString(lpTag)
    If lpString = 0 Then Err.Raise 7
    CopyMemory ByVal VarPtr(FListSubItemTag), lpString, 4
End If
End Property

Friend Property Let FListSubItemTag(ByRef SubPtr As Long, ByVal Value As String)
Dim lpTag As Long
CopyMemory lpTag, ByVal UnsignedAdd(SubPtr, 8), 4
If lpTag <> 0 Then
    SysFreeString lpTag
    lpTag = 0
End If
If LenB(Value) <> 0 Then lpTag = SysAllocString(StrPtr(Value))
CopyMemory ByVal UnsignedAdd(SubPtr, 8), lpTag, 4
End Property

Friend Property Get FListSubItemSubIndex(ByRef SubPtr As Long) As Long
CopyMemory FListSubItemSubIndex, ByVal UnsignedAdd(SubPtr, 12), 4
End Property

Friend Property Let FListSubItemSubIndex(ByRef SubPtr As Long, ByVal Value As Long)
CopyMemory ByVal UnsignedAdd(SubPtr, 12), Value, 4
End Property

Friend Property Get FListSubItemReportIcon(ByRef SubPtr As Long) As Long
CopyMemory FListSubItemReportIcon, ByVal UnsignedAdd(SubPtr, 16), 4
End Property

Friend Property Let FListSubItemReportIcon(ByRef SubPtr As Long, ByVal Value As Long)
Static Once As Boolean
If Value > 0 And Once = False Then
    If ListViewHandle <> 0 Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SUBITEMIMAGES, ByVal LVS_EX_SUBITEMIMAGES
        Once = True
    End If
End If
CopyMemory ByVal UnsignedAdd(SubPtr, 16), Value, 4
End Property

Friend Property Get FListSubItemBold(ByRef SubPtr As Long) As Boolean
CopyMemory FListSubItemBold, ByVal UnsignedAdd(SubPtr, 20), 2
End Property

Friend Property Let FListSubItemBold(ByRef SubPtr As Long, ByVal Value As Boolean)
CopyMemory ByVal UnsignedAdd(SubPtr, 20), Value, 2
End Property

Friend Property Get FListSubItemForeColor(ByRef SubPtr As Long) As OLE_COLOR
CopyMemory FListSubItemForeColor, ByVal UnsignedAdd(SubPtr, 22), 4
End Property

Friend Property Let FListSubItemForeColor(ByRef SubPtr As Long, ByVal Value As OLE_COLOR)
CopyMemory ByVal UnsignedAdd(SubPtr, 22), Value, 4
End Property

Public Property Get ColumnHeaders() As LvwColumnHeaders
Attribute ColumnHeaders.VB_Description = "Returns a reference to a collection of the column objects."
If PropColumnHeaders Is Nothing Then
    Set PropColumnHeaders = New LvwColumnHeaders
    PropColumnHeaders.FInit Me
End If
Set ColumnHeaders = PropColumnHeaders
End Property

Friend Sub FColumnHeadersAdd(Optional ByVal Index As Long, Optional ByVal Text As String, Optional ByVal Width As Single, Optional ByVal Alignment As LvwColumnHeaderAlignmentConstants, Optional ByVal Icon As Long)
Dim ColumnHeaderIndex As Long
If Index = 0 Then
    ColumnHeaderIndex = Me.ColumnHeaders.Count + 1
Else
    ColumnHeaderIndex = Index
End If
Dim LVC As LVCOLUMN
With LVC
.Mask = LVCF_FMT Or LVCF_WIDTH
If Not Text = vbNullString Then
    .Mask = .Mask Or LVCF_TEXT
    .pszText = StrPtr(Text)
    .cchTextMax = Len(Text) + 1
End If
If Width = 0 Then
    .CX = 96
ElseIf Width > 0 Then
    .CX = UserControl.ScaleX(Width, vbContainerSize, vbPixels)
Else
    Err.Raise 380
End If
If (ColumnHeaderIndex - 1) = 0 Then
    .fmt = LVCFMT_LEFT
Else
    Select Case Alignment
        Case LvwColumnHeaderAlignmentLeft
            .fmt = LVCFMT_LEFT
        Case LvwColumnHeaderAlignmentRight
            .fmt = LVCFMT_RIGHT
        Case LvwColumnHeaderAlignmentCenter
            .fmt = LVCFMT_CENTER
        Case Else
            Err.Raise 380
    End Select
End If
If Icon > 0 Then
    .fmt = .fmt Or LVCFMT_IMAGE
    .Mask = .Mask Or LVCF_IMAGE
    .iImage = Icon - 1
End If
End With
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
    SendMessage ListViewHandle, LVM_INSERTCOLUMN, ColumnHeaderIndex - 1, ByVal VarPtr(LVC)
    If (ColumnHeaderIndex - 1) = 0 Then
        ' According to MSDN:
        ' If a column is added to a list view control with index 0 (the leftmost column), it is always LVCFMT_LEFT.
        ' Workaround: Adjust the fmt value after the insert.
        If Alignment <> LvwColumnHeaderAlignmentLeft Then Me.FColumnHeaderAlignment(1) = Alignment
    End If
    Call SetColumnsSubItemIndex(1)
    Call RebuildListItems
    If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
End If
End Sub

Friend Sub FColumnHeadersRemove(ByVal Index As Long)
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
    SendMessage ListViewHandle, LVM_DELETECOLUMN, Index - 1, ByVal 0&
    Call SetColumnsSubItemIndex(-1)
    Call RebuildListItems
    If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
End If
End Sub

Friend Sub FColumnHeadersClear()
If ListViewHandle <> 0 Then Do While SendMessage(ListViewHandle, LVM_DELETECOLUMN, 0, ByVal 0&) = 1: Loop
Set PropColumnHeaders = Nothing
End Sub

Friend Property Get FColumnHeaderText(ByVal Index As Long) As String
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_TEXT
    Dim Buffer As String
    Buffer = String(260, vbNullChar)
    .pszText = StrPtr(Buffer)
    .cchTextMax = 260
    End With
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderText = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End If
End Property

Friend Property Let FColumnHeaderText(ByVal Index As Long, ByVal Value As String)
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_TEXT
    .pszText = StrPtr(Value)
    End With
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
End If
End Property

Friend Property Get FColumnHeaderIcon(ByVal Index As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    LVC.Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If (LVC.fmt And LVCFMT_IMAGE) = LVCFMT_IMAGE Then
        LVC.Mask = LVCF_IMAGE
        SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
        FColumnHeaderIcon = LVC.iImage + 1
    End If
End If
End Property

Friend Property Let FColumnHeaderIcon(ByVal Index As Long, ByVal Value As Long)
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    .Mask = .Mask Or LVCF_IMAGE
    .iImage = Value - 1
    If Value > 0 Then
        .fmt = .fmt Or LVCFMT_IMAGE
    Else
        .fmt = .fmt And Not LVCFMT_IMAGE
    End If
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderWidth(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    LVC.Mask = LVCF_WIDTH
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderWidth = UserControl.ScaleX(LVC.CX, vbPixels, vbContainerSize)
End If
End Property

Friend Property Let FColumnHeaderWidth(ByVal Index As Long, ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_WIDTH
    .CX = UserControl.ScaleX(Value, vbContainerSize, vbPixels)
    End With
    If PropView = LvwViewReport Then
        SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    Else
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
        SendMessage ListViewHandle, LVM_UPDATE, 0, ByVal 0&
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
    End If
End If
End Property

Friend Property Get FColumnHeaderAlignment(ByVal Index As Long) As LvwColumnHeaderAlignmentConstants
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    Select Case LVC.fmt And LVCFMT_JUSTIFYMASK
        Case LVCFMT_LEFT
            FColumnHeaderAlignment = LvwColumnHeaderAlignmentLeft
        Case LVCFMT_RIGHT
            FColumnHeaderAlignment = LvwColumnHeaderAlignmentRight
        Case LVCFMT_CENTER
            FColumnHeaderAlignment = LvwColumnHeaderAlignmentCenter
    End Select
    End With
End If
End Property

Friend Property Let FColumnHeaderAlignment(ByVal Index As Long, ByVal Value As LvwColumnHeaderAlignmentConstants)
If ListViewHandle <> 0 Then
    Select Case Value
        Case LvwColumnHeaderAlignmentLeft, LvwColumnHeaderAlignmentRight, LvwColumnHeaderAlignmentCenter
            Dim LVC As LVCOLUMN
            With LVC
            .Mask = LVCF_FMT
            SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
            .fmt = .fmt And Not LVCFMT_JUSTIFYMASK
            Select Case Value
                Case LvwColumnHeaderAlignmentLeft
                    .fmt = .fmt Or LVCFMT_LEFT
                Case LvwColumnHeaderAlignmentRight
                    .fmt = .fmt Or LVCFMT_RIGHT
                Case LvwColumnHeaderAlignmentCenter
                    .fmt = .fmt Or LVCFMT_CENTER
            End Select
            End With
            SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
        Case Else
            Err.Raise 380
    End Select
End If
End Property

Friend Property Get FColumnHeaderPosition(ByVal Index As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    LVC.Mask = LVCF_ORDER
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderPosition = LVC.iOrder + 1
End If
End Property

Friend Property Let FColumnHeaderPosition(ByVal Index As Long, ByVal Value As Long)
If ListViewHandle <> 0 Then
    If Value < 1 Or Value > Me.ColumnHeaders.Count Then
        Err.Raise 380
    Else
        Dim LVC As LVCOLUMN
        With LVC
        .Mask = LVCF_ORDER
        .iOrder = Value - 1
        End With
        SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
        Me.Refresh
    End If
End If
End Property

Friend Property Get FColumnHeaderSortArrow(ByVal Index As Long) As LvwColumnHeaderSortArrowConstants
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If (.fmt And HDF_SORTUP) = HDF_SORTUP Then
        FColumnHeaderSortArrow = LvwColumnHeaderSortArrowUp
    ElseIf (.fmt And HDF_SORTDOWN) = HDF_SORTDOWN Then
        FColumnHeaderSortArrow = LvwColumnHeaderSortArrowDown
    Else
        FColumnHeaderSortArrow = LvwColumnHeaderSortArrowNone
    End If
    End With
End If
End Property

Friend Property Let FColumnHeaderSortArrow(ByVal Index As Long, ByVal Value As LvwColumnHeaderSortArrowConstants)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    Select Case Value
        Case LvwColumnHeaderSortArrowNone
            .fmt = .fmt And Not (HDF_SORTDOWN Or HDF_SORTUP)
        Case LvwColumnHeaderSortArrowDown
            .fmt = .fmt Or HDF_SORTDOWN
            .fmt = .fmt And Not HDF_SORTUP
        Case LvwColumnHeaderSortArrowUp
            .fmt = .fmt Or HDF_SORTUP
            .fmt = .fmt And Not HDF_SORTDOWN
    End Select
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderIconOnRight(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderIconOnRight = CBool((.fmt And HDF_BITMAP_ON_RIGHT) = HDF_BITMAP_ON_RIGHT)
    End With
End If
End Property

Friend Property Let FColumnHeaderIconOnRight(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If Value = True Then
        .fmt = .fmt Or HDF_BITMAP_ON_RIGHT
    Else
        .fmt = .fmt And Not HDF_BITMAP_ON_RIGHT
    End If
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderResizable(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderResizable = Not CBool((.fmt And HDF_FIXEDWIDTH) = HDF_FIXEDWIDTH)
    End With
Else
    FColumnHeaderResizable = True
End If
End Property

Friend Property Let FColumnHeaderResizable(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If Value = True Then
        .fmt = .fmt And Not HDF_FIXEDWIDTH
    Else
        .fmt = .fmt Or HDF_FIXEDWIDTH
    End If
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderSplitButton(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderSplitButton = CBool((.fmt And HDF_SPLITBUTTON) = HDF_SPLITBUTTON)
    End With
End If
End Property

Friend Property Let FColumnHeaderSplitButton(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If Value = True Then
        .fmt = .fmt Or HDF_SPLITBUTTON
    Else
        .fmt = .fmt And Not HDF_SPLITBUTTON
    End If
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderCheckBox(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderCheckBox = CBool((.fmt And HDF_CHECKBOX) = HDF_CHECKBOX)
    End With
End If
End Property

Friend Property Let FColumnHeaderCheckBox(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If Value = True Then
        .fmt = .fmt Or HDF_CHECKBOX
    Else
        .fmt = .fmt And Not HDF_CHECKBOX
    End If
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderChecked(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderChecked = CBool((.fmt And HDF_CHECKED) = HDF_CHECKED)
    End With
End If
End Property

Friend Property Let FColumnHeaderChecked(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If Value = True Then
        .fmt = .fmt Or HDF_CHECKED
    Else
        .fmt = .fmt And Not HDF_CHECKED
    End If
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderLeft(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim i As Long
    For i = 1 To Index
        If i = Index Then Exit For
        FColumnHeaderLeft = FColumnHeaderLeft + Me.FColumnHeaderWidth(i)
    Next i
End If
End Property

Friend Sub FColumnHeaderAutoSize(ByVal Index As Long, ByVal Value As LvwColumnHeaderAutoSizeConstants)
If ListViewHandle <> 0 Then
    Dim Flag As Long
    Select Case Value
        Case LvwColumnHeaderAutoSizeToItems
            Flag = LVSCW_AUTOSIZE
        Case LvwColumnHeaderAutoSizeToHeader
            Flag = LVSCW_AUTOSIZE_USEHEADER
        Case Else
            Err.Raise 380
    End Select
    SendMessage ListViewHandle, LVM_SETCOLUMNWIDTH, Index - 1, ByVal Flag
End If
End Sub

Friend Function FColumnHeaderSubItemIndex(ByVal Index As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    LVC.Mask = LVCF_SUBITEM
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderSubItemIndex = LVC.iSubItem
End If
End Function

Private Sub CreateListView()
If ListViewHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or LVS_SHAREIMAGELISTS
Select Case PropBorderStyle
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
If Ambient.UserMode = True Then
    If ComCtlsSupportLevel() = 0 And PropView = LvwViewTile Then PropView = LvwViewIcon
    Select Case PropView
        Case LvwViewIcon
            dwStyle = dwStyle Or LVS_ICON
        Case LvwViewSmallIcon
            dwStyle = dwStyle Or LVS_SMALLICON
        Case LvwViewList
            dwStyle = dwStyle Or LVS_LIST
        Case LvwViewReport
            dwStyle = dwStyle Or LVS_REPORT
    End Select
Else
    dwStyle = dwStyle Or LVS_LIST
End If
Select Case PropArrange
    Case LvwArrangeAutoLeft
        dwStyle = dwStyle Or LVS_AUTOARRANGE Or LVS_ALIGNLEFT
    Case LvwArrangeAutoTop
        dwStyle = dwStyle Or LVS_AUTOARRANGE Or LVS_ALIGNTOP
    Case LvwArrangeLeft
        dwStyle = dwStyle Or LVS_ALIGNLEFT
    Case LvwArrangeTop
        dwStyle = dwStyle Or LVS_ALIGNTOP
End Select
If PropMultiSelect = False Then dwStyle = dwStyle Or LVS_SINGLESEL
If PropLabelEdit <> LvwLabelEditDisabled Then dwStyle = dwStyle Or LVS_EDITLABELS
If PropLabelWrap = False Then dwStyle = dwStyle Or LVS_NOLABELWRAP
If PropHideSelection = False Then dwStyle = dwStyle Or LVS_SHOWSELALWAYS
If PropHideColumnHeaders = True Then dwStyle = dwStyle Or LVS_NOCOLUMNHEADER
If Ambient.RightToLeft = True Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
ListViewHandle = CreateWindowEx(dwExStyle, StrPtr("SysListView32"), StrPtr("List View"), dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If PropView = LvwViewReport Then ListViewHeaderHandle = Me.hWndHeader
If ListViewHandle <> 0 Then If PropView = LvwViewTile Then SendMessage ListViewHandle, LVM_SETVIEW, LV_VIEW_TILE, ByVal 0&
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.BackColor = PropBackColor
Me.ForeColor = PropForeColor
If PropRedraw = False Then Me.Redraw = False
Me.AllowColumnReorder = PropAllowColumnReorder
Me.AllowColumnCheckboxes = PropAllowColumnCheckboxes
Me.FullRowSelect = PropFullRowSelect
Me.GridLines = PropGridLines
Me.Checkboxes = PropCheckboxes
Me.ShowInfoTips = PropShowInfoTips
Me.ShowLabelTips = PropShowLabelTips
Me.DoubleBuffer = PropDoubleBuffer
Me.HoverSelection = PropHoverSelection
Me.HoverSelectionTime = PropHoverSelectionTime
Me.HotTracking = PropHotTracking
Me.InsertMarkColor = PropInsertMarkColor
Me.TextBackground = PropTextBackground
Me.ClickableColumnHeaders = PropClickableColumnHeaders
Me.HighlightColumnHeaders = PropHighlightColumnHeaders
Me.TrackSizeColumnHeaders = PropTrackSizeColumnHeaders
Me.ResizableColumnHeaders = PropResizableColumnHeaders
If Not PropPicture Is Nothing Then Set Me.Picture = PropPicture
Me.TileViewLines = PropTileViewLines
Me.SnapToGrid = PropSnapToGrid
If ListViewHandle <> 0 Then
    If ComCtlsSupportLevel() = 0 Then
        ' According to MSDN:
        ' - Version 5 of comctl32 supports deleting of column zero, but only after you use CCM_SETVERSION to set the version to 5 or later.
        ' - If you change the font by returning CDRF_NEWFONT, the list view control might display clipped text.
        '   This behavior is necessary for backward compatibility with earlier versions of the common controls.
        '   If you want to change the font of a list view control, you will get better results if you send a CCM_SETVERSION message
        '   with the wParam value set to 5 before adding any items to the control.
        SendMessage ListViewHandle, CCM_SETVERSION, 5, ByVal 0&
    End If
End If
If Ambient.UserMode = True Then
    If ListViewHandle <> 0 Then Call ComCtlsSetSubclass(ListViewHandle, Me, 1)
End If
End Sub

Private Sub DestroyListView()
If ListViewHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(ListViewHandle)
ShowWindow ListViewHandle, SW_HIDE
SetParent ListViewHandle, 0
DestroyWindow ListViewHandle
ListViewHandle = 0
ListViewHeaderHandle = 0
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Function HitTest(ByVal X As Single, ByVal Y As Single, Optional ByRef SubItemIndex As Variant) As LvwListItem
Attribute HitTest.VB_Description = "Returns a reference to the list item object located at the coordinates of X and Y."
If ListViewHandle <> 0 Then
    Dim LVHTI As LVHITTESTINFO
    With LVHTI
    .PT.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    .PT.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    If IsMissing(SubItemIndex) = True Then
        If SendMessage(ListViewHandle, LVM_HITTEST, 0, ByVal VarPtr(LVHTI)) > -1 Then
            If (.Flags And LVHT_ONITEM) <> 0 Then Set HitTest = Me.ListItems(.iItem + 1)
        End If
    Else
        Select Case VarType(SubItemIndex)
            Case vbLong, vbInteger, vbByte
                If SendMessage(ListViewHandle, LVM_SUBITEMHITTEST, 0, ByVal VarPtr(LVHTI)) > -1 Then
                    If (.Flags And LVHT_ONITEM) <> 0 Then
                        Set HitTest = Me.ListItems(.iItem + 1)
                        SubItemIndex = .iSubItem
                    End If
                End If
            Case Else
                Err.Raise 13
        End Select
    End If
    End With
End If
End Function

Public Function HitTestInsertMark(ByVal X As Single, ByVal Y As Single, Optional ByRef After As Boolean) As LvwListItem
Attribute HitTestInsertMark.VB_Description = "Returns a reference to the list item object located at the coordinates of X and Y and retrieves a value that determines where the insertion point should appear. Requires comctl32.dll version 6.1 or higher."
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim P As POINTAPI, LVIM As LVINSERTMARK
    P.X = CLng(UserControl.ScaleX(X, vbContainerPosition, vbPixels))
    P.Y = CLng(UserControl.ScaleY(Y, vbContainerPosition, vbPixels))
    With LVIM
    .cbSize = LenB(LVIM)
    SendMessage ListViewHandle, LVM_INSERTMARKHITTEST, VarPtr(P), ByVal VarPtr(LVIM)
    If .iItem > -1 Then Set HitTestInsertMark = Me.ListItems(.iItem + 1)
    After = CBool(.dwFlags = LVIM_AFTER)
    End With
End If
End Function

Public Function FindItem(ByVal Text As String, Optional ByVal Index As Long, Optional ByVal Partial As Boolean, Optional ByVal Wrap As Boolean) As LvwListItem
Attribute FindItem.VB_Description = "Finds an item in the list and returns a reference to that item."
If ListViewHandle <> 0 Then
    If Index > 0 Then Index = Index - 1
    Dim LVFI As LVFINDINFO
    With LVFI
    .psz = StrPtr(Text)
    .Flags = LVFI_STRING
    If Partial = True Then .Flags = .Flags Or LVFI_PARTIAL
    If Wrap = True Then .Flags = .Flags Or LVFI_WRAP
    End With
    Index = SendMessage(ListViewHandle, LVM_FINDITEM, Index - 1, ByVal VarPtr(LVFI))
    If Index > -1 Then Set FindItem = Me.ListItems(Index + 1)
End If
End Function

Public Function GetFirstVisible() As LvwListItem
Attribute GetFirstVisible.VB_Description = "Retrieves a reference of the first list item visible in the client area."
If ListViewHandle <> 0 Then
    If Me.ListItems.Count > 0 Then
        Select Case PropView
            Case LvwViewReport, LvwViewList
                Set GetFirstVisible = PtrToObj(Me.FListItemPtr(SendMessage(ListViewHandle, LVM_GETTOPINDEX, 0, ByVal 0&) + 1))
            Case Else
                Dim LVRC As RECT, RC As RECT, i As Long
                SendMessage ListViewHandle, LVM_GETVIEWRECT, 0, ByVal VarPtr(LVRC)
                SetRect LVRC, 0, 0, (LVRC.Right - LVRC.Left), (LVRC.Bottom - LVRC.Top)
                For i = 1 To Me.ListItems.Count
                    SetRect RC, LVIR_BOUNDS, 0, 0, 0
                    SendMessage ListViewHandle, LVM_GETITEMRECT, i - 1, ByVal VarPtr(RC)
                    If RC.Right > LVRC.Left Then
                        If RC.Left < LVRC.Right Then
                            If RC.Bottom > LVRC.Top Then
                                If RC.Top < LVRC.Bottom Then
                                    Set GetFirstVisible = PtrToObj(Me.FListItemPtr(i))
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next i
        End Select
    End If
End If
End Function

Public Function GetVisibleCount() As Long
Attribute GetVisibleCount.VB_Description = "Returns the number of fully visible list items. If the list view is in 'icon', 'small icon' or 'tile' view then the return value is the total number of list items."
If ListViewHandle <> 0 Then GetVisibleCount = SendMessage(ListViewHandle, LVM_GETCOUNTPERPAGE, 0, ByVal 0&)
End Function

Public Function GetSelectedCount() As Long
Attribute GetSelectedCount.VB_Description = "Returns the number of selected items."
If ListViewHandle <> 0 Then GetSelectedCount = SendMessage(ListViewHandle, LVM_GETSELECTEDCOUNT, 0, ByVal 0&)
End Function

Public Function GetHeaderHeight() As Single
Attribute GetHeaderHeight.VB_Description = "Retrieves the height of the header control in 'report' view."
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim RC As RECT
        GetWindowRect ListViewHeaderHandle, RC
        GetHeaderHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
    End If
End If
End Function

Public Property Get SelectedItem() As LvwListItem
Attribute SelectedItem.VB_Description = "Returns/sets a reference to the currently selected list item."
Attribute SelectedItem.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, ByVal LVNI_FOCUSED)
    If iItem > -1 Then Set SelectedItem = Me.ListItems(iItem + 1)
End If
End Property

Public Property Let SelectedItem(ByVal Value As LvwListItem)
Set Me.SelectedItem = Value
End Property

Public Property Set SelectedItem(ByVal Value As LvwListItem)
If ListViewHandle <> 0 Then
    If Not Value Is Nothing Then
        Value.Selected = True
    Else
        Dim LVI As LVITEM
        With LVI
        .Mask = LVIF_STATE
        .StateMask = LVIS_FOCUSED
        .State = Not LVIS_FOCUSED
        End With
        SendMessage ListViewHandle, LVM_SETITEMSTATE, -1, ByVal VarPtr(LVI)
        Call CheckItemFocus(0)
    End If
End If
End Property

Public Function SelectedIndices() As Collection
Attribute SelectedIndices.VB_Description = "Returns a reference to a collection containing the indexes to the selected items."
If ListViewHandle <> 0 Then
    Set SelectedIndices = New Collection
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, ByVal LVNI_SELECTED)
    Do While iItem > -1
        SelectedIndices.Add (iItem + 1)
        iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, iItem, ByVal LVNI_SELECTED)
    Loop
End If
End Function

Public Property Get HotItem() As LvwListItem
Attribute HotItem.VB_Description = "Returns/sets a reference to the currently hot list item. This is only meaningful if the hot tracking property is set to true."
Attribute HotItem.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETHOTITEM, 0, ByVal 0&)
    If iItem > -1 Then Set HotItem = Me.ListItems(iItem + 1)
End If
End Property

Public Property Let HotItem(ByVal Value As LvwListItem)
Set Me.HotItem = Value
End Property

Public Property Set HotItem(ByVal Value As LvwListItem)
If ListViewHandle <> 0 Then
    If Not Value Is Nothing Then
        Value.Hot = True
    Else
        SendMessage ListViewHandle, LVM_SETHOTITEM, -1, ByVal 0&
    End If
End If
End Property

Public Property Get SelectedColumn() As LvwColumnHeader
Attribute SelectedColumn.VB_Description = "Returns/sets a reference to the currently selected column. Requires comctl32.dll version 6.0 or higher."
Attribute SelectedColumn.VB_MemberFlags = "400"
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim Index As Long
    Index = SendMessage(ListViewHandle, LVM_GETSELECTEDCOLUMN, 0, ByVal 0&)
    If Index > -1 Then Set SelectedColumn = Me.ColumnHeaders(Index + 1)
End If
End Property

Public Property Let SelectedColumn(ByVal Value As LvwColumnHeader)
Set Me.SelectedColumn = Value
End Property

Public Property Set SelectedColumn(ByVal Value As LvwColumnHeader)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If Value Is Nothing Then
        SendMessage ListViewHandle, LVM_SETSELECTEDCOLUMN, -1, ByVal 0&
    Else
        If Not PropPicture Is Nothing Then If PropPictureAlignment = LvwPictureAlignmentTile And PropPictureWatermark = False Then Exit Property
        SendMessage ListViewHandle, LVM_SETSELECTEDCOLUMN, Value.Index - 1, ByVal 0&
    End If
End If
End Property

Public Property Get SelectionMark() As LvwListItem
Attribute SelectionMark.VB_Description = "Returns/sets the selection mark. A selection mark is that list item from which a multiple selection starts."
Attribute SelectionMark.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETSELECTIONMARK, 0, ByVal 0&)
    If iItem > -1 Then Set SelectionMark = Me.ListItems(iItem + 1)
End If
End Property

Public Property Let SelectionMark(ByVal Value As LvwListItem)
Set Me.SelectionMark = Value
End Property

Public Property Set SelectionMark(ByVal Value As LvwListItem)
If ListViewHandle <> 0 Then
    If Not Value Is Nothing Then
        Dim iItem As Long
        iItem = Value.Index - 1
        SendMessage ListViewHandle, LVM_SETSELECTIONMARK, 0, ByVal iItem
    Else
        SendMessage ListViewHandle, LVM_SETSELECTIONMARK, 0, ByVal -1&
    End If
End If
End Property

Public Property Get ColumnWidth() As Single
Attribute ColumnWidth.VB_Description = "Returns/sets the width of a column in 'list' view."
Attribute ColumnWidth.VB_MemberFlags = "400"
If ListViewMemoryColumnWidth = 0 And PropView = LvwViewList Then
    ColumnWidth = UserControl.ScaleX(SendMessage(ListViewHandle, LVM_GETCOLUMNWIDTH, 0, ByVal 0&), vbPixels, vbContainerSize)
Else
    ColumnWidth = UserControl.ScaleX(ListViewMemoryColumnWidth, vbPixels, vbContainerSize)
End If
End Property

Public Property Let ColumnWidth(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
Dim IntValue As Integer
IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If IntValue > 0 Then
    ListViewMemoryColumnWidth = IntValue
    If ListViewHandle <> 0 And PropView = LvwViewList Then SendMessage ListViewHandle, LVM_SETCOLUMNWIDTH, 0, ByVal CLng(IntValue)
Else
    Err.Raise 380
End If
End Property

Public Sub StartLabelEdit()
Attribute StartLabelEdit.VB_Description = "Begins a label editing operation on a list item. This method will fail if the label edit property is set to disabled."
If ListViewHandle <> 0 Then
    ListViewStartLabelEdit = True
    SendMessage ListViewHandle, LVM_EDITLABEL, ListViewFocusIndex - 1, ByVal 0&
    ListViewStartLabelEdit = False
End If
End Sub

Public Sub Scroll(ByVal X As Single, ByVal Y As Single)
Attribute Scroll.VB_Description = "Scrolls the content. When the list view is in 'report' view, the X and Y arguments will be rounded up to the nearest number that form a whole line increment."
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_SCROLL, CLng(UserControl.ScaleX(X, vbContainerSize, vbPixels)), ByVal CLng(UserControl.ScaleX(Y, vbContainerSize, vbPixels))
End Sub

Public Sub SetExplorerTheme()
Attribute SetExplorerTheme.VB_Description = "Method that gives the list view the appearance of the windows explorer."
If ListViewHandle <> 0 And EnabledVisualStyles() = True Then SetWindowTheme ListViewHandle, StrPtr("Explorer"), 0
End Sub

Public Sub ResetEmptyMarkup()
Attribute ResetEmptyMarkup.VB_Description = "Method to force the control to request again for a markup text. Requires comctl32.dll version 6.1 or higher."
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then SendMessage ListViewHandle, LVM_RESETEMPTYTEXT, 0, ByVal 0&
End Sub

Public Sub ComputeControlSize(ByVal VisibleCount As Long, ByRef Width As Single, ByRef Height As Single, Optional ByVal ProposedWidth As Single, Optional ByVal ProposedHeight As Single)
Attribute ComputeControlSize.VB_Description = "A method that returns the width and height for a given number of visible list items."
If VisibleCount < 0 Then Err.Raise 380
If ListViewHandle <> 0 Then
    Dim RetVal As Long, RC(0 To 1) As RECT, ProposedX As Long, ProposedY As Long
    GetWindowRect ListViewHandle, RC(0)
    GetClientRect ListViewHandle, RC(1)
    With UserControl
    If ProposedWidth <> 0 Then
        ProposedX = CLng(.ScaleX(ProposedWidth, vbContainerSize, vbPixels))
    Else
        ProposedX = -1
    End If
    If ProposedHeight <> 0 Then
        ProposedY = CLng(.ScaleY(ProposedHeight, vbContainerSize, vbPixels))
    Else
        ProposedY = -1
    End If
    RetVal = SendMessage(ListViewHandle, LVM_APPROXIMATEVIEWRECT, IIf(PropView = LvwViewReport, VisibleCount - 1, VisibleCount), MakeDWord(ProposedX, ProposedY))
    If LoWord(RetVal) <> 0 Then Width = .ScaleX(LoWord(RetVal) + ((RC(0).Right - RC(0).Left) - (RC(1).Right - RC(1).Left)), vbPixels, vbContainerSize)
    If HiWord(RetVal) <> 0 Then Height = .ScaleY(HiWord(RetVal) + ((RC(0).Bottom - RC(0).Top) - (RC(1).Bottom - RC(1).Top)), vbPixels, vbContainerSize)
    End With
End If
End Sub

Public Function TextWidth(ByVal Text As String) As Single
Attribute TextWidth.VB_Description = "Returns the text width of the given string using the current font of the list view."
If ListViewHandle <> 0 Then
    Dim Pixels As Long
    Pixels = SendMessage(ListViewHandle, LVM_GETSTRINGWIDTH, 0, ByVal StrPtr(Text))
    If Pixels > 0 Then TextWidth = UserControl.ScaleX(Pixels, vbPixels, vbContainerSize)
End If
End Function

Public Property Get DropHighlight() As LvwListItem
Attribute DropHighlight.VB_Description = "Returns/sets a reference to a list item and highlights it with the system highlight color."
Attribute DropHighlight.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, ByVal LVNI_DROPHILITED)
    If iItem > -1 Then Set DropHighlight = Me.ListItems(iItem + 1)
End If
End Property

Public Property Let DropHighlight(ByVal Value As LvwListItem)
Set Me.DropHighlight = Value
End Property

Public Property Set DropHighlight(ByVal Value As LvwListItem)
If ListViewHandle <> 0 Then
    Dim iItem As Long, LVI As LVITEM
    LVI.StateMask = LVIS_DROPHILITED
    If Not Value Is Nothing Then
        iItem = Value.Index - 1
        If iItem <> SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, LVNI_DROPHILITED) Then
            With LVI
            .State = Not LVIS_DROPHILITED
            SendMessage ListViewHandle, LVM_SETITEMSTATE, -1, ByVal VarPtr(LVI)
            If iItem > -1 Then
                .State = LVIS_DROPHILITED
                SendMessage ListViewHandle, LVM_SETITEMSTATE, iItem, ByVal VarPtr(LVI)
            End If
            End With
        End If
    Else
        LVI.State = Not LVIS_DROPHILITED
        SendMessage ListViewHandle, LVM_SETITEMSTATE, -1, ByVal VarPtr(LVI)
    End If
End If
End Property

Public Property Get InsertMark(Optional ByRef After As Boolean) As LvwListItem
Attribute InsertMark.VB_Description = "Returns/sets a reference to a list item where an insertion mark is positioned. Requires comctl32.dll version 6.1 or higher."
Attribute InsertMark.VB_MemberFlags = "400"
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVIM As LVINSERTMARK
    With LVIM
    .cbSize = LenB(LVIM)
    SendMessage ListViewHandle, LVM_GETINSERTMARK, 0, ByVal VarPtr(LVIM)
    If .iItem > -1 Then
        Set InsertMark = Me.ListItems(.iItem + 1)
        Dim Buffer(0 To 1) As Long, Flag As Long
        Buffer(0) = .dwFlags
        Buffer(1) = vbDropEffectScroll
        Flag = Buffer(0) - Buffer(1)
        After = CBool(Flag = LVIM_AFTER)
    End If
    End With
End If
End Property

Public Property Let InsertMark(Optional ByRef After As Boolean, ByVal Value As LvwListItem)
Set Me.InsertMark(After) = Value
End Property

Public Property Set InsertMark(Optional ByRef After As Boolean, ByVal Value As LvwListItem)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVIM As LVINSERTMARK
    With LVIM
    .cbSize = LenB(LVIM)
    If Value Is Nothing Then
        .iItem = -1
        .dwFlags = 0
    Else
        .iItem = Value.Index - 1
        .dwFlags = IIf(After = True, LVIM_AFTER, 0)
    End If
    End With
    SendMessage ListViewHandle, LVM_SETINSERTMARK, 0, ByVal VarPtr(LVIM)
End If
End Property

Public Property Get OLEDraggedItem() As LvwListItem
Attribute OLEDraggedItem.VB_Description = "Returns a reference to the currently dragged list item during an OLE drag/drop operation."
Attribute OLEDraggedItem.VB_MemberFlags = "400"
If ListViewDragIndex > 0 Then
    Dim Ptr As Long
    Ptr = Me.FListItemPtr(ListViewDragIndex)
    If Ptr <> 0 Then Set OLEDraggedItem = PtrToObj(Ptr)
End If
End Property

Public Property Get WorkAreas() As Variant
Attribute WorkAreas.VB_Description = "Returns/sets the working areas of the list view in 'icon' and 'small icon' view. All the client coordinates (left, top, right and bottom) are in pixels."
Attribute WorkAreas.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    Dim StructCount As Long
    SendMessage ListViewHandle, LVM_GETNUMBEROFWORKAREAS, 0, ByVal VarPtr(StructCount)
    If StructCount > 0 Then
        Dim RC() As RECT
        ReDim RC(0 To (StructCount - 1)) As RECT
        SendMessage ListViewHandle, LVM_GETWORKAREAS, StructCount, ByVal VarPtr(RC(0))
        Dim ArgList() As Long
        ReDim ArgList(0 To ((StructCount * 4) - 1)) As Long
        CopyMemory ArgList(0), ByVal VarPtr(RC(0)), StructCount * 16
        WorkAreas = ArgList()
    Else
        WorkAreas = Empty
    End If
End If
End Property

Public Property Let WorkAreas(ByVal ArgList As Variant)
If ListViewHandle <> 0 Then
    If IsArray(ArgList) Then
        Dim Ptr As Long
        CopyMemory Ptr, ByVal UnsignedAdd(VarPtr(ArgList), 8), 4
        If Ptr <> 0 Then
            Dim RetVal As Long
            CopyMemory ByVal VarPtr(RetVal), Ptr, 4
            If RetVal <> 0 Then
                Dim DimensionCount As Integer
                CopyMemory DimensionCount, ByVal Ptr, 2
                If DimensionCount = 1 Then
                    Dim Arr() As Long, Count As Long, i As Long
                    For i = LBound(ArgList) To UBound(ArgList)
                        Select Case VarType(ArgList(i))
                            Case vbLong, vbInteger, vbByte
                                If ArgList(i) >= 0 Then
                                    ReDim Preserve Arr(0 To Count) As Long
                                    Arr(Count) = ArgList(i)
                                    Count = Count + 1
                                End If
                        End Select
                    Next i
                    If Count > 0 Then
                        If Count Mod 4 = 0 Then
                            Dim StructCount As Long
                            StructCount = (Count / 4)
                            If StructCount > LV_MAX_WORKAREAS Then StructCount = LV_MAX_WORKAREAS
                            SendMessage ListViewHandle, LVM_SETWORKAREAS, StructCount, ByVal VarPtr(Arr(0))
                        Else
                            Err.Raise 5
                        End If
                    Else
                        SendMessage ListViewHandle, LVM_SETWORKAREAS, 0, ByVal 0&
                    End If
                Else
                    Err.Raise Number:=5, Description:="Array must be single dimensioned"
                End If
            Else
                Err.Raise Number:=91, Description:="Array is not allocated"
            End If
        Else
            Err.Raise 5
        End If
    ElseIf IsEmpty(ArgList) Then
        SendMessage ListViewHandle, LVM_SETWORKAREAS, 0, ByVal 0&
    Else
        Err.Raise 380
    End If
End If
End Property

Private Sub SetVisualStylesHeader()
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 And EnabledVisualStyles() = True Then
        Select Case Me.VisualStyles
            Case True
                ActivateVisualStyles ListViewHeaderHandle
            Case False
                RemoveVisualStyles ListViewHeaderHandle
        End Select
    End If
End If
End Sub

Private Sub SetColumnsSubItemIndex(Optional ByVal CountOffset As Long)
If ListViewHandle = 0 Then Exit Sub
If (Me.ColumnHeaders.Count + CountOffset) > 0 Then
    Dim LVC As LVCOLUMN
    LVC.Mask = LVCF_SUBITEM
    Dim i As Long
    For i = 1 To (Me.ColumnHeaders.Count + CountOffset)
        LVC.iSubItem = i - 1
        SendMessage ListViewHandle, LVM_SETCOLUMN, i - 1, ByVal VarPtr(LVC)
    Next i
End If
End Sub

Private Sub RebuildListItems()
If Me.ListItems.Count > 0 Then
    Dim i As Long, ii As Long
    With Me.ListItems
    For i = 1 To .Count
        With .Item(i)
        .Text = .Text
        If .ListSubItems.Count > 0 Then
            For ii = 1 To Me.ColumnHeaders.Count
                If ii <= .ListSubItems.Count Then
                    Me.FListItemText(i, ii) = .ListSubItems(ii).Text
                Else
                    Me.FListItemText(i, ii) = vbNullString
                End If
            Next ii
        End If
        End With
    Next i
    End With
    Me.Refresh
End If
End Sub

Private Sub CheckHeaderControl()
If ListViewHeaderHandle = 0 Then
    ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        If Not PropColumnHeaderIconsName = "(None)" Then
            If PropColumnHeaderIconsControl Is Nothing Then
                Me.ColumnHeaderIcons = PropColumnHeaderIconsName
            End If
        End If
        Call SetVisualStylesHeader
        Me.AllowColumnCheckboxes = PropAllowColumnCheckboxes
        Me.ClickableColumnHeaders = PropClickableColumnHeaders
        Me.HighlightColumnHeaders = PropHighlightColumnHeaders
        Me.TrackSizeColumnHeaders = PropTrackSizeColumnHeaders
        Me.ResizableColumnHeaders = PropResizableColumnHeaders
        SendMessage ListViewHandle, LVM_UPDATE, 0, ByVal 0&
        Me.Refresh
    End If
End If
End Sub

Private Sub CheckItemFocus(ByVal Index As Long)
Dim ParamValid As Boolean, ModularValid As Boolean
ParamValid = CBool(Index > 0 And Index <= Me.ListItems.Count)
ModularValid = CBool(ListViewFocusIndex > 0)
If (ParamValid = True And ModularValid = True And (Index <> ListViewFocusIndex)) Or (ParamValid Xor ModularValid) Then
    RaiseEvent ItemFocus(Me.ListItems(Index))
    ListViewFocusIndex = Index
Else
    ListViewFocusIndex = 0
End If
End Sub

Private Sub SortListItems()
If ListViewHandle <> 0 Then
    If SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&) > 0 Then
        If PropSortKey > Me.ColumnHeaders.Count Then PropSortKey = Me.ColumnHeaders.Count
        Dim Address As Long
        Select Case PropSortType
            Case LvwSortTypeBinary
                Address = ProcPtr(AddressOf LvwSortingFunctionBinary)
            Case LvwSortTypeText
                Address = ProcPtr(AddressOf LvwSortingFunctionText)
            Case LvwSortTypeNumeric
                Address = ProcPtr(AddressOf LvwSortingFunctionNumeric)
            Case LvwSortTypeCurrency
                Address = ProcPtr(AddressOf LvwSortingFunctionCurrency)
            Case LvwSortTypeDate
                Address = ProcPtr(AddressOf LvwSortingFunctionDate)
        End Select
        If Address <> 0 Then SendMessageSort ListViewHandle, LVM_SORTITEMSEX, Me, ByVal Address
    End If
End If
End Sub

Private Function ListItemsSortingFunctionBinary(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FListItemText(lParam1 + 1, PropSortKey)
Text2 = Me.FListItemText(lParam2 + 1, PropSortKey)
ListItemsSortingFunctionBinary = lstrcmp(StrPtr(Text1), StrPtr(Text2))
If PropSortOrder = LvwSortOrderDescending Then ListItemsSortingFunctionBinary = -ListItemsSortingFunctionBinary
End Function

Private Function ListItemsSortingFunctionText(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FListItemText(lParam1 + 1, PropSortKey)
Text2 = Me.FListItemText(lParam2 + 1, PropSortKey)
ListItemsSortingFunctionText = lstrcmpi(StrPtr(Text1), StrPtr(Text2))
If PropSortOrder = LvwSortOrderDescending Then ListItemsSortingFunctionText = -ListItemsSortingFunctionText
End Function

Private Function ListItemsSortingFunctionNumeric(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FListItemText(lParam1 + 1, PropSortKey)
Text2 = Me.FListItemText(lParam2 + 1, PropSortKey)
Dim DblBlank As Double
Dim Dbl1 As Double, Dbl2 As Double
On Error GoTo Handler
Dbl1 = CDbl(Text1)
Dbl2 = CDbl(Text2)
If 0& > 1& Then
Handler: Dbl1 = DblBlank: Dbl2 = DblBlank
End If
On Error GoTo 0
ListItemsSortingFunctionNumeric = Sgn(Dbl1 - Dbl2)
If PropSortOrder = LvwSortOrderDescending Then ListItemsSortingFunctionNumeric = -ListItemsSortingFunctionNumeric
End Function

Private Function ListItemsSortingFunctionCurrency(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FListItemText(lParam1 + 1, PropSortKey)
Text2 = Me.FListItemText(lParam2 + 1, PropSortKey)
Dim CurBlank As Currency
Dim Cur1 As Currency, Cur2 As Currency
On Error GoTo Handler
Cur1 = CCur(Text1)
Cur2 = CCur(Text2)
If 0& > 1& Then
Handler: Cur1 = CurBlank: Cur2 = CurBlank
End If
On Error GoTo 0
ListItemsSortingFunctionCurrency = Sgn(Cur1 - Cur2)
If PropSortOrder = LvwSortOrderDescending Then ListItemsSortingFunctionCurrency = -ListItemsSortingFunctionCurrency
End Function

Private Function ListItemsSortingFunctionDate(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FListItemText(lParam1 + 1, PropSortKey)
Text2 = Me.FListItemText(lParam2 + 1, PropSortKey)
Dim DateBlank As Date
Dim Date1 As Date, Date2 As Date
On Error GoTo Handler
Date1 = CDate(Text1)
Date2 = CDate(Text2)
If 0& > 1& Then
Handler: Date1 = DateBlank: Date2 = DateBlank
End If
On Error GoTo 0
ListItemsSortingFunctionDate = Sgn(Date1 - Date2)
If PropSortOrder = LvwSortOrderDescending Then ListItemsSortingFunctionDate = -ListItemsSortingFunctionDate
End Function

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcLabelEdit(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
    Case 10
        ISubclass_Message = ListItemsSortingFunctionBinary(wParam, lParam)
    Case 11
        ISubclass_Message = ListItemsSortingFunctionText(wParam, lParam)
    Case 12
        ISubclass_Message = ListItemsSortingFunctionNumeric(wParam, lParam)
    Case 13
        ISubclass_Message = ListItemsSortingFunctionCurrency(wParam, lParam)
    Case 14
        ISubclass_Message = ListItemsSortingFunctionDate(wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_MOUSEACTIVATE
        Static InProc As Boolean
        Dim LabelEditHandle As Long
        LabelEditHandle = Me.hWndLabelEdit
        If GetFocus() <> ListViewHandle And (GetFocus() <> LabelEditHandle Or LabelEditHandle = 0) Then
            If InProc = True Or LoWord(lParam) = HTBORDER Then WindowProcControl = MA_NOACTIVATEANDEAT: Exit Function
            Select Case HiWord(lParam)
                Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
                    On Error Resume Next
                    If Extender.CausesValidation = True Then
                        InProc = True
                        Screen.ActiveForm.ValidateControls
                        InProc = False
                        If Err.Number = 380 Then
                            WindowProcControl = MA_NOACTIVATEANDEAT
                        Else
                            SetFocusAPI UserControl.hWnd
                            WindowProcControl = MA_NOACTIVATE
                        End If
                    Else
                        SetFocusAPI UserControl.hWnd
                        WindowProcControl = MA_NOACTIVATE
                    End If
                    On Error GoTo 0
                    Exit Function
            End Select
        End If
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(0, MousePointerID(PropMousePointer))
                WindowProcControl = 1
                Exit Function
            ElseIf PropMousePointer = 99 Then
                If Not PropMouseIcon Is Nothing Then
                    SetCursor PropMouseIcon.Handle
                    WindowProcControl = 1
                    Exit Function
                End If
            End If
        End If
    Case WM_KEYDOWN, WM_KEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftState())
        ElseIf wMsg = WM_KEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftState())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        KeyChar = CUIntToInt(wParam And &HFFFF&)
        RaiseEvent KeyPress(KeyChar)
        wParam = CIntToUInt(KeyChar)
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftState(), X, Y)
                ListViewButtonDown = vbLeftButton
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftState(), X, Y)
                ListViewButtonDown = vbMiddleButton
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftState(), X, Y)
                ListViewButtonDown = vbRightButton
            Case WM_MOUSEMOVE
                RaiseEvent MouseMove(GetMouseState(), GetShiftState(), X, Y)
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
                Select Case wMsg
                    Case WM_LBUTTONUP
                        RaiseEvent MouseUp(vbLeftButton, GetShiftState(), X, Y)
                    Case WM_MBUTTONUP
                        RaiseEvent MouseUp(vbMiddleButton, GetShiftState(), X, Y)
                    Case WM_RBUTTONUP
                        RaiseEvent MouseUp(vbRightButton, GetShiftState(), X, Y)
                End Select
                Dim P As POINTAPI
                GetCursorPos P
                If WindowFromPoint(P.X, P.Y) = hWnd Then RaiseEvent Click
        End Select
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = ListViewHeaderHandle Then
            Dim Cancel As Boolean
            Dim NMHDR As NMHEADER
            Select Case NM.Code
                Case HDN_BEGINTRACK
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then
                        RaiseEvent ColumnBeforeSize(Me.ColumnHeaders(NMHDR.iItem + 1), Cancel)
                        If Cancel = True Then
                            WindowProcControl = 1
                            Exit Function
                        End If
                    End If
                Case HDN_ENDTRACK
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then RaiseEvent ColumnAfterSize(Me.ColumnHeaders(NMHDR.iItem + 1))
                Case HDN_BEGINDRAG
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then RaiseEvent ColumnBeforeDrag(Me.ColumnHeaders(NMHDR.iItem + 1))
                Case HDN_ENDDRAG
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then
                        Dim HDI As HDITEM
                        CopyMemory HDI, ByVal NMHDR.lPtrHDItem, LenB(HDI)
                        RaiseEvent ColumnAfterDrag(Me.ColumnHeaders(NMHDR.iItem + 1), HDI.iOrder + 1, Cancel)
                        If Cancel = True Then
                            WindowProcControl = 1
                            Exit Function
                        End If
                    End If
                Case HDN_DROPDOWN
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then RaiseEvent ColumnDropDown(Me.ColumnHeaders(NMHDR.iItem + 1))
                Case HDN_ITEMSTATEICONCLICK
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then
                        With Me.ColumnHeaders(NMHDR.iItem + 1)
                        .Checked = Not .Checked
                        End With
                        RaiseEvent ColumnCheck(Me.ColumnHeaders(NMHDR.iItem + 1))
                        Exit Function
                    End If
            End Select
        End If
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcLabelEdit(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg = WM_IME_CHAR Then SendMessage hWnd, WM_CHAR, wParam, ByVal lParam: Exit Function
WindowProcLabelEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SHOWWINDOW
        If ListViewHandle = 0 Then Call CreateListView
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = ListViewHandle Then
            Dim Length As Long, Cancel As Boolean
            Dim NMLV As NMLISTVIEW, NMIA As NMITEMACTIVATE, NMLVDI As NMLVDISPINFO
            Select Case NM.Code
                Case LVN_INSERTITEM
                    If ListViewListItemsControl = 0 Then
                        Me.FListItemSelected(1) = True
                        ListViewFocusIndex = 1
                    End If
                    ListViewListItemsControl = ListViewListItemsControl + 1
                Case LVN_DELETEITEM
                    ListViewListItemsControl = ListViewListItemsControl - 1
                Case LVN_ITEMCHANGED
                    CopyMemory NMLV, ByVal lParam, LenB(NMLV)
                    With NMLV
                    If .uChanged = LVIF_STATE Then
                        If CBool((.uNewState And LVIS_FOCUSED) = LVIS_FOCUSED) Xor CBool((.uOldState And LVIS_FOCUSED) = LVIS_FOCUSED) Then
                            If (.uNewState And LVIS_FOCUSED) = LVIS_FOCUSED Then Call CheckItemFocus(.iItem + 1)
                        End If
                        If CBool((.uNewState And LVIS_SELECTED) = LVIS_SELECTED) Xor CBool((.uOldState And LVIS_SELECTED) = LVIS_SELECTED) Then
                            Me.FListItemRedraw .iItem + 1
                            RaiseEvent ItemSelect(Me.ListItems(.iItem + 1), CBool((.uNewState And LVIS_SELECTED) = LVIS_SELECTED))
                        End If
                        If CBool((.uNewState And &H2000&) = &H2000&) Xor CBool((.uOldState And &H2000&) = &H2000&) Then RaiseEvent ItemCheck(Me.ListItems(.iItem + 1), CBool((.uNewState And &H2000&) = &H2000&))
                    End If
                    End With
                Case LVN_BEGINLABELEDIT, LVN_ENDLABELEDIT
                    Static LabelEditHandle As Long
                    Select Case NM.Code
                        Case LVN_BEGINLABELEDIT
                            If PropLabelEdit = LvwLabelEditManual And ListViewStartLabelEdit = False Then
                                WindowProcUserControl = 1
                            Else
                                RaiseEvent BeforeLabelEdit(Cancel)
                                If Cancel = True Then
                                    WindowProcUserControl = 1
                                Else
                                    WindowProcUserControl = 0
                                    LabelEditHandle = Me.hWndLabelEdit
                                    If LabelEditHandle <> 0 Then Call ComCtlsSetSubclass(LabelEditHandle, Me, 2)
                                    ListViewLabelInEdit = True
                                End If
                            End If
                        Case LVN_ENDLABELEDIT
                            CopyMemory NMLVDI, ByVal lParam, LenB(NMLVDI)
                            With NMLVDI.Item
                            If .pszText <> 0 Then
                                Dim NewText As String
                                Length = lstrlen(.pszText)
                                NewText = String(Length, vbNullChar)
                                CopyMemory ByVal StrPtr(NewText), ByVal .pszText, Length * 2
                                RaiseEvent AfterLabelEdit(Cancel, NewText)
                                If Cancel = False Then
                                    With Me.ListItems(.iItem + 1)
                                    .FInit Me, .Index, .Key, NMLVDI.Item.lParam, NewText, .Icon, .SmallIcon
                                    End With
                                    WindowProcUserControl = 1
                                Else
                                    WindowProcUserControl = 0
                                End If
                            Else
                                WindowProcUserControl = 0
                            End If
                            End With
                            If LabelEditHandle <> 0 Then
                                Call ComCtlsRemoveSubclass(LabelEditHandle)
                                LabelEditHandle = 0
                            End If
                            ListViewLabelInEdit = False
                    End Select
                    Exit Function
                Case LVN_BEGINDRAG, LVN_BEGINRDRAG
                    CopyMemory NMLV, ByVal lParam, LenB(NMLV)
                    If NMLV.iItem > -1 Then
                        ListViewDragIndexBuffer = NMLV.iItem + 1
                        If NM.Code = LVN_BEGINDRAG Then
                            RaiseEvent ItemDrag(Me.ListItems(NMLV.iItem + 1), vbLeftButton)
                            If PropOLEDragMode = vbOLEDragAutomatic Then Me.OLEDrag
                        ElseIf NM.Code = LVN_BEGINRDRAG Then
                            RaiseEvent ItemDrag(Me.ListItems(NMLV.iItem + 1), vbRightButton)
                        End If
                        ListViewDragIndexBuffer = 0
                    End If
                Case LVN_COLUMNCLICK
                    CopyMemory NMLV, ByVal lParam, LenB(NMLV)
                    RaiseEvent ColumnClick(Me.ColumnHeaders(NMLV.iSubItem + 1))
                Case LVN_ITEMACTIVATE
                    CopyMemory NMIA, ByVal lParam, LenB(NMIA)
                    Dim Shift As Integer
                    Select Case NMIA.uKeyFlags
                        Case LVKF_ALT
                            Shift = vbAltMask
                        Case LVKF_CONTROL
                            Shift = vbCtrlMask
                        Case LVKF_SHIFT
                            Shift = vbShiftMask
                    End Select
                    RaiseEvent ItemActivate(Me.ListItems(NMIA.iItem + 1), NMIA.iSubItem, Shift)
                Case NM_CLICK, NM_RCLICK
                    CopyMemory NMIA, ByVal lParam, LenB(NMIA)
                    If NMIA.iItem > -1 Then
                        If NM.Code = NM_CLICK Then
                            RaiseEvent ItemClick(Me.ListItems(NMIA.iItem + 1), vbLeftButton)
                        ElseIf NM.Code = NM_RCLICK Then
                            RaiseEvent ItemClick(Me.ListItems(NMIA.iItem + 1), vbRightButton)
                        End If
                    End If
                    If NMIA.iItem > -1 Or (NMIA.iItem = -1 And (PropView = LvwViewReport Or PropView = LvwViewList)) Then
                        Dim P1 As POINTAPI
                        GetCursorPos P1
                        ScreenToClient ListViewHandle, P1
                        RaiseEvent MouseUp(ListViewButtonDown, GetShiftState(), UserControl.ScaleX(P1.X, vbPixels, vbTwips), UserControl.ScaleY(P1.Y, vbPixels, vbTwips))
                        ListViewButtonDown = 0
                        RaiseEvent Click
                    End If
                Case NM_DBLCLK, NM_RDBLCLK
                    CopyMemory NMIA, ByVal lParam, LenB(NMIA)
                    If NMIA.iItem > -1 Then
                        If NM.Code = NM_DBLCLK Then
                            RaiseEvent ItemDblClick(Me.ListItems(NMIA.iItem + 1), vbLeftButton)
                        ElseIf NM.Code = NM_RDBLCLK Then
                            RaiseEvent ItemDblClick(Me.ListItems(NMIA.iItem + 1), vbRightButton)
                        End If
                    End If
                    RaiseEvent DblClick
                Case NM_CUSTOMDRAW
                    Dim FontHandle As Long
                    Dim ListItem As LvwListItem
                    Dim NMLVCD As NMLVCUSTOMDRAW
                    CopyMemory NMLVCD, ByVal lParam, LenB(NMLVCD)
                    Select Case NMLVCD.NMCD.dwDrawStage
                        Case CDDS_PREPAINT
                            WindowProcUserControl = CDRF_NOTIFYITEMDRAW
                            Exit Function
                        Case CDDS_ITEMPREPAINT
                            FontHandle = ListViewFontHandle
                            If NMLVCD.NMCD.lItemlParam <> 0 Then
                                Set ListItem = PtrToObj(NMLVCD.NMCD.lItemlParam)
                                With ListItem
                                If NMLVCD.iSubItem = 0 Then
                                    If (NMLVCD.NMCD.uItemState And CDIS_HOT) = 0 Then
                                        If .Bold = True Then FontHandle = ListViewBoldFontHandle
                                        NMLVCD.ClrText = WinColor(.ForeColor)
                                    Else
                                        If PropUnderlineHot = True Then
                                            If .Bold = True Then
                                                FontHandle = ListViewBoldUnderlineFontHandle
                                            Else
                                                FontHandle = ListViewUnderlineFontHandle
                                            End If
                                        Else
                                            If .Bold = True Then FontHandle = ListViewBoldFontHandle
                                        End If
                                        If PropHighlightHot = False Then NMLVCD.ClrText = WinColor(.ForeColor)
                                    End If
                                    RaiseEvent ItemBkColor(ListItem, NMLVCD.ClrTextBk)
                                End If
                                End With
                            End If
                            SelectObject NMLVCD.NMCD.hDC, FontHandle
                            CopyMemory ByVal lParam, NMLVCD, LenB(NMLVCD)
                            WindowProcUserControl = CDRF_NEWFONT Or CDRF_NOTIFYSUBITEMDRAW
                            Exit Function
                        Case (CDDS_ITEMPREPAINT Or CDDS_SUBITEM)
                            FontHandle = ListViewFontHandle
                            If NMLVCD.NMCD.lItemlParam <> 0 Then
                                Set ListItem = PtrToObj(NMLVCD.NMCD.lItemlParam)
                                With ListItem
                                If NMLVCD.iSubItem > 0 Then
                                    With .ListSubItems
                                    If .Count > 0 Then
                                        If NMLVCD.iSubItem <= .Count Then
                                            If (NMLVCD.NMCD.uItemState And CDIS_HOT) = 0 Then
                                                If .Item(NMLVCD.iSubItem).Bold = True Then FontHandle = ListViewBoldFontHandle
                                                NMLVCD.ClrText = WinColor(.Item(NMLVCD.iSubItem).ForeColor)
                                            Else
                                                If PropUnderlineHot = True Then
                                                    If .Item(NMLVCD.iSubItem).Bold = True Then
                                                        FontHandle = ListViewBoldUnderlineFontHandle
                                                    Else
                                                        FontHandle = ListViewUnderlineFontHandle
                                                    End If
                                                Else
                                                    If .Item(NMLVCD.iSubItem).Bold = True Then FontHandle = ListViewBoldFontHandle
                                                End If
                                                If PropHighlightHot = False Then NMLVCD.ClrText = WinColor(.Item(NMLVCD.iSubItem).ForeColor)
                                            End If
                                        End If
                                    End If
                                    End With
                                Else
                                    If (NMLVCD.NMCD.uItemState And CDIS_HOT) = 0 Then
                                        If .Bold = True Then FontHandle = ListViewBoldFontHandle
                                        NMLVCD.ClrText = WinColor(.ForeColor)
                                    Else
                                        If PropUnderlineHot = True Then
                                            If .Bold = True Then
                                                FontHandle = ListViewBoldUnderlineFontHandle
                                            Else
                                                FontHandle = ListViewUnderlineFontHandle
                                            End If
                                        Else
                                            If .Bold = True Then FontHandle = ListViewBoldFontHandle
                                        End If
                                        If PropHighlightHot = False Then NMLVCD.ClrText = WinColor(.ForeColor)
                                    End If
                                    RaiseEvent ItemBkColor(ListItem, NMLVCD.ClrTextBk)
                                End If
                                End With
                            End If
                            SelectObject NMLVCD.NMCD.hDC, FontHandle
                            CopyMemory ByVal lParam, NMLVCD, LenB(NMLVCD)
                            WindowProcUserControl = CDRF_NEWFONT
                            Exit Function
                    End Select
                Case NM_SETFOCUS, NM_KILLFOCUS
                    If PropView = LvwViewReport Then
                        If PropFullRowSelect = False Then
                            If Not PropSmallIconsName = "(None)" Then
                                If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_REDRAWITEMS, 0, ByVal SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&)
                            End If
                        End If
                    End If
                Case LVN_GETINFOTIP
                    Dim NMLVGIT As NMLVGETINFOTIP
                    CopyMemory NMLVGIT, ByVal lParam, LenB(NMLVGIT)
                    With NMLVGIT
                    If .iItem > -1 And .pszText <> 0 Then
                        If .dwFlags = LVGIT_UNFOLDED Then
                            Dim ToolTipText As String
                            ToolTipText = Me.ListItems(.iItem + 1).ToolTipText
                            If Not ToolTipText = vbNullString Then
                                ToolTipText = ToolTipText & vbNullChar
                                Length = LenB(ToolTipText)
                                If Length > .cchTextMax Then Length = .cchTextMax
                                If Length > 0 Then CopyMemory ByVal .pszText, ByVal StrPtr(ToolTipText), Length
                            Else
                                CopyMemory ByVal .pszText, 0&, 4
                            End If
                        End If
                    End If
                    End With
                Case LVN_GETDISPINFO
                    CopyMemory NMLVDI, ByVal lParam, LenB(NMLVDI)
                    With NMLVDI.Item
                    If .iItem > -1 Then
                        If .iSubItem = 0 Then
                            Select Case PropView
                                Case LvwViewIcon, LvwViewTile
                                    .iImage = Me.ListItems(.iItem + 1).Icon - 1
                                Case LvwViewSmallIcon, LvwViewList, LvwViewReport
                                    .iImage = Me.ListItems(.iItem + 1).SmallIcon - 1
                            End Select
                        Else
                            With Me.ListItems(.iItem + 1).ListSubItems
                            If NMLVDI.Item.iSubItem <= .Count Then NMLVDI.Item.iImage = .Item(NMLVDI.Item.iSubItem).ReportIcon - 1
                            End With
                        End If
                        CopyMemory ByVal lParam, NMLVDI, LenB(NMLVDI)
                    End If
                    End With
                Case LVN_GETEMPTYMARKUP
                    Dim Text As String, Centered As Boolean
                    RaiseEvent GetEmptyMarkup(Text, Centered)
                    If Not Text = vbNullString Then
                        Dim NMLVEMU As NMLVEMPTYMARKUP
                        CopyMemory NMLVEMU, ByVal lParam, LenB(NMLVEMU)
                        If Len(Text) > L_MAX_URL_LENGTH Then
                            Length = L_MAX_URL_LENGTH * 2
                        Else
                            Length = LenB(Text)
                        End If
                        Dim TextB() As Byte
                        TextB() = Text
                        CopyMemory NMLVEMU.szMarkup(0), TextB(0), Length
                        If Centered = True Then NMLVEMU.dwFlags = EMF_CENTERED
                        CopyMemory ByVal lParam, NMLVEMU, LenB(NMLVEMU)
                        WindowProcUserControl = 1
                        Exit Function
                    End If
                Case LVN_MARQUEEBEGIN
                    RaiseEvent BeginMarqueeSelection(Cancel)
                    If Cancel = True Then
                        WindowProcUserControl = 1
                    Else
                        WindowProcUserControl = 0
                    End If
                    Exit Function
            End Select
        End If
    Case WM_CONTEXTMENU
        If wParam = ListViewHandle Then
            Dim P2 As POINTAPI
            P2.X = Get_X_lParam(lParam)
            P2.Y = Get_Y_lParam(lParam)
            If P2.X > 0 And P2.Y > 0 Then
                ScreenToClient ListViewHandle, P2
                RaiseEvent ContextMenu(UserControl.ScaleX(P2.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P2.Y, vbPixels, vbContainerPosition))
            ElseIf P2.X = -1 And P2.Y = -1 Then
                ' According to MSDN:
                ' If the context menu is generated from the keyboard - for example
                ' if the user types SHIFT + F10 — then the X and Y coordinates
                ' are -1 and the application should display the context menu at the
                ' location of the current selection rather than at (XPos, YPos).
                RaiseEvent ContextMenu(-1, -1)
            End If
        End If
    Case WM_NOTIFYFORMAT
        Const NF_QUERY As Long = 3
        If lParam = NF_QUERY Then
            Const NFR_UNICODE As Long = 2
            Const NFR_ANSI As Long = 1
            WindowProcUserControl = NFR_UNICODE
            Exit Function
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS Then SetFocusAPI ListViewHandle
End Function
