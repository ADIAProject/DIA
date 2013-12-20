Attribute VB_Name = "mSortListView"
Option Explicit

Public sOrder                 As Boolean
Public lSortAs                As Long
Public m_lColumn              As Long
Public m_PRECEDE              As Long
Public m_FOLLOW               As Long

Private Const LVM_FIRST       As Long = &H1000
Private Const LVIF_TEXT       As Long = &H1
Private Const LVIF_IMAGE      As Long = &H2
Private Const LVIF_PARAM      As Long = &H4
Private Const LVIF_STATE      As Long = &H8
Private Const LVIF_INDENT     As Long = &H10
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)

Private Type LVITEM_lp
    Mask                                    As Long
    iItem                               As Long
    iSubItem                            As Long
    State                               As Long
    StateMask                           As Long
    pszText                             As Long
    cchTextMax                          As Long
    iImage                              As Long
    lParam                              As Long
    iIndent                             As Long
End Type

Private m_uLVI As LVITEM_lp

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function hSortFunc
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lParam1 (Long)
'                              lParam2 (Long)
'                              lngHWnd (Long)
'!--------------------------------------------------------------------------------
Public Function hSortFunc(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal lngHWnd As Long) As Long

    Select Case lSortAs

        Case stText
            hSortFunc = IIf(pvGetItemText(lngHWnd, lParam1) > pvGetItemText(lngHWnd, lParam2), m_PRECEDE, m_FOLLOW)

        Case stIndex
            hSortFunc = IIf(lParam1 > lParam2, m_PRECEDE, m_FOLLOW)

        Case stDate
            hSortFunc = IIf(CDate(pvGetItemText(lngHWnd, lParam1)) > CDate(pvGetItemText(lngHWnd, lParam2)), m_PRECEDE, m_FOLLOW)

        Case stNumber
            'hSortFunc = IIf(Val(pvGetItemText(lngHWnd, lParam1)) > Val(pvGetItemText(lngHWnd, lParam2)), m_PRECEDE, m_FOLLOW)
            hSortFunc = IIf(CLng(pvGetItemText(lngHWnd, lParam1)) > CLng(pvGetItemText(lngHWnd, lParam2)), m_PRECEDE, m_FOLLOW)
    End Select

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function pvGetItemText
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lngHWnd (Long)
'                              lParam (Long)
'!--------------------------------------------------------------------------------
Private Function pvGetItemText(ByVal lngHWnd As Long, ByVal lParam As Long) As String

    Dim a(261) As Byte
    Dim lLen   As Long

    With m_uLVI
        .Mask = LVIF_TEXT
        .pszText = VarPtr(a(0))
        .cchTextMax = UBound(a)
        .iSubItem = m_lColumn
    End With

    'M_ULVI
    lLen = SendMessage(lngHWnd, LVM_GETITEMTEXT, lParam, m_uLVI)
    pvGetItemText = VBA.Left$(StrConv(a(), vbUnicode), lLen)
End Function
