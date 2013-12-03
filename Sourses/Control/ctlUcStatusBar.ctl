VERSION 5.00
Begin VB.UserControl ctlUcStatusBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DrawWidth       =   56
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ctlUcStatusBar.ctx":0000
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   600
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "ctlUcStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'+  File Description:
'       ucStatusBar - A Selfsubclassed Theme Aware ucStatusBar Control which Provides Dynamic Properties
'
'   Product Name:
'       ucStatusBar.ctl
'
'   Compatability:
'       Windows: 9x, ME, NT, 2K, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'       (Paul Caton - Self-Subclassser)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'       (Randy Birch - IsWinXP)
'           http://vbnet.mvps.org/code/system/getversionex.htm
'       (Dieter Otter - GetCurrentThemeName)
'           http://www.vbarchiv.net/archiv/tipp_805.html
'
'   Legal Copyright & Trademarks:
'       Copyright © 2007, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2007, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Advance Research Systems shall not be liable for
'       any incidental or consequential damages suffered by any use of this software.
'       This software is owned by Paul R. Territo, Ph.D and is free for use
'       in accordance with the terms of the License Agreement in the accompanying
'       documentation.
'
'   Contact Information:
'       For Technical Assistance:
'       pwterrito@insightbb.com
'
'-  Modification(s) History:
'
'       13Jul07 - Initial Usercontrol Build
'       14Jul07 - Fixed Aligment bug in the PanelAlign method which passed the wrong constant values to the
'                 drawing routines.
'               - Added Private StatusBar constants for clarity of the text alignments
'               - Added Theme Support (non-subclassed).
'               - Added usbClassic Theme Style for Win9x drawing support.
'               - Added Version property
'               - Added HitTest for Events to allow for determining which panel we are over
'               - Optimized the Drawing routines to prevent flicker on resize
'               - Added All Normal UserControl Events
'               - Added Panel Specific Events
'       15Jul07 - Added BoundControl Method for Binding External Objects into Panels
'               - Added Boundry checking for the Index property variables to ensure we are in bounds
'               - Optimized PaintPanels method to group activities by Icon or BoundObject states.
'               - Optimized BoundObject handling for resizing and auto hide if the control has
'                 a minimum width property...like ComboBoxes etc...
'               - Added Subclass support for SysColor, Theme, NonClient Paint uMsgs
'               - Added MouseEnter & MouseExit events with subclasser uMsgs
'               - Added Editable Property and updated AddPanel to reflect this
'               - Added txtEdit to allow for direct Panel modifications in DblClick.
'       16Jul07 - Added addtional drawing optimizations for painting in the IDE
'               - Added Theme Color Specific AlphaBlends for the top line of the gradient under XP LnF.
'               - Added alignmnet adjustments for Edit TextBox in usbClassic theme
'               - Added Auto selection of text on focus for Edit TextBox
'               - Fixed BoundObject Width in usbClassic theme
'               - Removed AutoHide of BoundObject when usbNoSize
'               - Fixed Grip highlight Painting for usbClassic theme
'       17Jul07 - Added painting refinements to the top gradient within PaintGradient
'       03Aug07 - Added Sizable property to allow for removal of this functionality
'       08Aug07 - Fixed Minor Redraw bug in the Refresh method which did not allow all panels
'                 to repaint correctly when updated.
'
'   Build Date & Time: 8/3/2007 11:43:17 AM
Const Major                             As Long = 1
Const Minor                             As Long = 0
Const Revision                          As Long = 60
Const DateTime                          As String = "8/3/2007 11:43:17 AM "

'
'   Force Declarations
Option Explicit

'   MouseDown Message Constants for Corner Drag
Private Const HTBOTTOMRIGHT = 17
Private Const WM_NCLBUTTONDOWN = &HA1

'   Private Local StatusBar Text Alignment Constants
Private Const DT_SB_LEFT = (DT_VCENTER Or DT_LEFT Or DT_WORD_ELLIPSIS Or DT_WORDBREAK)
Private Const DT_SB_CENTER = (DT_VCENTER Or DT_CENTER Or DT_WORD_ELLIPSIS Or DT_WORDBREAK)
Private Const DT_SB_RIGHT = (DT_VCENTER Or DT_RIGHT Or DT_WORD_ELLIPSIS Or DT_WORDBREAK)

'   Constants used by new transparent support in NT.
Private Const CAPS1 = 94                 '  other caps
Private Const C1_TRANSPARENT = &H1       '  new raster cap
Private Const NEWTRANSPARENT = 3         '  use with SetBkMode()

'   Ternary raster operations
Private Const SRCCOPY = &HCC0020         ' (DWORD) dest = source

Public Enum usbAlignEnum
    usbLeft = &H0
    usbCenter = &H1
    usbRight = &H2

End Enum

#If False Then

    Const usbLeft = &H0
    Const usbCenter = &H1
    Const usbRight = &H2

#End If

Public Enum usbGripEnum
    usbNone = &H0
    usbSquare = &H1
    usbBars = &H2

End Enum

#If False Then

    Const usbNone = &H0
    Const usbSquare = &H1
    Const usbBars = &H2

#End If

Public Enum usbSizeEnum
    usbNoSize = &H0
    usbAutoSize = &H1

End Enum

#If False Then

    Const usbNoSize = &H0
    Const usbAutoSize = &H1

#End If

Public Enum usbStateEnum
    usbEnabled = &H0
    usbDisabled = &H1

End Enum

#If False Then

    Const usbEnabled = &H0
    Const usbDisabled = &H1

#End If

Public Enum usbThemeEnum
    usbAuto = &H0
    usbClassic = &H1
    usbBlue = &H2
    usbHomeStead = &H3
    usbMetallic = &H4

End Enum

#If False Then

    Const usbAuto = &H0
    Const usbClassic = &H1
    Const usbBlue = &H2
    Const usbHomeStead = &H3
    Const usbMetallic = &H4

#End If

'   Private StatusBar Item Type
Private Type PanelItem
    Alignment                               As Long
    AutoSize                            As Boolean
    BoundObject                         As Object
    BoundParent                         As Long
    BoundSize                           As usbSizeEnum
    Editable                            As Boolean
    ForeColor                           As OLE_COLOR
    Font                                As StdFont
    Icon                                As StdPicture
    IconState                           As usbStateEnum
    ItemRect                            As RECT
    MaskColor                           As OLE_COLOR
    Text                                As String
    ToolTipText                         As String
    UseMaskColor                        As Boolean
    Width                               As Long

End Type

Private m_ActivePanel                   As Long             'Current Active Panel
Private m_BackColor                     As OLE_COLOR        'UserControl BackColor
Private m_Forecolor                     As OLE_COLOR        'UserControl ForeColor
Private m_Font                          As StdFont          'UserControl Font
Private m_GripRect                      As RECT             'Grip Retangle
Private m_GripShape                     As usbGripEnum      'Grip Shape...Auto Set when Theme is Set
Private m_Sizable                       As Boolean          'Resizable
Private m_PanelCount                    As Long             'Panel Count
Private m_PanelItems()                  As PanelItem        'Panel Items
Private m_Theme                         As usbThemeEnum     'Theme Set by the User
Private m_iTheme                        As usbThemeEnum     'Theme Stored internally for determination of named themes + auto equivelant

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PanelClick(Index As Long)
Public Event PanelDblClick(Index As Long)
Public Event PanelMouseDown(Index As Long, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

Public Event PanelMouseMove(Index As Long, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

Public Event PanelMouseUp(Index As Long, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

'==================================================================================================
' ucSubclass - A template UserControl for control authors that require self-subclassing without ANY
'              external dependencies. IDE safe.
'
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0.0000 20040525 First cut.....................................................................
' v1.1.0000 20040602 Multi-subclassing version.....................................................
' v1.1.0001 20040604 Optimized the subclass code...................................................
' v1.1.0002 20040607 Substituted byte arrays for strings for the code buffers......................
' v1.1.0003 20040618 Re-patch when adding extra hWnds..............................................
' v1.1.0004 20040619 Optimized to death version....................................................
' v1.1.0005 20040620 Use allocated memory for code buffers, no need to re-patch....................
' v1.1.0006 20040628 Better protection in zIdx, improved comments..................................
' v1.1.0007 20040629 Fixed InIDE patching oops.....................................................
' v1.1.0008 20040910 Fixed bug in UserControl_Terminate, zSubclass_Proc procedure hidden...........
'==================================================================================================
'Subclasser declarations
Public Event MouseEnter()
Public Event MouseLeave()

Private bTrack                          As Boolean
Private bTrackUser32                    As Boolean
Private bInCtrl                         As Boolean
Private bSubClass                       As Boolean

Private Enum eMsgWhen
MSG_AFTER = 1                                                                   'Message calls back after the original (previous) WndProc
MSG_BEFORE = 2                                                                  'Message calls back before the original (previous) WndProc
MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                  'Message calls back before and after the original (previous) WndProc

End Enum

#If False Then

    Const MSG_AFTER = 1                                                                   'Message calls back after the original (previous) WndProc
    Const MSG_BEFORE = 2                                                                  'Message calls back before the original (previous) WndProc
    Const MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                  'Message calls back before and after the original (previous) WndProc

#End If

Private Type tSubData                                                               'Subclass data type
hWnd                                    As Long                                      'Handle of the window being subclassed
nAddrSub                                As Long                                      'The address of our new WndProc (allocated memory).
nAddrOrig                               As Long                                      'The address of the pre-existing WndProc
nMsgCntA                                As Long                                      'Msg after table entry count
nMsgCntB                                As Long                                      'Msg before table entry count
aMsgTblA()                              As Long                                      'Msg after table array
aMsgTblB()                              As Long                                      'Msg Before table array

End Type

Private m_bIsWinXpOrLater As Boolean
Private sc_aSubData()                   As tSubData                                    'Subclass data array

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, _
                          ByRef bHandled As Boolean, _
                          ByRef lReturn As Long, _
                          ByRef lng_hWnd As Long, _
                          ByRef uMsg As Long, _
                          ByRef wParam As Long, _
                          ByRef lParam As Long)

'Parameters:
'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
'hWnd     - The window handle
'uMsg     - The message number
'wParam   - Message related data
'lParam   - Message related data
'Notes:
'If you really know what you're doing, it's possible to change the values of the
'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
'values get passed to the default handler.. and optionaly, the 'after' callback
    Select Case uMsg

        Case WM_MOUSEMOVE

            If Not bInCtrl Then
                bInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
                RaiseEvent MouseEnter

            End If

        Case WM_MOUSELEAVE
            bInCtrl = False
            RaiseEvent MouseLeave

        Case WM_NCPAINT
            Refresh

        Case WM_SIZING
            Refresh

        Case WM_SYSCOLORCHANGE
            Refresh

        Case WM_THEMECHANGED
            Refresh

    End Select

End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines
'Add a message to the table of those that will invoke a callback. You should Subclass_Subclass first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, _
                            ByVal uMsg As Long, _
                            Optional ByVal When As eMsgWhen = MSG_AFTER)

'Parameters:
'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))

        If When And eMsgWhen.MSG_BEFORE Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)

        End If

        If When And eMsgWhen.MSG_AFTER Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)

        End If

    End With

End Sub

Private Function IsFunctionExported(ByVal sFunction As String, _
                                    ByVal sModule As String) As Boolean

Dim hMod                            As Long
Dim bLibLoaded                      As Boolean
Dim lngStrPtr                       As Long

    lngStrPtr = StrPtr(sModule)
    hMod = GetModuleHandle(lngStrPtr)

    If hMod = 0 Then
        hMod = LoadLibrary(lngStrPtr)

        If hMod Then
            bLibLoaded = True

        End If

    End If

    If hMod Then
        If GetProcAddress(hMod, sFunction) Then
            IsFunctionExported = True

        End If

    End If

    If bLibLoaded Then
        Call FreeLibrary(hMod)

    End If

End Function

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)

End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long

'Parameters:
'lng_hWnd  - The handle of the window to be subclassed
'Returns;
'The sc_aSubData() index
Const CODE_LEN                          As Long = 200

    'Length of the machine code in bytes
Const FUNC_CWP                          As String = "CallWindowProcA"

    'We use CallWindowProc to call the original WndProc
Const FUNC_EBM                          As String = "EbMode"

    'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_SWL                          As String = "SetWindowLongA"

    'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
Const MOD_USER                          As String = "user32.dll"

    'Location of the SetWindowLongA & CallWindowProc functions
Const MOD_VBA5                          As String = "vba5"

    'Location of the EbMode function if running VB5
Const MOD_VBA6                          As String = "vba6"

    'Location of the EbMode function if running VB6
Const PATCH_01                          As Long = 18

    'Code buffer offset to the location of the relative address to EbMode
Const PATCH_02                          As Long = 68

    'Address of the previous WndProc
Const PATCH_03                          As Long = 78

    'Relative address of SetWindowsLong
Const PATCH_06                          As Long = 116

    'Address of the previous WndProc
Const PATCH_07                          As Long = 121

    'Relative address of CallWindowProc
Const PATCH_0A                          As Long = 186

    'Address of the owner object
Static aBuf(1 To CODE_LEN)              As Byte

    'Static code buffer byte array
Static pCWP                             As Long

    'Address of the CallWindowsProc
Static pEbMode                          As Long

    'Address of the EbMode IDE break/stop/running function
Static pSWL                             As Long

    'Address of the SetWindowsLong function
Dim i                                   As Long

    'Loop index
Dim j                                   As Long

    'Loop index
Dim nSubIdx                             As Long

    'Subclass data index
Dim sHex                                As String
Dim miLensHex                           As Integer

    'Hex code string
    'If it's the first time through here..
    If aBuf(1) = 0 Then
        'The hex pair machine code representation.
        'sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        'Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1

        '        miLensHex = Len(sHex)
        '        Do While j < CODE_LEN
        '            j = j + 1
        '            If i < miLensHex Then
        '                'aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
        '                aBuf(j) = CLng("&H" & Mid$(sHex, i, 2))
        '            Else
        '                aBuf(j) = 0
        '            End If
        '
        '            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
        '            i = i + 2
        '        Loop

        For j = 1 To CODE_LEN
            'bytBuffer(lngCount) = Val("&H" & Left$(strHex, 2))
            aBuf(j) = CLng("&H" & Left$(sHex, 2))
            sHex = Mid$(sHex, 3)
        Next

        'Next pair of hex characters
        'Get API function addresses
        If Subclass_InIDE Then
            'If we're running in the VB IDE
            aBuf(16) = &H90
            'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90
            'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)

            'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then
                'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)

                'VB5 perhaps
            End If

        End If

        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData
        'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)

        If nSubIdx = -1 Then
            'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1
            'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData

            'Create a new sc_aSubData element
        End If

        Subclass_Start = nSubIdx

    End If

    With sc_aSubData(nSubIdx)
        .hWnd = lng_hWnd
        'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)
        'Set our WndProc in place
        Call CopyMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
        'Copy the machine code from the static byte array to the code array in sc_aSubData
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
        'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
        'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
        'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
        'Patch the relative address of the CallWindowProc api function
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))

        'Patch the address of this object instance into the static machine code buffer
    End With

End Function

'Stop all subclassing
Private Sub Subclass_StopAll()

Dim i                                   As Long

    i = UBound(sc_aSubData())

    'Get the upper bound of the subclass data array
    Do While i >= 0

        'Iterate through each element
        With sc_aSubData(i)

            If .hWnd <> 0 Then
                'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hWnd)

                'Subclass_Stop
            End If

        End With

        i = i - 1
        'Next element
    Loop

End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)

'Parameters:
'lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)
        'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)
        'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)
        'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)
        'Release the machine code memory
        .hWnd = 0
        'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0
        'Clear the before table
        .nMsgCntA = 0
        'Clear the after table
        Erase .aMsgTblB
        'Erase the before table
        Erase .aMsgTblA

        'Erase the after table
    End With

End Sub

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

Dim TME                                 As TRACKMOUSEEVENT_STRUCT

    If bTrack Then

        With TME
            .cbSize = Len(TME)
            .dwFlags = TME_LEAVE
            .hWndTrack = lng_hWnd

        End With

        If bTrackUser32 Then
            Call TrackMouseEvent(TME)
        Else
            Call TrackMouseEventComCtl(TME)

        End If

    End If

End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.
'Worker sub for sc_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, _
                    ByRef aMsgTbl() As Long, _
                    ByRef nMsgCnt As Long, _
                    ByVal When As eMsgWhen, _
                    ByVal nAddr As Long)

Dim nEntry                              As Long

    'Message table entry index
Dim nOff1                               As Long

    'Machine code buffer offset 1
Dim nOff2                               As Long

    'Machine code buffer offset 2
    If uMsg = ALL_MESSAGES Then
        'If all messages
        nMsgCnt = ALL_MESSAGES
        'Indicates that all messages will callback
    Else

        'Else a specific message number
        Do While nEntry < nMsgCnt
            'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1

            If aMsgTbl(nEntry) = 0 Then
                'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg
                'Re-use this entry
                Exit Sub
                'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then
                'The msg is already in the table!
                Exit Sub

                'Bail
            End If

        Loop
        'Next entry
        nMsgCnt = nMsgCnt + 1
        'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg

        'Store the message number in the table
    End If

    If When = eMsgWhen.MSG_BEFORE Then
        'If before
        nOff1 = PATCH_04
        'Offset to the Before table
        nOff2 = PATCH_05
        'Offset to the Before table entry count
    Else
        'Else after
        nOff1 = PATCH_08
        'Offset to the After table
        nOff2 = PATCH_09

        'Offset to the After table entry count
    End If

    If uMsg <> ALL_MESSAGES Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))

        'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If

    Call zPatchVal(nAddr, nOff2, nMsgCnt)

    'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandle(StrPtr(sDLL)), sProc)
    Debug.Assert zAddrFunc

    'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, _
                      Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably sc_AddMsg-ing before Subclass_Start
    zIdx = UBound(sc_aSubData)

    Do While zIdx >= 0

        'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)

            If .hWnd = lng_hWnd Then

                'If the hWnd of this element is the one we're looking for
                If Not bAdd Then
                    'If we're searching not adding
                    Exit Function

                    'Found
                End If

            ElseIf .hWnd = 0 Then

                'If this an element marked for reuse.
                If bAdd Then
                    'If we're adding
                    Exit Function

                    'Re-use it
                End If

            End If

        End With

        zIdx = zIdx - 1
        'Decrement the index
    Loop

    If Not bAdd Then
        Debug.Assert False

        'hWnd not found, programmer error
    End If

    'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, _
                      ByVal nOffset As Long, _
                      ByVal nTargetAddr As Long)
    Call CopyMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)

End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call CopyMemory(ByVal nAddr + nOffset, nValue, 4)

End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True

End Function

'======================================================================================================
'   End SubClass Sections
'======================================================================================================
Public Function AddPanel(Optional ByVal sText As String, _
                         Optional ByVal uTextAlign As usbAlignEnum = usbLeft, _
                         Optional ByVal bAutoSize As Boolean = True, _
                         Optional ByVal bEditable As Boolean, _
                         Optional ByVal oIcon As StdPicture, _
                         Optional ByVal bIconState As usbStateEnum = usbEnabled, _
                         Optional ByVal bUseMaskColor As Boolean, _
                         Optional ByVal lMaskColor As OLE_COLOR = vbMagenta, _
                         Optional ByVal lForeColor As OLE_COLOR = vbButtonText, _
                         Optional ByVal oFont As StdFont, _
                         Optional ByVal sToolTipText As String, _
                         Optional ByVal lWidth As Long = 40) As Boolean

'   Handle Any Errors
    On Error GoTo Func_ErrHandler

    m_PanelCount = m_PanelCount + 1
    ReDim Preserve m_PanelItems(1 To m_PanelCount)

    With m_PanelItems(m_PanelCount)

        Select Case uTextAlign

            Case usbLeft
                .Alignment = DT_SB_LEFT

            Case usbCenter
                .Alignment = DT_SB_CENTER

            Case usbRight
                .Alignment = DT_SB_RIGHT

        End Select

        .AutoSize = bAutoSize
        .Editable = bEditable

        If Not oFont Is Nothing Then
            Set .Font = oFont
        Else

            If Not m_Font Is Nothing Then
                Set .Font = m_Font
            Else
                Set .Font = Ambient.Font

            End If

        End If

        .ForeColor = lForeColor

        If Not oIcon Is Nothing Then
            Set .Icon = oIcon

        End If

        .IconState = bIconState
        .MaskColor = lMaskColor
        .Text = sText
        .ToolTipText = sToolTipText
        .UseMaskColor = bUseMaskColor

        If lWidth > 0 Then
            .Width = lWidth
        Else
            .Width = 40

        End If

    End With

    Refresh
Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    err.Raise err.Number, "ucStatusBar.AddPanel", err.Description, err.HelpFile, err.HelpContext

    Resume Func_ErrHandlerExit:

End Function

Private Function AlphaBlend(ByVal FirstColor As Long, _
                            ByVal SecondColor As Long, _
                            ByVal AlphaValue As Long) As Long

Dim iForeColor                          As RGBQUAD
Dim iBackColor                          As RGBQUAD

    OleTranslateColorByRef FirstColor, 0, VarPtr(iForeColor)
    OleTranslateColorByRef SecondColor, 0, VarPtr(iBackColor)

    With iForeColor
        .rgbRed = (.rgbRed * AlphaValue + iBackColor.rgbRed * (255 - AlphaValue)) / 255
        .rgbGreen = (.rgbGreen * AlphaValue + iBackColor.rgbGreen * (255 - AlphaValue)) / 255
        .rgbBlue = (.rgbBlue * AlphaValue + iBackColor.rgbBlue * (255 - AlphaValue)) / 255

    End With

    CopyMemoryLong VarPtr(AlphaBlend), VarPtr(iForeColor), 4

End Function

Private Sub APILine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lColor As Long)

'Use the API LineTo for Fast Drawing
Dim PT                                  As POINT
Dim hPen As Long, hPenOld               As Long

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    lColor = TranslateColor(lColor)
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(UserControl.hDC, hPen)
    MoveToEx UserControl.hDC, X1, Y1, PT
    LineTo UserControl.hDC, X2, Y2
    SelectObject UserControl.hDC, hPenOld
    DeleteObject hPen
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.APILine", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Public Property Get BackColor() As OLE_COLOR

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    BackColor = m_BackColor
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.BackColor", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'Description: Use this color for drawing
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_BackColor = NewValue
    UserControl.BackColor = m_BackColor
    Refresh
    PropertyChanged "BackColor"
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.BackColor", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Sub BoundControl(ByVal Index As Long, _
                        Control As Object, _
                        ByVal SizeMethod As usbSizeEnum)

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If m_PanelCount < 1 Then
        Exit Sub

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    If Not Control Is Nothing Then
        m_PanelItems(Index).BoundParent = GetParent(Control.hWnd)
        Set m_PanelItems(Index).BoundObject = Control
        SetParent m_PanelItems(Index).BoundObject.hWnd, UserControl.hWnd
    Else

        '   See if the control exists, if so, then we should set the parent back
        '   and destroy the reference to it...
        If Not m_PanelItems(Index).BoundObject Is Nothing Then
            SetParent m_PanelItems(Index).BoundObject.hWnd, m_PanelItems(Index).BoundParent
            Set m_PanelItems(Index).BoundObject = Nothing

        End If

    End If

    m_PanelItems(Index).BoundSize = SizeMethod
    Refresh
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.BoundControl", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Public Sub Clear()

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Dim lpRect                          As RECT
    Dim hBrush                          As Long
    Dim lColor                          As Long

    With lpRect
        .Left = 0
        .Top = 0
        .Right = ScaleWidth
        .Bottom = ScaleHeight

    End With

    lColor = TranslateColor(m_BackColor)
    hBrush = CreateSolidBrush(lColor)
    Call FillRect(UserControl.hDC, lpRect, hBrush)
    Call DeleteObject(hBrush)
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.Clear", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Public Property Get ForeColor() As OLE_COLOR

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ForeColor = m_Forecolor
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.ForeColor", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

'Description: Use this color for drawing
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_BackColor = NewValue
    UserControl.ForeColor = m_Forecolor
    Refresh
    PropertyChanged "ForeColor"
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.ForeColor", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Get Font() As StdFont

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set Font = m_Font
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.Font", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Set Font(ByVal NewFont As StdFont)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set m_Font = NewFont
    Refresh
    PropertyChanged "Font"
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.Font", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Private Function GetPanelIndex() As Long

Dim i                                   As Long
Dim tPt                                 As POINT
Dim lpRect                              As RECT

    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    If m_PanelCount > 0 Then
        '   Get our position
        Call GetCursorPos(tPt)
        '   Convert coordinates
        Call ScreenToClient(UserControl.hWnd, tPt)

        '   Loop Over the RECTs a see if it is in
        For i = 1 To m_PanelCount
            lpRect = m_PanelItems(i).ItemRect

            If Not m_PanelItems(i).Icon Is Nothing Then
                If m_PanelItems(i).Alignment = DT_SB_LEFT Then
                    OffsetRect lpRect, -16, 0
                ElseIf m_PanelItems(i).Alignment = DT_SB_CENTER Then
                    OffsetRect lpRect, -8, 0
                ElseIf m_PanelItems(i).Alignment = DT_SB_RIGHT Then
                    InflateRect lpRect, 2, 0

                End If

            End If

            If i > 1 Then
                If (m_PanelItems(i - 1).ItemRect.Right + 10) < lpRect.Left Then
                    OffsetRect lpRect, -8, 0
                    InflateRect lpRect, 6, 0

                End If

            End If

            If PtInRect(lpRect, tPt.X, tPt.Y) Then
                GetPanelIndex = i
                Exit For

            End If

        Next

    End If

Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    err.Raise err.Number, "ucStatusBar.GetPanelIndex", err.Description, err.HelpFile, err.HelpContext

    Resume Func_ErrHandlerExit:

End Function

Private Function GetThemeInfo() As String

Dim lPtrColorName As Long
Dim lPtrThemeFile As Long
Dim hTheme As Long
Dim sColorName As String
Dim sThemeFile As String

    If m_bIsWinXpOrLater Then
        hTheme = OpenThemeData(hWnd, StrPtr("Button"))
        If hTheme Then
            ReDim bThemeFile(0 To 260 * 2) As Byte
            lPtrThemeFile = VarPtr(bThemeFile(0))
            ReDim bColorName(0 To 260 * 2) As Byte
            lPtrColorName = VarPtr(bColorName(0))
            
            If GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0) <> &H0 Then
                GetThemeInfo = "UxTheme_Error"
                Exit Function
            Else
                sThemeFile = TrimNull(bThemeFile)
                sColorName = TrimNull(bColorName)
            End If
            
            CloseThemeData hTheme
        End If
    End If
    
    If LenB(Trim$(sColorName)) = 0 Then sColorName = "None"
    GetThemeInfo = sColorName

End Function

Private Sub GrayBlt(ByVal hDstDC As Long, _
                    ByVal hSrcDC As Long, _
                    ByVal nWidth As Long, _
                    ByVal nHeight As Long)

Dim MakePal                             As Long
Dim DIBInf                              As BITMAPINFO8
Dim gsDIB                               As Long
Dim hTmpDC                              As Long
Dim OldDIB                              As Long

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    hTmpDC = CreateCompatibleDC(hSrcDC)

    With DIBInf
        With .bmiHeader
            ' Same size as picture
            .biWidth = nWidth
            .biHeight = nHeight
            .biBitCount = 8
            .biPlanes = 1
            .biClrUsed = 256
            .biClrImportant = 256
            .biSize = Len(DIBInf.bmiHeader)

        End With

        ' Palette is Greyscale
        For MakePal = 0 To 255

            With .bmiColors(MakePal)
                .rgbRed = MakePal
                .rgbGreen = MakePal
                .rgbBlue = MakePal

            End With

        Next

    End With

    gsDIB = CreateDIBSection8(hTmpDC, DIBInf, 0, ByVal 0&, 0, 0)

    If (hTmpDC) Then
        ' Validate and select DIB
        OldDIB = SelectObject(hTmpDC, gsDIB)
        ' Draw original picture to the greyscale DIB
        BitBlt hTmpDC, 0, 0, DIBInf.bmiHeader.biWidth, DIBInf.bmiHeader.biHeight, hSrcDC, 0, 0, vbSrcCopy
        ' Draw the greyscale image back to the hDC
        BitBlt hDstDC, 0, 0, DIBInf.bmiHeader.biWidth, DIBInf.bmiHeader.biHeight, hTmpDC, 0, 0, vbSrcCopy
        ' Clean up DIB
        SelectObject hTmpDC, OldDIB
        DeleteObject gsDIB
        DeleteObject hTmpDC

    End If

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.GrayBlt", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Public Property Get GripShape() As usbGripEnum

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    GripShape = m_GripShape
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.GripShape", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Let GripShape(lShape As usbGripEnum)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    '   Check to see if this changed, otherwise we get an
    '   "Out of Stack Space" error with recursive changes...
    If lShape <> m_GripShape Then
        m_GripShape = lShape
        Refresh
        PropertyChanged "GripShape"

    End If

Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.GripShape", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Private Sub PaintGradients()

Dim i                                   As Long
Dim Y1                                  As Long
Dim BtnFace                             As Long
Dim lColor                              As Long

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With UserControl

        If (m_iTheme = usbClassic) Then
            '   Clear the control to start using the
            '   optimized repaint method instead Cls to avoid flicker
            Clear
        Else
            '   Get the BackColor and Offset it by 2 Units
            BtnFace = ShiftColor(.BackColor, -&H1)
            '   Clear the control to start using the
            '   optimized repaint method instead Cls to avoid flicker
            Clear

            '   Draw the Smooth Gradient across the whole control
            For i = 0 To ScaleHeight
                Y1 = i
                APILine 0, Y1, .ScaleWidth, Y1, AlphaBlend(&HFFFFFF, BtnFace, (i / ScaleHeight) * 48)
            Next

            '   Draw The Top Lines
            Select Case m_iTheme

                Case usbBlue
                    lColor = AlphaBlend(ShiftColor(BtnFace, -&H40), &HB99D7F, 128)

                Case usbHomeStead
                    lColor = AlphaBlend(ShiftColor(BtnFace, -&H40), &H69A18B, 128)

                Case usbMetallic
                    lColor = AlphaBlend(ShiftColor(BtnFace, -&H40), &H947C7C, 128)

                Case Else
                    lColor = ShiftColor(BtnFace, -&H50)

            End Select

            APILine 0, 0, .ScaleWidth, 0, &HFFFFFF
            'AlphaBlend(ShiftColor(BtnFace, -&H8), &HFFFFFF, 128)
            APILine 0, 1, .ScaleWidth, 1, lColor
            '   Draw the Top Gradient
            APILine 0, 2, .ScaleWidth, 2, ShiftColor(BtnFace, -&H25)
            APILine 0, 3, .ScaleWidth, 3, ShiftColor(BtnFace, -&H9)

            '   Draw the Bottom Gradient
            For i = 0 To 5
                Y1 = .ScaleHeight - 5 + i
                APILine 0, Y1, .ScaleWidth, Y1, ShiftColor(BtnFace, -&H1 * ((((i / 3) * 100) * .ScaleHeight) / 100))
            Next

        End If

    End With

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PaintGradients", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub PaintGrip()

Dim AdjWidth                            As Long
Dim AdjHeight                           As Long

    '   Custom reoutine, to paint/repaint the shapes on the
    '   screen to represent the Grip Style selected...
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With UserControl
        AdjWidth = (.ScaleWidth - 15)
        AdjHeight = (.ScaleHeight - 16)

        '   See if this is XP, if so then paint the correct Resize Button
        If (m_GripShape = usbSquare) And (m_iTheme <> usbClassic) Then
            '   Paint the Shadows first....
            .ForeColor = vbWhite
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 4, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 4, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 5, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 5, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 9, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 9, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 13, .ForeColor
            SetPixelV .hDC, AdjWidth + 13, AdjHeight + 13, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 9, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 9, .ForeColor
            SetPixelV .hDC, AdjWidth + 9, AdjHeight + 9, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 9, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 13, .ForeColor
            SetPixelV .hDC, AdjWidth + 9, AdjHeight + 13, .ForeColor
            SetPixelV .hDC, AdjWidth + 4, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 5, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 4, AdjHeight + 13, .ForeColor
            SetPixelV .hDC, AdjWidth + 5, AdjHeight + 13, .ForeColor
            '   Shift the Color to be a Blend of the BackColor and Medium Grey
            .ForeColor = AlphaBlend(&H909090, .BackColor, 128)
            '   Paint the Grips Next....
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 3, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 3, .ForeColor
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 4, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 4, .ForeColor
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 7, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 7, .ForeColor
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 12, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 11, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 7, AdjHeight + 7, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 7, .ForeColor
            SetPixelV .hDC, AdjWidth + 7, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 8, .ForeColor
            SetPixelV .hDC, AdjWidth + 7, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 7, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 8, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 3, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 4, AdjHeight + 11, .ForeColor
            SetPixelV .hDC, AdjWidth + 3, AdjHeight + 12, .ForeColor
            SetPixelV .hDC, AdjWidth + 4, AdjHeight + 12, .ForeColor
        ElseIf (m_GripShape = usbBars) And (m_iTheme = usbClassic) Then
            '   Draw the White Highlight Lines First in groups of two
            .ForeColor = vbWhite
            APILine AdjWidth + 12, AdjHeight + 13, AdjWidth + 14, AdjHeight + 11, .ForeColor
            APILine AdjWidth + 9, AdjHeight + 13, AdjWidth + 14, AdjHeight + 8, .ForeColor
            APILine AdjWidth + 6, AdjHeight + 13, AdjWidth + 14, AdjHeight + 5, .ForeColor
            APILine AdjWidth + 3, AdjHeight + 13, AdjWidth + 14, AdjHeight + 2, .ForeColor
            '   Now Draw the Lowlight Lines in groups of two
            .ForeColor = AlphaBlend(vbWhite, ShiftColor(.BackColor, -&H70), 128)
            APILine AdjWidth + 13, AdjHeight + 14, AdjWidth + 14, AdjHeight + 13, .ForeColor
            APILine AdjWidth + 12, AdjHeight + 14, AdjWidth + 14, AdjHeight + 12, .ForeColor
            APILine AdjWidth + 10, AdjHeight + 14, AdjWidth + 14, AdjHeight + 10, .ForeColor
            APILine AdjWidth + 9, AdjHeight + 14, AdjWidth + 14, AdjHeight + 9, .ForeColor
            APILine AdjWidth + 7, AdjHeight + 14, AdjWidth + 14, AdjHeight + 7, .ForeColor
            APILine AdjWidth + 6, AdjHeight + 14, AdjWidth + 14, AdjHeight + 6, .ForeColor
            APILine AdjWidth + 4, AdjHeight + 14, AdjWidth + 14, AdjHeight + 4, .ForeColor
            APILine AdjWidth + 3, AdjHeight + 14, AdjWidth + 14, AdjHeight + 3, .ForeColor

        End If

    End With

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PaintGrip", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub PaintPanels()

Dim i                                   As Long
Dim lX                                  As Long
Dim lForeColor                          As Long
Dim lIconOffset                         As Long
Dim lGripSize                           As Long
Dim bMinWidth                           As Boolean

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    lForeColor = UserControl.ForeColor
    lIconOffset = 18

    If (m_iTheme = usbClassic) Then
        lGripSize = 16
    Else
        lGripSize = 18

    End If

    For i = 1 To PanelCount

        With m_PanelItems(i)
            '   Set the Individual ForeColor & Font
            UserControl.ForeColor = .ForeColor
            Set UserControl.Font = .Font

            '   Autosize the Text + Icon?
            If .AutoSize Then
                '   Set the Left & Top
                .ItemRect.Left = lX
                .ItemRect.Top = 5

                '   Do we have a valid Icon?
                If .Icon Is Nothing Then
                    '   Compute the Distance we need to Extend the Rect
                    .ItemRect.Right = lX + TextWidth(.Text) + 8
                Else
                    '   Compute the Distance we need to Extend the Rect + Icon Distance
                    .ItemRect.Right = lX + TextWidth(.Text) + 8 + lIconOffset

                End If

                '   Set the Bottom of the Rect
                .ItemRect.Bottom = ScaleHeight - 5

                '   Use a default for blank text
                If LenB(.Text) > 0 Then
                    lX = .ItemRect.Right
                Else
                    lX = lX + 20

                End If

                '   Check to see if the control is smaller then the
                '   right most separator, if so correct it
                If lX >= (ScaleWidth - lGripSize) Then
                    '   Yep, so make the Rect scaller to match
                    .ItemRect.Right = (ScaleWidth - lGripSize)
                    lX = .ItemRect.Right

                End If

            Else
                '   Set the Left & Top
                .ItemRect.Left = lX
                .ItemRect.Top = 5

                '   Do we have a valid Icon?
                If .Icon Is Nothing Then
                    '   Compute the Distance we need to Extend the Rect
                    .ItemRect.Right = lX + .Width
                Else
                    '   Compute the Distance we need to Extend the Rect + Icon Distance
                    .ItemRect.Right = lX + .Width + lIconOffset

                End If

                '   Set the Bottom of the Rect
                .ItemRect.Bottom = ScaleHeight - 5
                lX = .ItemRect.Right

                '   Check to see if the control is smaller then the
                '   right most separator, if so correct it
                If lX >= (ScaleWidth - lGripSize) Then
                    '   Yep, so make the Rect scaller to match
                    .ItemRect.Right = (ScaleWidth - lGripSize)
                    lX = .ItemRect.Right

                End If

            End If

            '   Now draw the Theme Based Borders....
            If (m_iTheme = usbClassic) Then
                '   Draw the Panels as Sunken Boxes as per 9x LnF
                InflateRect .ItemRect, 0, 3
                DrawEdge UserControl.hDC, .ItemRect, EDGE_SUNKEN, BF_RECT
                InflateRect .ItemRect, -5, -3
            Else
                '   Draw the Lines for the Dividors as per XP LnF
                APILine lX, .ItemRect.Top, lX, .ItemRect.Bottom, AlphaBlend(&H909090, m_BackColor, 128)
                APILine lX + 1, .ItemRect.Top, lX + 1, .ItemRect.Bottom, vbWhite
                '   Decrease the RECT by 4
                InflateRect .ItemRect, -4, 0

            End If

            '   Does this have a bound object?
            If .BoundObject Is Nothing Then

                '   Do we have a Picture?
                If Not .Icon Is Nothing Then

                    '   Adjust the Initial Items RECT to line up correctly
                    If i = 1 Then
                        OffsetRect .ItemRect, -2, 0

                    End If

                    '   See if the size of the StatusBar is too small for an Icon + Padding
                    If (.ItemRect.Left + lIconOffset) <= (ScaleWidth - lGripSize) Then
                        '   Yep, so paint it centered vertically
                        TransBltEx UserControl.hDC, .ItemRect.Left, ScaleHeight \ 2 - 8, 16, 16, .Icon, 0, 0, BackColor, IIf(.IconState = usbEnabled, False, True)
                        '   Now offset th RECT so the text starts in the corect position
                        OffsetRect .ItemRect, lIconOffset \ 2, 0
                        InflateRect .ItemRect, -lIconOffset \ 2, 0

                        '   Perform adjustments as needed depending on Aligment
                        If .Alignment = DT_SB_LEFT Then

                            '   Adjust the Right most extent if the item is smaller
                            '   than the RECT....
                            If lX >= (ScaleWidth - lGripSize) Then
                                '   Yep, so make the Rect scaller to match
                                .ItemRect.Right = (ScaleWidth - lGripSize)

                            End If

                        ElseIf .Alignment = DT_SB_CENTER Then

                            '   Adjust the Right most extent if the item is smaller
                            '   than the RECT....
                            If lX >= (ScaleWidth - lGripSize) Then
                                '   Yep, so make the Rect scaller to match
                                OffsetRect .ItemRect, 0, 0
                                .ItemRect.Right = (ScaleWidth - lGripSize) - 2

                            End If

                        ElseIf .Alignment = DT_SB_RIGHT Then

                            '   Adjust the Right most extent if the item is smaller
                            '   than the RECT....
                            If lX >= (ScaleWidth - lGripSize) Then
                                '   Yep, so make the Rect scaller to match
                                OffsetRect .ItemRect, lIconOffset, 0
                                InflateRect .ItemRect, lIconOffset, 0
                                .ItemRect.Right = (ScaleWidth - lGripSize) - 2

                            End If

                        End If

                    End If

                    '   See if the size of the StatusBar is too small for an Icon + Padding
                    '   if so then we don't want to paint the text where the icon was located
                    If (.Alignment = DT_SB_LEFT) Or (.Alignment = DT_SB_RIGHT) Then

                        '   If there is enough room, print the text
                        If ((.ItemRect.Left + lIconOffset) <= (ScaleWidth - lGripSize)) Or ((.ItemRect.Right - .ItemRect.Left) > 16) Then
                            DrawText UserControl.hDC, .Text, -1, .ItemRect, .Alignment

                        End If

                    Else

                        '   If there is enough room, print the text
                        If (.ItemRect.Left + lIconOffset \ 2) <= (ScaleWidth - lGripSize) Or ((.ItemRect.Right - .ItemRect.Left) > 16) Then
                            DrawText UserControl.hDC, .Text, -1, .ItemRect, .Alignment

                        End If

                    End If

                Else

                    '   If there is enough room, print the text
                    If (.ItemRect.Left + 2) <= (ScaleWidth - lGripSize) Then
                        DrawText UserControl.hDC, .Text, -1, .ItemRect, .Alignment

                    End If

                End If

            Else

                '   Set the Bound Object onto the Control
                '
                '   Handle errors quietly in this section as we are late bound
                '   so it is hard to predict if all controls will support certain
                '   object interfaces....
                On Error Resume Next

                '   Only deal with real controls
                If Not .BoundObject Is Nothing Then

                    '   Is this going to be resized or not....
                    If .BoundSize = usbNoSize Then

                        '   Keep the Width, but set the Left, Top and Height
                        With .BoundObject
                            .Left = m_PanelItems(i).ItemRect.Left * Screen.TwipsPerPixelX
                            .Top = m_PanelItems(i).ItemRect.Top * Screen.TwipsPerPixelY
                            .Height = 16 * Screen.TwipsPerPixelY
                            '   Under development....;-)
                            '   Should be hidden if too small to fit the control..
                            'If (.Width <= ((m_PanelItems(i).ItemRect.Right - m_PanelItems(i).ItemRect.Left)) * Screen.TwipsPerPixelX) Then
                            '    .Visible = False
                            'Else
                            '    .Visible = True
                            'End If
                            .ZOrder 0

                        End With

                    Else

                        With .BoundObject

                            '   Resize all properties to make it fit
                            If m_iTheme <> usbClassic Then
                                .Left = (m_PanelItems(i).ItemRect.Left) * Screen.TwipsPerPixelX
                                .Width = ((m_PanelItems(i).ItemRect.Right - m_PanelItems(i).ItemRect.Left)) * Screen.TwipsPerPixelX

                                '   See if we were avel to resize the controls width, if not
                                '   then the control might have a minimum width (i.e. ComboBox)
                                '   so we can simply use this as an indicator to hide the control...
                                If (.Width <> (((m_PanelItems(i).ItemRect.Right - m_PanelItems(i).ItemRect.Left)) * Screen.TwipsPerPixelX)) Then
                                    bMinWidth = True
                                Else
                                    bMinWidth = False

                                End If

                            Else
                                .Left = (m_PanelItems(i).ItemRect.Left - 4) * Screen.TwipsPerPixelX
                                .Width = ((m_PanelItems(i).ItemRect.Right - m_PanelItems(i).ItemRect.Left) + 9) * Screen.TwipsPerPixelX

                                '   See if we were avel to resize the controls width, if not
                                '   then the control might have a minimum width (i.e. ComboBox)
                                '   so we can simply use this as an indicator to hide the control...
                                If (.Width <> (((m_PanelItems(i).ItemRect.Right - m_PanelItems(i).ItemRect.Left) + 9) * Screen.TwipsPerPixelX)) Then
                                    bMinWidth = True
                                Else
                                    bMinWidth = False

                                End If

                            End If

                            .Height = 16 * Screen.TwipsPerPixelY
                            .Top = Height \ 2 - .Height \ 2

                            If (.Width <= 30) Or (bMinWidth = True) Then
                                .Visible = False
                            Else
                                .Visible = True

                            End If

                            .ZOrder 0

                        End With

                    End If

                End If

                '   Turn the normal Error handing back on....
                On Error GoTo 0

            End If

        End With

    Next
    '   Set the ForeColor back...
    UserControl.ForeColor = lForeColor
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PaintPanels", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Public Property Get PanelCount() As Long

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_PanelCount = UBoundEx(m_PanelItems)
    PanelCount = m_PanelCount
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelCount", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Get PanelAlignment(ByVal Index As Long) As usbAlignEnum

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount

    Select Case m_PanelItems(Index).Alignment

        Case DT_SB_LEFT
            PanelAlignment = usbLeft

        Case DT_SB_CENTER
            PanelAlignment = usbCenter

        Case DT_SB_RIGHT
            PanelAlignment = usbRight

    End Select

Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelAlignment", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Let PanelAlignment(ByVal Index As Long, ByVal NewValue As usbAlignEnum)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then
        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount

    Select Case NewValue

        Case usbLeft
            m_PanelItems(Index).Alignment = DT_SB_LEFT

        Case usbCenter
            m_PanelItems(Index).Alignment = DT_SB_CENTER

        Case usbRight
            m_PanelItems(Index).Alignment = DT_SB_RIGHT

    End Select

    Refresh
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelAlignment", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Get PanelAutoSize(ByVal Index As Long) As Boolean

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelAutoSize = m_PanelItems(Index).AutoSize
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelAutoSize", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Let PanelAutoSize(ByVal Index As Long, ByVal NewValue As Boolean)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then
        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).AutoSize = NewValue
    Refresh
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelAutoSize", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Get PanelEditable(ByVal Index As Long) As Boolean

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelEditable = m_PanelItems(Index).Editable
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelEditable", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Let PanelEditable(ByVal Index As Long, ByVal NewValue As Boolean)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then
        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).Editable = NewValue
    Refresh
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelEditable", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Get PanelForeColor(ByVal Index As Long) As OLE_COLOR

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelForeColor = m_PanelItems(Index).ForeColor
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelForeColor", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Let PanelForeColor(ByVal Index As Long, ByVal NewItem As OLE_COLOR)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then
        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).ForeColor = NewItem
    Refresh
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelForeColor", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Get PanelFont(ByVal Index As Long) As StdFont

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    Set PanelFont = m_PanelItems(Index).Font
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelFont", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Let PanelFont(ByVal Index As Long, ByVal NewItem As StdFont)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then
        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    Set m_PanelItems(Index).Font = NewItem
    Refresh
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelFont", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Get PanelIcon(ByVal Index As Long) As StdPicture

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    Set PanelIcon = m_PanelItems(Index).Icon
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelIcon", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Set PanelIcon(ByVal Index As Long, ByVal NewItem As StdPicture)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then
        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    Set m_PanelItems(Index).Icon = NewItem
    Refresh
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelIcon", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Get PanelText(ByVal Index As Long) As String

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelText = m_PanelItems(Index).Text
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelText", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Let PanelText(ByVal Index As Long, ByVal NewItem As String)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then
        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).Text = NewItem
    Refresh
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelText", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Get PanelToolTipText(ByVal Index As Long) As String

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelToolTipText = m_PanelItems(Index).ToolTipText
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelToolTipText", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Let PanelToolTipText(ByVal Index As Long, NewValue As String)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then
        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).ToolTipText = NewValue
    Refresh
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelToolTipText", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Get PanelWidth(ByVal Index As Long) As Long

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    PanelWidth = m_PanelItems(Index).Width
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelWidth", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Let PanelWidth(ByVal Index As Long, ByVal NewItem As Long)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If m_PanelCount < 1 Then
        Exit Property

    End If

    If Index < 1 Then Index = 1
    If Index > m_PanelCount Then Index = m_PanelCount
    m_PanelItems(Index).Width = NewItem
    Refresh
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PanelWidth", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Private Function PtInRect(ByRef lpRect As RECT, X As Long, Y As Long) As Boolean

'   This is a replacemnt for the PtInRect API call which seems to always
'   return 0 depite the X & Y Points being in the RECT...
'   Handle Any Errors
    On Error GoTo Func_ErrHandler

    If (X >= lpRect.Left) Then
        If (X <= lpRect.Right) Then
            If (Y >= lpRect.Top) Then
                If (Y <= lpRect.Bottom) Then
                    PtInRect = True
                End If
            End If
        End If
    End If

Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    err.Raise err.Number, "ucStatusBar.PtInRect", err.Description, err.HelpFile, err.HelpContext

    Resume Func_ErrHandlerExit:

End Function

Public Sub Refresh()

Dim AutoTheme                           As String

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Select Case m_Theme

        Case [usbAuto]
            AutoTheme = GetThemeInfo

            Select Case AutoTheme

                Case "None", "UxTheme_Error"
                    m_iTheme = usbClassic

                    If m_GripShape <> usbNone Then
                        m_GripShape = usbBars

                    End If

                Case "NormalColor"
                    m_iTheme = usbBlue

                    If m_GripShape <> usbNone Then
                        m_GripShape = usbSquare

                    End If

                Case "HomeStead"
                    m_iTheme = usbHomeStead

                    If m_GripShape <> usbNone Then
                        m_GripShape = usbSquare

                    End If

                Case "Metallic"
                    m_iTheme = usbMetallic

                    If m_GripShape <> usbNone Then
                        m_GripShape = usbSquare

                    End If

                Case Else
                    m_iTheme = usbBlue

                    If m_GripShape <> usbNone Then
                        m_GripShape = usbSquare

                    End If

            End Select

        Case [usbClassic]
            m_iTheme = usbClassic

            If m_GripShape <> usbNone Then
                m_GripShape = usbBars

            End If

        Case [usbBlue]
            m_iTheme = usbBlue

            If m_GripShape <> usbNone Then
                m_GripShape = usbSquare

            End If

        Case [usbHomeStead]
            m_iTheme = usbHomeStead

            If m_GripShape <> usbNone Then
                m_GripShape = usbSquare

            End If

        Case [usbMetallic]
            m_iTheme = usbMetallic

            If m_GripShape <> usbNone Then
                m_GripShape = usbSquare

            End If

        Case Else
            m_iTheme = usbBlue

            If m_GripShape <> usbNone Then
                m_GripShape = usbSquare

            End If

    End Select

    '   Paint the Gradient for the whole control
    PaintGradients
    '   Now Paint the Grip according to style
    PaintGrip
    '   Paint the Divisions which represent the panels
    PaintPanels

    '   Only refresh if in the IDE (Otherwise it will Flicker!!)
    If Not Ambient.UserMode Then
        AutoRedraw = False
    Else
        AutoRedraw = True
        '   Refresh the Window
        UserControl.Refresh

    End If

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.Refresh", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long) As Long

Dim lR                                  As Long
Dim lg                                  As Long
Dim lb                                  As Long

    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    Color = TranslateColor(Color)
    lR = (Color And &HFF) + Value
    lg = ((Color \ &H100) Mod &H100) + Value
    lb = ((Color \ &H10000) Mod &H100)
    lb = lb + ((lb * Value) \ &HC0)

    If Value > 0 Then
        If lR > 255 Then lR = 255
        If lg > 255 Then lg = 255
        If lb > 255 Then lb = 255
    ElseIf Value < 0 Then

        If lR < 0 Then lR = 0
        If lg < 0 Then lg = 0
        If lb < 0 Then lb = 0

    End If

    ShiftColor = lR + 256& * lg + 65536 * lb
Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    err.Raise err.Number, "ucStatusBar.ShiftColor", err.Description, err.HelpFile, err.HelpContext

    Resume Func_ErrHandlerExit:

End Function

Public Property Get Sizable() As Boolean

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Sizable = m_Sizable
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.Sizable", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Let Sizable(ByVal NewValue As Boolean)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_Sizable = NewValue

    If m_Sizable Then
        If IsWinXPOrLater Then
            m_GripShape = usbSquare
        Else
            m_GripShape = usbBars

        End If

    Else
        m_GripShape = usbNone

    End If

    Refresh
    PropertyChanged "Sizable"
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.Sizable", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Get Theme() As usbThemeEnum

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Theme = m_Theme
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.Theme", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Property Let Theme(ByVal NewValue As usbThemeEnum)

'   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_Theme = NewValue
    Refresh
    PropertyChanged "Theme"
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    err.Raise err.Number, "ucStatusBar.Theme", err.Description, err.HelpFile, err.HelpContext

    Resume Prop_ErrHandlerExit:

End Property

Public Sub TransBltEx(ByVal hDestDC As Long, _
                      ByVal X As Long, _
                      ByVal Y As Long, _
                      ByVal nWidth As Long, _
                      ByVal nHeight As Long, _
                      ByVal hSrcImg As StdPicture, _
                      ByVal XSrc As Long, _
                      ByVal YSrc As Long, _
                      ByVal TransColor As Long, _
                      ByVal Disabled As Boolean)

'
'   32-Bit Transparent BitBlt Function
'   Written by Karl E. Peterson, 9/20/96.
'   Portions borrowed and modified from KB.
'   Other portions modified following input from users. <g>
'
'   Modified by Paul R. Territo, Ph.D 02Apr07 to allow
'   passing of a StdPicture object and populating a private
'   hSrcDC instead of the original method which passed the hScrDC
'
'   Modified by Paul R. Territo, Ph.D 11Apr07 to allow for GrayScaling of
'   the passed image via the GrayBlt method implemented in the UserControl.
'
'Parameters ************************************************************
'   hDestDC:     Destination device context
'   x, y:        Upper-left destination coordinates (pixels)
'   nWidth:      Width of destination
'   nHeight:     Height of destination
'   hSrcImg:     Source StdPicture Object
'   xSrc, ySrc:  Upper-left source coordinates (pixels)
'   TransColor:  RGB value for transparent pixels, typically &HC0C0C0.
'***********************************************************************
'
Dim OrigColor                           As Long

    ' Holds original background color
Dim OrigMode                            As Long

    ' Holds original background drawing mode
Dim hSrcDC                              As Long
Dim tObj                                As Long
Dim hBrush                              As Long

    'Handle to the Brush we are using for MaskColor
Dim hTmp                                As Long

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Create a DC which is compatible with the destination DC
    hSrcDC = CreateCompatibleDC(hDestDC)

    '   Check if it is an Icon or a Bitmap
    If hSrcImg.Type = vbPicTypeBitmap Then
        '   Bitmap, so simply Select it into the DC
        tObj = SelectObject(hSrcDC, hSrcImg.Handle)
        DeleteObject tObj
    Else
        '   This is an Icon, so we need to Draw this into the DC
        '   at the new size....we are using the TransColor here as the
        '   MaskColor so pass the handled to the brush
        hTmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
        tObj = SelectObject(hSrcDC, hTmp)
        hBrush = CreateSolidBrush(TransColor)
        'MaskColor)
        DrawIconEx hSrcDC, 0, 0, hSrcImg.Handle, nWidth, nHeight, 0, hBrush, &H1 Or &H2
        '   Clean up the brush
        DeleteObject hBrush
        DeleteObject hTmp
        DeleteObject tObj

    End If

    If (GetDeviceCaps(hDestDC, CAPS1) And C1_TRANSPARENT) Then
        '
        ' Some NT machines support this *super* simple method!
        ' Save original settings, Blt, restore settings.
        '
        OrigMode = SetBkMode(hDestDC, NEWTRANSPARENT)
        OrigColor = SetBkColor(hDestDC, TransColor)

        '
        '   Check to see if this is a GreyScale Image, if so then GrayBlt it
        '   to the DC it is located on...
        '
        If Disabled Then
            GrayBlt hSrcDC, hSrcDC, nWidth, nHeight

        End If

        Call BitBlt(hDestDC, X, Y, nWidth, nHeight, hSrcDC, XSrc, YSrc, SRCCOPY)
        Call SetBkColor(hDestDC, OrigColor)
        Call SetBkMode(hDestDC, OrigMode)
    Else

        Dim saveDC                      As Long

        ' Backup copy of source bitmap
        Dim maskDC                      As Long

        ' Mask bitmap (monochrome)
        Dim invDC                       As Long

        ' Inverse of mask bitmap (monochrome)
        Dim resultDC                    As Long

        ' Combination of source bitmap & background
        Dim hSaveBmp                    As Long

        ' Bitmap stores backup copy of source bitmap
        Dim hMaskBmp                    As Long

        ' Bitmap stores mask (monochrome)
        Dim hInvBmp                     As Long

        ' Bitmap holds inverse of mask (monochrome)
        Dim hResultBmp                  As Long

        ' Bitmap combination of source & background
        Dim hSavePrevBmp                As Long

        ' Holds previous bitmap in saved DC
        Dim hMaskPrevBmp                As Long

        ' Holds previous bitmap in the mask DC
        Dim hInvPrevBmp                 As Long

        ' Holds previous bitmap in inverted mask DC
        Dim hDestPrevBmp                As Long

        ' Holds previous bitmap in destination DC
        '
        ' Create DCs to hold various stages of transformation.
        '
        saveDC = CreateCompatibleDC(hDestDC)
        maskDC = CreateCompatibleDC(hDestDC)
        invDC = CreateCompatibleDC(hDestDC)
        resultDC = CreateCompatibleDC(hDestDC)
        '
        ' Create monochrome bitmaps for the mask-related bitmaps.
        '
        hMaskBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
        hInvBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
        '
        ' Create color bitmaps for final result & stored copy of source.
        '
        hResultBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
        hSaveBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
        '
        ' Select bitmaps into DCs.
        '
        hSavePrevBmp = SelectObject(saveDC, hSaveBmp)
        hMaskPrevBmp = SelectObject(maskDC, hMaskBmp)
        hInvPrevBmp = SelectObject(invDC, hInvBmp)
        hDestPrevBmp = SelectObject(resultDC, hResultBmp)
        '
        ' Create mask: set background color of source to transparent color.
        '
        OrigColor = SetBkColor(hSrcDC, TransColor)
        Call BitBlt(maskDC, 0, 0, nWidth, nHeight, hSrcDC, XSrc, YSrc, vbSrcCopy)
        TransColor = SetBkColor(hSrcDC, OrigColor)
        '
        ' Create inverse of mask to AND w/ source & combine w/ background.
        '
        Call BitBlt(invDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbNotSrcCopy)
        '
        ' Copy background bitmap to result.
        '
        Call BitBlt(resultDC, 0, 0, nWidth, nHeight, hDestDC, X, Y, vbSrcCopy)
        '
        ' AND mask bitmap w/ result DC to punch hole in the background by
        ' painting black area for non-transparent portion of source bitmap.
        '
        Call BitBlt(resultDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbSrcAnd)

        '
        '   Check to see if this is a GreyScale Image, if so then GrayBlt it
        '   to the DC it is located on...
        '
        If Disabled Then
            GrayBlt hSrcDC, hSrcDC, nWidth, nHeight

        End If

        '
        ' get overlapper
        '
        Call BitBlt(saveDC, 0, 0, nWidth, nHeight, hSrcDC, XSrc, YSrc, vbSrcCopy)
        '
        ' AND with inverse monochrome mask
        '
        Call BitBlt(saveDC, 0, 0, nWidth, nHeight, invDC, 0, 0, vbSrcAnd)
        '
        ' XOR these two
        '
        Call BitBlt(resultDC, 0, 0, nWidth, nHeight, saveDC, 0, 0, vbSrcInvert)
        '
        ' Display transparent bitmap on background.
        '
        Call BitBlt(hDestDC, X, Y, nWidth, nHeight, resultDC, 0, 0, vbSrcCopy)
        '
        ' Select original objects back.
        '
        Call SelectObject(saveDC, hSavePrevBmp)
        Call SelectObject(resultDC, hDestPrevBmp)
        Call SelectObject(maskDC, hMaskPrevBmp)
        Call SelectObject(invDC, hInvPrevBmp)
        '
        ' Deallocate system resources.
        '
        Call DeleteObject(hSaveBmp)
        Call DeleteObject(hMaskBmp)
        Call DeleteObject(hInvBmp)
        Call DeleteObject(hResultBmp)
        Call DeleteDC(saveDC)
        Call DeleteDC(invDC)
        Call DeleteDC(maskDC)
        Call DeleteDC(resultDC)

    End If

    Call DeleteDC(hSrcDC)
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.TransBltEx", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Function TranslateColor(ByVal lColor As Long) As Long

'   Handle Any Errors
    On Error GoTo Func_ErrHandler

    If OleTranslateColor(lColor, 0, TranslateColor) Then
        TranslateColor = -1

    End If

Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    err.Raise err.Number, "ucStatusBar.TranslateColor", err.Description, err.HelpFile, err.HelpContext

    Resume Func_ErrHandlerExit:

End Function

Private Function UBoundEx(uArr() As PanelItem) As Long

    On Error Resume Next

    UBoundEx = UBound(uArr, 1)

End Function

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Select Case KeyCode

        Case vbKeyEscape

            If txtEdit.Visible = True Then
                txtEdit.Visible = False

            End If

        Case vbKeyReturn

            If txtEdit.Visible = True Then
                m_PanelItems(m_ActivePanel).Text = txtEdit.Text
                txtEdit.Visible = False
                Refresh

            End If

    End Select

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.txtEdit_KeyUp", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub txtEdit_LostFocus()

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If txtEdit.Visible = True Then
        m_PanelItems(m_ActivePanel).Text = txtEdit.Text
        txtEdit.Visible = False

    End If

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.txtEdit_LostFocus", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_Click()

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If txtEdit.Visible = True Then
        txtEdit.Visible = False

    End If

    RaiseEvent Click
    RaiseEvent PanelClick(GetPanelIndex())
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_Click", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_DblClick()

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    m_ActivePanel = GetPanelIndex()

    If m_ActivePanel > 0 Then

        With m_PanelItems(m_ActivePanel)

            If .Editable Then
                If m_iTheme <> usbClassic Then
                    txtEdit.BackColor = m_BackColor
                    txtEdit.Left = .ItemRect.Left
                    txtEdit.Height = 16
                    txtEdit.Top = ScaleHeight \ 2 - txtEdit.Height \ 2
                    txtEdit.Width = ((.ItemRect.Right - .ItemRect.Left))
                Else
                    txtEdit.BackColor = m_BackColor

                    If Not .Icon Is Nothing Then
                        txtEdit.Left = .ItemRect.Left - 1
                    Else
                        txtEdit.Left = .ItemRect.Left - 4

                    End If

                    txtEdit.Height = 16 - 12
                    txtEdit.Top = (ScaleHeight \ 2 - txtEdit.Height \ 2) - 1
                    txtEdit.Width = ((.ItemRect.Right - .ItemRect.Left)) + 8

                End If

                txtEdit.Text = .Text
                txtEdit.SelStart = 0
                txtEdit.SelLength = Len(.Text)
                txtEdit.Visible = True
                txtEdit.ZOrder 0
                txtEdit.SetFocus

            End If

        End With

    End If

    RaiseEvent DblClick
    RaiseEvent PanelDblClick(m_ActivePanel)
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_DblClick", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_Initialize()
    m_bIsWinXpOrLater = IsWinXPOrLater
End Sub

Private Sub UserControl_InitProperties()

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    m_BackColor = vbButtonFace
    m_Forecolor = vbButtonText
    Set m_Font = UserControl.Font
    m_GripShape = usbSquare
    m_Sizable = True
    m_Theme = usbAuto
    m_iTheme = m_Theme
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_InitProperties", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent KeyDown(KeyCode, Shift)
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_KeyDown", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent KeyPress(KeyAscii)
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_KeyPress", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent KeyUp(KeyCode, Shift)
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_KeyUp", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_LostFocus()

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If txtEdit.Visible = True Then
        m_PanelItems(m_ActivePanel).Text = txtEdit.Text
        txtEdit.Visible = False

    End If

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_LostFocus", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If m_Sizable Then
        If PtInRect(m_GripRect, CLng(X), CLng(Y)) Then
            '   Relase any events captured previously
            ReleaseCapture
            '   Send a message that we are resizing the form
            SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
        End If
    End If

    RaiseEvent MouseDown(Button, Shift, X, Y)
    RaiseEvent PanelMouseDown(GetPanelIndex(), Button, Shift, X, Y)
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_MouseDown", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If PtInRect(m_GripRect, CLng(X), CLng(Y)) Then
        UserControl.MousePointer = vbSizeNWSE
    Else
        UserControl.MousePointer = vbDefault

    End If

    RaiseEvent MouseMove(Button, Shift, X, Y)
    RaiseEvent PanelMouseMove(GetPanelIndex(), Button, Shift, X, Y)
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_MouseMove", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent MouseUp(Button, Shift, X, Y)
    RaiseEvent PanelMouseUp(GetPanelIndex(), Button, Shift, X, Y)
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_MouseUp", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_Paint()

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Refresh
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_Paint", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With PropBag
        m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        m_Forecolor = .ReadProperty("ForeColor", Ambient.ForeColor)
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        m_GripShape = .ReadProperty("GripShape", usbSquare)
        m_PanelCount = .ReadProperty("PanelCount", 0)
        m_Sizable = .ReadProperty("Sizable", True)
        Theme = .ReadProperty("Theme", usbAuto)

    End With

    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_Forecolor
    Set UserControl.Font = m_Font
    UserControl.Extender.Align = vbAlignBottom
    m_iTheme = m_Theme

    If Ambient.UserMode Then
        'If we're not in design mode
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "user32.dll")

        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False

            End If

        End If

        If bTrack Then

            'OS supports mouse leave so subclass for it
            With UserControl
                'Start subclassing the UserControl
                Call Subclass_Start(.hWnd)
                Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE, MSG_AFTER)
                Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE, MSG_AFTER)
                Call Subclass_AddMsg(.hWnd, WM_NCPAINT, MSG_AFTER)
                Call Subclass_AddMsg(.hWnd, WM_THEMECHANGED, MSG_AFTER)
                Call Subclass_AddMsg(.hWnd, WM_SIZING, MSG_AFTER)
                Call Subclass_AddMsg(.hWnd, WM_SYSCOLORCHANGE, MSG_AFTER)

            End With

            bSubClass = True

        End If

    End If

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_ReadProperties", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_Resize()

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With UserControl
        '.Height = 360
        .Height = 700

    End With

    With m_GripRect
        .Left = ScaleWidth - 15
        .Top = ScaleHeight - 15
        .Right = .Left + 15
        .Bottom = .Top + 15

    End With

    UserControl.Refresh
    Refresh
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_Resize", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_Show()

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Refresh
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_Show", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub

Private Sub UserControl_Terminate()

Dim i                                   As Long

    'The control is terminating - a good place to stop the subclasser
    'On Error GoTo Catch
    On Error Resume Next

    '   Set the Parents of the Object Back....
    For i = 1 To m_PanelCount

        With m_PanelItems(i)

            If Not .BoundObject Is Nothing Then
                SetParent .BoundObject.hWnd, .BoundParent

            End If

        End With

    Next

    If bSubClass Then
        Call Subclass_StopAll
        bSubClass = False
    End If

Catch:

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With PropBag
        Call .WriteProperty("BackColor", m_BackColor, Ambient.BackColor)
        Call .WriteProperty("ForeColor", m_Forecolor, Ambient.ForeColor)
        Call .WriteProperty("Font", m_Font, Ambient.Font)
        Call .WriteProperty("GripShape", m_GripShape, usbSquare)
        Call .WriteProperty("PanelCount", m_PanelCount, 0)
        Call .WriteProperty("Sizable", m_Sizable, True)
        Call .WriteProperty("Theme", m_Theme, usbAuto)

    End With

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    err.Raise err.Number, "ucStatusBar.UserControl_WriteProperties", err.Description, err.HelpFile, err.HelpContext

    Resume Sub_ErrHandlerExit:

End Sub
