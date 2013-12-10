Attribute VB_Name = "mBrowseForFolder"
Option Explicit

' ***** Глобальные переменные для связи между fBrowseForFolder() и callback-функцией: *******************************
' пользовательский заголовок диалога
Public g_DialogTitle                    As String

' центрировать ли диалог на экране
Public g_CenterOnScreen                 As Boolean

' если = True, то диалог будет поверх всех открытых окон
Private g_TopMost                       As Boolean
Private g_newLeft                       As Long
Private g_newTop                        As Long

Public g_CurrentDirectory               As String

' *******************************************************************************************************************
Public Const BFFM_INITIALIZED           As Long = 1

Private Const BFFM_SELCHANGED           As Long = 2
Private Const BFFM_SETSTATUSTEXT        As Long = (WM_USER + 100)
Private Const BFFM_ENABLEOK             As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTION         As Long = (WM_USER + 102)
Private Const BFFM_SETEXPANDED = (WM_USER + 106)

Public Type BROWSEINFO
    hWndOwner                               As Long
    pIDLRoot                            As Long
    pszDisplayName                      As Long
    lpszTitle                           As Long
    ulFlags                             As Long
    lpfnCallback                        As Long
    lParam                              As Long
    iImage                              As Long

End Type

'
' NOTE: Many of these flags only work with certain versions of Shell32.dll:
'
Public Enum WhatBrowse
    'Only return file system directories. If the user selects
    'folders that are not part of the file system, the OK
    'button is grayed:
    BIF_RETURNONLYFSDIRS = &H1
    'The browse dialog will display files as well as folders:
    BIF_BROWSEINCLUDEFILES = &H1 Or &H4000
    'Only return computers. If the user selects anything
    'other than a computer, the OK button is grayed:
    BIF_BROWSEFORCOMPUTER = &H1000
    'Only return printers. If the user selects anything
    'other than a printer, the OK button is grayed:
    BIF_BROWSEFORPRINTER = &H2000
    'Do not include network folders below the domain
    'level in the tree view control:
    BIF_DONTGOBELOWDOMAIN = &H2
    'Include a status area in the dialog box. The callback
    'function can set the status text by sending messages
    'to the dialog box:
    BIF_STATUSTEXT = &H4
    'Use the new user-interface providing the user with a larger
    'resizable dialog box which includes drag and drop, reordering,
    'context menus, new folders, delete, and other context menu
    'commands:
    BIF_NEWDIALOGSTYLE = &H40
    'Include an edit control in the dialog box:
    BIF_EDITBOX = &H10
    'Equivalent to BIF_EDITBOX | BIF_NEWDIALOGSTYLE:
    BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
    'Only return file system ancestors. If the user
    'selects anything other than a file system ancestor,
    'the OK button is grayed:
    BIF_SHAREABLE = &H8000
    BIF_BROWSEFILEJUNCTIONS = &H10000
    BIF_UAHINT = &H100
    BIF_NONEWFOLDERBUTTON = &H200
    BIF_NOTRANSLATETARGETS = &H400
    BIF_BROWSEINCLUDEURLS = &H80
    BIF_RETURNFSANCESTORS = &H8
    BIF_VALIDATE = &H20
    BIF_DEFAULT = BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_STATUSTEXT

End Enum

#If False Then

    Private BIF_RETURNONLYFSDIRS, BIF_BROWSEINCLUDEFILES, BIF_BROWSEFORCOMPUTER, BIF_BROWSEFORPRINTER, BIF_DONTGOBELOWDOMAIN
    Private BIF_STATUSTEXT, BIF_NEWDIALOGSTYLE, BIF_EDITBOX, BIF_USENEWUI, BIF_RETURNFSANCESTORS
#End If

Private Declare Function SHBrowseForFolder _
                          Lib "shell32" _
                              Alias "SHBrowseForFolderW" (lpBrowseInfo As BROWSEINFO) As Long

'Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" (ByVal pidList As Long, ByRef lpBuffer As Byte) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" (ByVal pIDList As Long, ByVal lpBuffer As Long) As Long

Public Function BrowseCallbackProc(ByVal hWnd As Long, _
                                   ByVal uMsg As Long, _
                                   ByVal lParam As Long, _
                                   ByVal lpData As Long) As Long

Dim lRet                                As Long
Dim sBuffer                             As String
Dim Fhwnd                               As Long
Dim szPath()                            As Byte

    On Error GoTo errhandler

    Select Case uMsg

        Case BFFM_INITIALIZED
            'DebugMode "Show Open Dialog: Func - 'BrowseCallbackProc:BFFM_INITIALIZED'", 2
            BFFSetPath hWnd, lpData, True
            Fhwnd = FindWindowEx(hWnd, ByVal 0&, "Edit", ByVal lpData)
            SendMessage Fhwnd, EM_NOSETFOCUS, 0&, 0&

            ' << надо только центрировать диалог на экране
            If g_CenterOnScreen Then
                'DebugMode "Show Open Dialog: Func - 'BrowseCallbackProc:CenterDialog'", 2
                CenterDialog hWnd

            End If

        Case BFFM_SELCHANGED
            'DebugMode "Show Open Dialog: Func - 'BrowseCallbackProc:BFFM_SELCHANGED'", 2
            sBuffer = String$(MAX_PATH_UNICODE, vbNullChar)
            'DebugMode "Show Open Dialog: Func - 'BrowseCallbackProc:SHGetPathFromIDList'", 2
            lRet = SHGetPathFromIDList(lParam, StrPtr(sBuffer))

            If lRet = 1 Then
                'DebugMode "Show Open Dialog: Func - 'BrowseCallbackProc:FindWindowEx'", 2
                Fhwnd = FindWindowEx(hWnd, ByVal 0&, "Edit", vbNullString)
                szPath = BackslashAdd2Path(TrimNull(sBuffer)) & vbNullChar
                'DebugMode "Show Open Dialog: Func - 'BrowseCallbackProc:WM_SETTEXT'", 2
                SendMessageLong Fhwnd, WM_SETTEXT, 0, VarPtr(szPath(0))
                'DebugMode "Show Open Dialog: Func - 'BrowseCallbackProc:EM_SETREADONLY'", 2
                SendMessage Fhwnd, EM_SETREADONLY, True, 0&
                'DebugMode "Show Open Dialog: Func - 'BrowseCallbackProc:WM_KILLFOCUS'", 2
                SendMessage Fhwnd, WM_KILLFOCUS, 0&, ByVal 0&
                'DebugMode "Show Open Dialog: Func - 'BrowseCallbackProc:BFFEnableOKButton'", 2
                BFFEnableOKButton hWnd, CBool(Len(TrimNull(sBuffer)))
            Else
                BFFEnableOKButton hWnd, False

            End If

    End Select

errhandler:
    BrowseCallbackProc = 0

End Function

Private Sub CenterDialog(ByVal lngHWnd As Long)

Dim screenWidth                         As Long
Dim screenHeight                        As Long
Dim winWidth                            As Long
Dim winHeight                           As Long
Dim R                                   As RECT

    ' Определяем размеры экрана:
    screenWidth = Screen.Width / Screen.TwipsPerPixelX
    screenHeight = Screen.Height / Screen.TwipsPerPixelY
    ' Определяем размеры окна диалога:
    GetWindowRect lngHWnd, R
    ' Рассчитываем текущие размеры диалога:
    winWidth = (R.Right - R.Left)
    winHeight = (R.Bottom - R.Top)
    g_newLeft = (screenWidth - winWidth) / 2
    g_newTop = (screenHeight - winHeight) / 2
    ' Центрируем на экране:
    SetWindowPos lngHWnd, 0, g_newLeft, g_newTop, winWidth, winHeight, SWP_SHOWWINDOW

End Sub

Public Function fBrowseForFolder(ByVal hWnd_Owner As Long, _
                                 ByVal WhatBr As WhatBrowse, _
                                 Optional ByVal sPrompt As String = "Please Select a Folder:", _
                                 Optional ByVal InitDir As String = vbNullString, _
                                 Optional ByVal CenterOnScreen As Boolean = False, _
                                 Optional ByVal TopMost As Boolean) As String

' *** Модифицированная функция fBrowseForFolder. Получает на вход следующие аргументы:
' ***   hWnd_Owner      - hWnd вызывающего объекта,
' ***   sPrompt         - текст подсказки,
' ***   WhatBr          - комбинация (через OR) элементов перечислителя WhatBrowse,
' ***   DialogTitle     - пользовательский заголовок диалога
' ***   initDir         - начальная папка обзора (если не задана, то "Мой компьютер"),
' ***   CenterOnScreen  - центрировать ли диалог на экране
' ***   TopMost         - если = True, то диалог будет поверх всех открытых окон
Dim lpIDList                            As Long
Dim sBuffer                             As String
Dim szTitleInfo()                       As Byte
Dim udtBI                               As BROWSEINFO

    ' Устанавливаем глобальные переменные, чтобы аргументы fBrowseForFolder() "дошли" до callback-функции:
    'DebugMode "Show Open Dialog: fBrowseForFolder - Set Global variation", 2
    szTitleInfo = sPrompt & vbNullChar
    g_CenterOnScreen = CenterOnScreen
    g_TopMost = TopMost

    With udtBI
        .hWndOwner = hWnd_Owner
        .lpszTitle = VarPtr(szTitleInfo(0))
        .ulFlags = WhatBr
        .lParam = StrPtr(InitDir & vbNullChar)
        'DebugMode "Show Open Dialog: Func - 'GetAddressofFunction(AddressOf BrowseCallbackProc)'", 2
        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)

    End With

    'DebugMode "Show Open Dialog: Func - 'SHBrowseForFolder(udtBI)'", 2
    lpIDList = SHBrowseForFolder(udtBI)

    If lpIDList Then
        sBuffer = String$(MAX_PATH_UNICODE, vbNullChar)
        'DebugMode "Show Open Dialog: Func - 'SHGetPathFromIDList'", 2

        If SHGetPathFromIDList(lpIDList, StrPtr(sBuffer)) Then
            'sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            'DebugMode "Show Open Dialog: Func - 'BackslashAdd2Path'", 2
            fBrowseForFolder = BackslashAdd2Path(TrimNull(sBuffer))
        Else
            fBrowseForFolder = vbNullString

        End If

        'DebugMode "Show Open Dialog: Func - 'CoTaskMemFree lpIDList'", 2
        CoTaskMemFree lpIDList
    Else
        fBrowseForFolder = vbNullString
    End If

    'DebugMode "Show Open Dialog: Func - 'CoTaskMemFree(udtBI.lParam)'", 2
    Call CoTaskMemFree(udtBI.lParam)
    ' Result in Debuglog
    'DebugMode "Show Open Dialog: fBrowseForFolder =" & fBrowseForFolder, 2

End Function

Private Function GetAddressofFunction(Add As Long) As Long
    ' This function allows you to assign a function pointer to a vaiable.
    GetAddressofFunction = Add

End Function

'   Used to enable or disable the dialog's OK button from the BFF callback
Private Sub BFFEnableOKButton(ByVal hWndDialog As Long, ByVal Enable As Boolean)
    SendMessageLong hWndDialog, BFFM_ENABLEOK, 0, ByVal Abs(Enable)

End Sub

Private Sub BFFSetPath(ByVal hWndDialog As Long, _
                       lpData As Long, _
                       ByVal UseStrPath As Boolean)
    SendMessageLong hWndDialog, BFFM_SETSELECTION, Abs(UseStrPath), ByVal lpData
    'If IsWin7 Then 'если этого не делать, то скроллинг на Win7 не гарантирован
    'http://connect.microsoft.com/VisualStudio/feedback/details/518103/bffm-setselection-does-not-work-with-shbrowseforfolder-on-windows-7
    Sleep 200
    PostMessage hWndDialog, BFFM_SETEXPANDED, Abs(UseStrPath), ByVal lpData

    'End If
End Sub
