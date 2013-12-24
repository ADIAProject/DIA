Attribute VB_Name = "mShellWait"
Option Explicit

'=======================================================================================================================
'···············································    C O N S T A N T S    ···············································
'=======================================================================================================================
Public Const INFINITE As Long = &HFFFFFFFF   'Infinite timeout. Pass INFINITE to ShellW to wait indefinitely until the process terminates
Public Const USER_TIMER_MINIMUM = &HA&       'If uElapse is less than USER_TIMER_MINIMUM (0x0000000A), the timeout is set to USER_TIMER_MINIMUM.
Public Const USER_TIMER_MAXIMUM = &H7FFFFFFF    'If uElapse is greater than USER_TIMER_MAXIMUM (0x7FFFFFFF), the timeout is set to USER_TIMER_MAXIMUM.

'=======================================================================================================================
'············································    E N U M E R A T I O N S    ············································
'=======================================================================================================================
Private Enum SEE_Mask
    SEE_MASK_DEFAULT = &H0                  'Use default values.
    SEE_MASK_CLASSNAME = &H1                'Use the class name given by the lpClass member. If both SEE_MASK_CLASSKEY
    'and SEE_MASK_CLASSNAME are set, the class key is used.
    SEE_MASK_CLASSKEY = &H3                 'Use the class key given by the hkeyClass member. If both SEE_MASK_CLASSKEY
    'and SEE_MASK_CLASSNAME are set, the class key is used.
    SEE_MASK_IDLIST = &H4                   'Use the item identifier list given by the lpIDList member. The lpIDList
    'member must point to an ITEMIDLIST structure.
    SEE_MASK_INVOKEIDLIST = &HC             'Use the IContextMenu interface of the selected item's shortcut menu handler.
    'Use either lpFile to identify the item by its file system path or lpIDList
    'to identify the item by its PIDL. This flag allows applications to use
    'ShellExecuteEx to invoke verbs from shortcut menu extensions instead of the
    'static verbs listed in the registry.
    'Note:  SEE_MASK_INVOKEIDLIST overrides and implies SEE_MASK_IDLIST.
    SEE_MASK_ICON = &H10                    'Use the icon given by the hIcon member. This flag cannot be combined with
    'SEE_MASK_HMONITOR.
    'Note:  This flag is used only in Windows XP and earlier. It is ignored as
    'of Windows Vista.
    SEE_MASK_HOTKEY = &H20                  'Use the keyboard shortcut given by the dwHotKey member.
    SEE_MASK_NOCLOSEPROCESS = &H40          'Use to indicate that the hProcess member receives the process handle. This
    'handle is typically used to allow an application to find out when a process
    'created with ShellExecuteEx terminates. In some cases, such as when
    'execution is satisfied through a DDE conversation, no handle will be
    'returned. The calling application is responsible for closing the handle
    'when it is no longer needed.
    SEE_MASK_CONNECTNETDRV = &H80           'Validate the share and connect to a drive letter. This enables reconnection
    'of disconnected network drives. The lpFile member is a UNC path of a file
    'on a network.
    SEE_MASK_NOASYNC = &H100                'Wait for the execute operation to complete before returning. This flag
    'should be used by callers that are using ShellExecute forms that might
    'result in an async activation, for example DDE, and create a process that
    'might be run on a background thread. (Note: ShellExecuteEx runs on a
    'background thread by default if the caller's threading model is not
    'Apartment.) Calls to ShellExecuteEx from processes already running on
    'background threads should always pass this flag. Also, applications that
    'exit immediately after calling ShellExecuteEx should specify this flag.
    'If the execute operation is performed on a background thread and the caller
    'did not specify the SEE_MASK_ASYNCOK flag, then the calling thread waits
    'until the new process has started before returning. This typically means
    'that either CreateProcess has been called, the DDE communication has
    'completed, or that the custom execution delegate has notified
    'ShellExecuteEx that it is done. If the SEE_MASK_WAITFORINPUTIDLE flag is
    'specified, then ShellExecuteEx calls WaitForInputIdle and waits for the new
    'process to idle before returning, with a maximum timeout of 1 minute.
    'For further discussion on when this flag is necessary, see the Remarks
    'section.
    SEE_MASK_FLAG_DDEWAIT = &H100           'Do not use; use SEE_MASK_NOASYNC instead.
    SEE_MASK_DOENVSUBST = &H200             'Expand any environment variables specified in the string given by the
    'lpDirectory or lpFile member.
    SEE_MASK_FLAG_NO_UI = &H400             'Do not display an error message box if an error occurs.
    SEE_MASK_UNICODE = &H4000               'Use this flag to indicate a Unicode application.
    SEE_MASK_NO_CONSOLE = &H8000&           'Use to inherit the parent's console for the new process instead of having
    'it create a new console. It is the opposite of using a CREATE_NEW_CONSOLE
    'flag with CreateProcess.
    SEE_MASK_ASYNCOK = &H100000             'The execution can be performed on a background thread and the call should
    'return immediately without waiting for the background thread to finish.
    'Note that in certain cases ShellExecuteEx ignores this flag and waits for
    'the process to finish before returning.
    SEE_MASK_HMONITOR = &H200000            'Use this flag when specifying a monitor on multi-monitor systems. The
    'monitor is specified in the hMonitor member. This flag cannot be combined
    'with SEE_MASK_ICON.
    SEE_MASK_NOZONECHECKS = &H800000        'Introduced in Windows XP. Do not perform a zone check. This flag allows
    'ShellExecuteEx to bypass zone checking put into place by IAttachmentExecute.
    SEE_MASK_NOQUERYCLASSSTORE = &H1000000  'Not used.
    SEE_MASK_WAITFORINPUTIDLE = &H2000000   'After the new process is created, wait for the process to become idle
    'before returning, with a one minute timeout. See WaitForInputIdle for more
    'details.
    SEE_MASK_FLAG_LOG_USAGE = &H4000000     'Introduced in Windows XP. Keep track of the number of times this
    'application has been launched. Applications with sufficiently high counts
    'appear in the Start Menu's list of most frequently used programs.
    'SEE_MASK_FLAG_HINST_IS_SITE = &H8000000 'Introduced in Windows 8. The hInstApp member is used to specify the
    'IUnknown of the object that will be used as a site pointer. The site
    'pointer is used to provide services to the ShellExecute function, the
    'handler binding process, and invoked verb handlers.
    'http://msdn.microsoft.com/en-us/library/bb759784(v=vs.85).aspx
End Enum

#If False Then
    Dim SEE_MASK_DEFAULT, SEE_MASK_CLASSNAME, SEE_MASK_CLASSKEY, SEE_MASK_IDLIST, SEE_MASK_INVOKEIDLIST, SEE_MASK_ICON, SEE_MASK_HOTKEY, SEE_MASK_NOCLOSEPROCESS, SEE_MASK_CONNECTNETDRV, SEE_MASK_NOASYNC, SEE_MASK_FLAG_DDEWAIT, SEE_MASK_DOENVSUBST, _
                                SEE_MASK_FLAG_NO_UI, SEE_MASK_UNICODE, SEE_MASK_NO_CONSOLE, SEE_MASK_ASYNCOK, SEE_MASK_HMONITOR, SEE_MASK_NOZONECHECKS, SEE_MASK_NOQUERYCLASSSTORE, SEE_MASK_WAITFORINPUTIDLE, SEE_MASK_FLAG_LOG_USAGE, _
                                SEE_MASK_FLAG_HINST_IS_SITE
#End If

'=======================================================================================================================
'=======================================================================================================================
Private Enum E_ShowCmd
    SW_HIDE = 0            'Hides the window and activates another window.
    SW_SHOWNORMAL = 1      'Activates and displays a window. If the window is minimized or maximized, Windows restores
    'it to its original size and position. An application should specify this flag when
    'displaying the window for the first time.
    SW_SHOWMINIMIZED = 2   'Activates the window and displays it as a minimized window.
    SW_SHOWMAXIMIZED = 3   'Activates the window and displays it as a maximized window.
    SW_MAXIMIZE = 3        'Maximizes the specified window.
    SW_SHOWNOACTIVATE = 4  'Displays a window in its most recent size and position. The active window remains active.
    SW_SHOW = 5            'Activates the window and displays it in its current size and position.
    SW_MINIMIZE = 6        'Minimizes the specified window and activates the next top-level window in the z-order.
    SW_SHOWMINNOACTIVE = 7    'Displays the window as a minimized window. The active window remains active.
    SW_SHOWNA = 8          'Displays the window in its current state. The active window remains active.
    SW_RESTORE = 9         'Activates and displays the window. If the window is minimized or maximized, Windows restores
    'it to its original size and position. An application should specify this flag when restoring
    'a minimized window.
    SW_SHOWDEFAULT = 10    'Sets the show state based on the SW_ flag specified in the STARTUPINFO structure passed to
    'the CreateProcess function by the program that started the application. An application
    'should call ShowWindow with this flag to set the initial show state of its main window.
End Enum

#If False Then             'http://msdn.microsoft.com/en-us/library/bb762153(v=vs.85).aspx
    Dim SW_HIDE, SW_SHOWNORMAL, SW_SHOWMINIMIZED, SW_SHOWMAXIMIZED, SW_MAXIMIZE, SW_SHOWNOACTIVATE, SW_SHOW, SW_MINIMIZE, SW_SHOWMINNOACTIVE, SW_SHOWNA, SW_RESTORE, SW_SHOWDEFAULT
#End If

'=======================================================================================================================
'=======================================================================================================================
Public Enum AppWinStyle                         'WindowStyle constants for ShellW
    vbHide = SW_HIDE
    vbShowNormal = SW_SHOWNORMAL
    vbShowMinimized = SW_SHOWMINIMIZED
    vbShowMaximized = SW_SHOWMAXIMIZED
    vbMaximize = SW_MAXIMIZE
    vbShowNoActivate = SW_SHOWNOACTIVATE
    vbShow = SW_SHOW
    vbMinimize = SW_MINIMIZE
    vbShowMinNoActive = SW_SHOWMINNOACTIVE
    vbShowNA = SW_SHOWNA
    vbRestore = SW_RESTORE
    vbShowDefault = SW_SHOWDEFAULT
End Enum

#If False Then
    Dim vbHide, vbShowNormal, vbShowMinimized, vbShowMaximized, vbMaximize, vbShowNoActivate, vbShow, vbMinimize, vbShowMinNoActive, vbShowNA, vbRestore, vbShowDefault
#End If

'=======================================================================================================================
'=======================================================================================================================
'·······································    T Y P E   D E C L A R A T I O N S    ·······································
'=======================================================================================================================
'=======================================================================================================================
Private Type POINT     'The POINT structure defines the x- and y- coordinates of a point.
    X                                       As Long       'The x-coordinate of the point.
    Y                                       As Long       'The y-coordinate of the point.
End Type            'http://msdn.microsoft.com/en-us/library/dd162805(v=vs.85).aspx

'=======================================================================================================================
'=======================================================================================================================
'Contains message information from a thread's message queue.
'Minimum supported client: Windows 2000 Professional
'http://msdn.microsoft.com/en-us/library/ms644958(v=vs.85).aspx
Private Type Msg
    hWnd                                    As Long   'A handle to the window whose window procedure receives the message. This member is NULL when the message is a thread message.
    Message                                 As Long   'The message identifier. Applications can only use the low word; the high word is reserved by the system.
    wParam                                  As Long   'Additional information about the message. The exact meaning depends on the value of the message member.
    lParam                                  As Long   'Additional information about the message. The exact meaning depends on the value of the message member.
    time                                    As Long   'The time at which the message was posted.
    PT                                      As POINT  'The cursor position, in screen coordinates, when the message was posted.
End Type

'=======================================================================================================================
'=======================================================================================================================
    'Contains information used by ShellExecuteEx.
Private Type SHELLEXECUTEINFO
    cbSize                                  As Long     'Required. The size of this structure, in bytes.
    fMask                                   As SEE_Mask 'Flags that indicate the content and validity of the other structure members; a combination of the following values: (See Enum SEE_Mask above)
    hWnd                                    As Long     'Optional. A handle to the parent window, used to display any message boxes that the system might produce while executing this function. This value can be NULL.
    lpVerb                                  As String   'A string, referred to as a verb, that specifies the action to be performed.
                                                        'The set of available verbs depends on the particular file or folder. Generally, the actions
                                                        'available from an object's shortcut menu are available verbs. This parameter can be NULL,
                                                        'in which case the default verb is used if available. If not, the "open" verb is used. If
                                                        'neither verb is available, the system uses the first verb listed in the registry. The
                                                            'following verbs are commonly used:
                                                            'edit       : Launches an editor and opens the document for editing. If lpFile is not a
                                                            '             document file, the function will fail.
                                                            'explore    : Explores the folder specified by lpFile.
                                                            'find       : Initiates a search starting from the specified directory.
                                                            'open       : Opens the file specified by the lpFile parameter. The file can be an
                                                            '             executable file, a document file, or a folder.
                                                            'openas     : Displays the "Open with" dialog for a file.
                                                            'print      : Prints the document file specified by lpFile. If lpFile is not a document
                                                            '             file, the function will fail.
                                                            'properties : Displays the file or folder's properties.
                                                            'runas      : Grants the user the ability to launch an application with different
                                                            '             credentials.
    lpFile                                  As String   'The address of a null-terminated string that specifies the name of the file or object on
                                                        'which ShellExecuteEx will perform the action specified by the lpVerb parameter. The
                                                        'system registry verbs that are supported by the ShellExecuteEx function include "open"
                                                        'for executable files and document files and "print" for document files for which a print
                                                        'handler has been registered. Other applications might have added Shell verbs through the
                                                        'system registry, such as "play" for .avi and .wav files. To specify a Shell namespace
                                                        'object, pass the fully qualified parse name and set the SEE_MASK_INVOKEIDLIST flag in the
                                                        'fMask parameter.
                                                        'Note:  If the SEE_MASK_INVOKEIDLIST flag is set, you can use either lpFile or lpIDList to
                                                        'identify the item by its file system path or its PIDL respectively. One of the two
                                                        'values — lpFile or lpIDList — must be set.
                                                        'Note:  If the path is not included with the name, the current directory is assumed.
    lpParameters                            As String   'Optional. The address of a null-terminated string that contains the application
                                                        'parameters. The parameters must be separated by spaces. If the lpFile member specifies a
                                                        'document file, lpParameters should be NULL.
    lpDirectory                             As String     'Optional. The address of a null-terminated string that specifies the name of the working directory. If this member is NULL, the current directory is used as the working directory.
    nShow                                   As E_ShowCmd  'Required. Flags that specify how an application is to be shown when it is opened; one of
    'the SW_ values listed for the ShellExecute function. If lpFile specifies a document file,
    'the flag is simply passed to the associated application. It is up to the application to
    'decide how to handle it.
    hInstApp                                As Long     'If SEE_MASK_NOCLOSEPROCESS is set and the ShellExecuteEx call succeeds, it sets this
                                                        'member to a value greater than 32. If the function fails, it is set to an SE_ERR_XXX
                                                        'error value that indicates the cause of the failure. Although hInstApp is declared as an
                                                        'HINSTANCE for compatibility with 16-bit Windows applications, it is not a true HINSTANCE.
                                                        'It can be cast only to an int and compared to either 32 or the following SE_ERR_XXX error codes.
    lpIDList                                As Long      'The address of an absolute ITEMIDLIST structure (PCIDLIST_ABSOLUTE) to contain an item
    'identifier list that uniquely identifies the file to execute. This member is ignored if
    'the fMask member does not include SEE_MASK_IDLIST or SEE_MASK_INVOKEIDLIST.
    lpClass                                 As String    'The address of a null-terminated string that specifies the name of a file type or a GUID.
    'This member is ignored if fMask does not include SEE_MASK_CLASSNAME.
    hkeyClass                               As Long      'A handle to the registry key for the file type. The access rights for this registry key
    'should be set to KEY_READ. This member is ignored if fMask does not include
    'SEE_MASK_CLASSKEY.
    dwHotKey                                As Long      'A keyboard shortcut to associate with the application. The low-order word is the virtual
    'key code, and the high-order word is a modifier flag (HOTKEYF_). For a list of modifier
    'flags, see the description of the WM_SETHOTKEY message. This member is ignored if fMask
    'does not include SEE_MASK_HOTKEY.
    #If True Then
        hIcon                               As Long      'A handle to the icon for the file type. This member is ignored if fMask does not include SEE_MASK_ICON. This value is used only in Windows XP and earlier. It is ignored as of Windows Vista.
    #Else
        hMonitor                            As Long      'A handle to the monitor upon which the document is to be displayed. This member is ignored if fMask does not include SEE_MASK_HMONITOR.
    #End If
    hProcess                            As Long      'A handle to the newly started application. This member is set on return and is always
    'NULL unless fMask is set to SEE_MASK_NOCLOSEPROCESS. Even if fMask is set to
    'SEE_MASK_NOCLOSEPROCESS, hProcess will be NULL if no process was launched. For example,
    'if a document to be launched is a URL and an instance of Internet Explorer is already
    'running, it will display the document. No new process is launched, and hProcess will be
    'NULL.
    'Note:  ShellExecuteEx does not always return an hProcess, even if a process is launched
    'as the result of the call. For example, an hProcess does not return when you use
    'SEE_MASK_INVOKEIDLIST to invoke IContextMenu.
    'Remarks  --------------------------------------------------------------------------------
    'The SEE_MASK_NOASYNC flag must be specified if the thread calling ShellExecuteEx does not
    'have a message loop or if the thread or process will terminate soon after ShellExecuteEx
    'returns. Under such conditions, the calling thread will not be available to complete the
    'DDE conversation, so it is important that ShellExecuteEx complete the conversation before
    'returning control to the calling application. Failure to complete the conversation can
    'result in an unsuccessful launch of the document.
    'If the calling thread has a message loop and will exist for some time after the call to
    'ShellExecuteEx returns, the SEE_MASK_NOASYNC flag is optional. If the flag is omitted,
    'the calling thread's message pump will be used to complete the DDE conversation. The
    'calling application regains control sooner, since the DDE conversation can be completed
    'in the background.
    'When populating the most frequently used program list using the SEE_MASK_FLAG_LOG_USAGE
    'flag in fMask, counts are made differently for the classic and Windows XP-style Start
    'menus. The classic style menu only counts hits to the shortcuts in the Program menu. The
    'Windows XP-style menu counts both hits to the shortcuts in the Program menu and hits to
    'those shortcuts' targets outside of the Program menu. Therefore, setting lpFile to
    'myfile.exe would affect the count for the Windows XP-style menu regardless of whether
    'that file was launched directly or through a shortcut. The classic style — which would
    'require lpFile to contain a .lnk file name — would not be affected.
    'To include double quotation marks in lpParameters, enclose each mark in a pair of
    'quotation marks, as in the following example.
    '    sei.lpParameters = "An example: \"\"\"quoted text\"\"\"";
    'In this case, the application receives three parameters: An, example:, and "quoted text".
    'Minimum supported client: Windows XP
End Type                      'http://msdn.microsoft.com/en-us/library/bb759784(v=vs.85).aspx

'=======================================================================================================================
'=======================================================================================================================
'········································    A P I   D E C L A R A T I O N S    ········································
'=======================================================================================================================
Private Declare Sub PathRemoveArgsW Lib "shlwapi.dll" (ByVal pszPath As Long)
Private Declare Function ExpandEnvironmentStringsW Lib "kernel32.dll" (ByVal lpSrc As Long, Optional ByVal lpDst As Long, Optional ByVal nSize As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long    'BOOL
Private Declare Function WaitMessage Lib "user32.dll" () As Long    'BOOL
Private Declare Function GetProcessId Lib "kernel32.dll" (ByVal hProcess As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Boolean
Private Declare Function PathCanonicalizeW Lib "shlwapi.dll" (ByVal lpszDst As Long, ByVal lpszSrc As Long) As Boolean
Private Declare Function PathGetArgsW Lib "shlwapi.dll" (ByVal pszPath As Long) As Long
Private Declare Function PeekMessageW Lib "user32.dll" (ByRef lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Boolean
Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, Optional ByVal lpTimerFunc As Long) As Long
Private Declare Function ShellExecuteExW Lib "shell32.dll" (ByVal pExecInfo As Long) As Boolean
Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'=======================================================================================================================
'··········································    P U B L I C   M E T H O D S    ··········································
'=======================================================================================================================
'!--------------------------------------------------------------------------------
'! Procedure   (Ôóíêöèÿ)   :   Function ShellW
'! Description (Îïèñàíèå)  :   [Runs an executable program or document; optionally waits for a specified amount of time before resuming execution]
'! Parameters  (Ïåðåìåííûå):   PathName (String)
'                              WindowStyle (AppWinStyle = vbShowNormal)
'                              Wait (Long)
'!--------------------------------------------------------------------------------
Public Function ShellW(ByRef PathName As String, Optional ByVal WindowStyle As AppWinStyle = vbShowNormal, Optional ByVal Wait As Long) As Long

    Const o = 0&, l = 1&, MAX_PATH = 260&
    'Messages are not removed from the queue after processing by PeekMessage.
    Const PM_NOREMOVE = &H0&
    'Process all posted messages, including timers and hotkeys.
    Const PM_QS_POSTMESSAGE = &H980000
    'The time-out interval elapsed, and the object's state is nonsignaled.
    Const WAIT_TIMEOUT = &H102&
    'Posted to the installing thread's message queue when a timer expires.
    Const WM_TIMER = &H113&
    Dim TimedOut As Boolean, Tmr1 As Long, Tmr2 As Long, m As Msg, SEI As SHELLEXECUTEINFO
    Static Busy  As Boolean
    
    'Reset Err object everytime this function is called
    Err.Clear

    'This function shouldn't be called
    If Not Busy Then
        'more than once at any given time
        Busy = True

        'See if there's anything to do
        If LenB(PathName) Then
            
            With SEI
                .cbSize = LenB(SEI)
                .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_DOENVSUBST Or SEE_MASK_FLAG_NO_UI
                'Suppress error message
                .nShow = WindowStyle

                'Expand environment variables
                If InStr(PathName, Percentage) Then
                    SysReAllocStringLen VarPtr(.lpFile), , ExpandEnvironmentStringsW(StrPtr(PathName)) - l
                    ExpandEnvironmentStringsW StrPtr(PathName), StrPtr(.lpFile), Len(.lpFile) + l
                Else
                    .lpFile = PathName
                    'ShellExecuteEx doesn't expand environment vars in lpParameters
                End If

                Select Case True
                    
                    'Look for "\.", "\..", ".\" or "..\"
                    Case InStr(.lpFile, "\.") <> o, InStr(.lpFile, ".\") <> o

                        
                        If Len(.lpFile) < MAX_PATH Then
                            SysReAllocStringLen VarPtr(.lpVerb), , MAX_PATH - l

                            'Simplify the given path
                            If PathCanonicalizeW(StrPtr(.lpVerb), StrPtr(.lpFile)) Then
                                'by removing "." & ".."
                                SysReAllocString VarPtr(.lpFile), StrPtr(.lpVerb)
                            End If

                            .lpVerb = vbNullString
                        End If

                End Select

                'Extract arguments, if any
                SysReAllocString VarPtr(.lpParameters), PathGetArgsW(StrPtr(.lpFile))

                'If there are, then trim the
                If LenB(.lpParameters) Then
                    PathRemoveArgsW StrPtr(.lpFile)
                    'original args from lpFile PathRemoveArgsW StrPtr(.lpFile)
                    'original args from lpFile If InStr(.lpParameters, """") Then .lpParameters = Replace(.lpParameters, """", """""""")
                    'MSDN's instructions don't seem to work in XP
                End If
                
                'Run the specified executable or document
                If ShellExecuteExW(VarPtr(SEI)) Then
                    'Return the Task ID, a.k.a. Process ID
                    ShellW = GetProcessId(.hProcess)

                    'If specified, wait Wait millisecs before returning
                    If Wait Then
                    
                        'If specified waiting time isn't INFINITE or negative then set a timer with the given duration
                        If Wait > INFINITE Then Tmr1 = SetTimer(o, .hProcess, Wait)
                        'and another one with a very short interval used to ensure a constant flow of messages
                        Tmr2 = SetTimer(o, App.ThreadID, 250&)

                        Do
                            'This API suspends the thread if no new messages have arrived
                            WaitMessage

                            If Tmr1 Then

                                'Check the message queue for WM_TIMER messages only
                                If PeekMessageW(m, -l, WM_TIMER, WM_TIMER, PM_QS_POSTMESSAGE) Then
                                    If m.wParam = Tmr1 Then Err.Clear

                                    Exit Do
                                
                                'Reset Err (in case it was raised elsewhere) if the Timer ID
                                End If
                            
                            'is the one for the wait interval and break out of the loop
                            End If
                            
                            'Let the system perform other tasks
                            DoEvents
                            'The process becomes signaled when it ends, WaitForSingleObject thus returns WAIT_OBJECT_0
                            TimedOut = (WaitForSingleObject(.hProcess, o) = WAIT_TIMEOUT)
                        'http://msdn.microsoft.com/en-us/library/ms683189(v=vs.85).aspx#1
                        Loop While TimedOut
                        
                        Tmr2 = KillTimer(o, Tmr2)

                        If Tmr1 Then Tmr1 = KillTimer(o, Tmr1)

                        'WaitForSingleObject didn't timeout, therefore the process has terminated
                        If Not TimedOut Then
                            Tmr1 = GetExitCodeProcess(.hProcess, ShellW)
                            Debug.Assert Tmr1
                            'Return the terminated
                            Err = vbObjectError
                            'Set the Err object's properties instead of raising an error;
                            'this is similar to the API's use of Get/SetLastError
                            Err.Description = "Exit Code"
                        End If
                        
                    End If
                    
                    Tmr2 = CloseHandle(.hProcess)
                    Debug.Assert Tmr2
                End If

            End With

        End If
        
        'Reset flag
        Busy = False
    End If

End Function

'ShellW returns either the Process ID, the Exit Code or zero (check Err.Number to make sure)
'=======================================================================================================================
'··············································    R E F E R E N C E S    ··············································
'=======================================================================================================================
'How To Replace IsTask() with GetExitCodeProcess()
'http://support.microsoft.com/kb/108228
'Launching Applications (ShellExecute, ShellExecuteEx, SHELLEXECUTEINFO)
'http://msdn.microsoft.com/en-us/library/bb776886(v=vs.85).aspx
'Run Method (Windows Script Host)
'http://msdn.microsoft.com/en-us/library/d5fk67ky(v=vs.84).aspx
