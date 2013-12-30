Attribute VB_Name = "mTimedMsgBox"
Option Explicit

' IMPORTANT NOTE:
' Demo project showing how to use the Timed MessageBox
' by Anirudha Vengurlekar anirudhav@yahoo.com(http://domaindlx.com/anirudha)
' this demo is released into the public domain "as is" without
' warranty or guaranty of any kind.  In other words, use at your own risk.
' Please send me you comments or suggestions at anirudhav@yahoo.com
' Thanks in advance.
Private Const WH_CBT        As Integer = 5
Private Const HCBT_ACTIVATE As Integer = 5
Private Const BN_CLICKED    As Integer = 0

' Used for storing information
Private m_lMsgHandle        As Long
Private m_TimeMsgBox        As Long
Private m_lNoHandle         As Long
Private m_lhHook            As Long
Private bTimedOut           As Boolean
Private sMsgText            As String

Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function GetDlgCtrlID Lib "user32.dll" (ByVal hWnd As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function EnumChildWindowsProc
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lngHWnd (Long)
'                              lParam (Long)
'!--------------------------------------------------------------------------------
Private Function EnumChildWindowsProc(ByVal lngHWnd As Long, ByVal lParam As Long) As Long

    Dim lRet       As Long
    Dim sClassName As String

    sClassName = String$(100, vbNullChar)
    lRet = GetClassName(lngHWnd, sClassName, Len(sClassName))
    sClassName = Left$(sClassName, lRet)

    If UCase$(sClassName) = UCase$("Button") Then
        m_lNoHandle = lngHWnd
        EnumChildWindowsProc = 0
    Else
        EnumChildWindowsProc = 1
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetMessageBoxHandle
'! Description (Описание)  :   [THIS IS CALLBACK procedure. Will called by Hook procedure]
'! Parameters  (Переменные):   lMsg (Long)
'                              wParam (Long)
'                              lParam (Long)
'!--------------------------------------------------------------------------------
Private Function GetMessageBoxHandle(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If lMsg = HCBT_ACTIVATE Then
        'Release the CBT hook
        m_lMsgHandle = wParam
        ' Msg Box Window Handle
        UnhookWindowsHookEx m_lhHook
        m_lhHook = 0
        ' enumerate all the children so we can send a number
        ' button message to the No button if our box has one
        ' this avoids the Microsoft error in the message box
        ' Added by Daniels, Michael A (KPMG Group)
        EnumChildWindows m_lMsgHandle, AddressOf EnumChildWindowsProc, 0
    End If

    GetMessageBoxHandle = False
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub MessageBoxTimerEvent
'! Description (Описание)  :   [THIS IS CALLBACK procedure. Will called by timer procedure. This function is called when time out occurs by the timer]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub MessageBoxTimerEvent()

    Dim lButtonCommand As Integer

    If m_lNoHandle = 0 Then
        SendMessage m_lMsgHandle, WM_CLOSE, 0, 0
    Else
        lButtonCommand = (BN_CLICKED * (2 ^ 16)) And &HFFFF
        lButtonCommand = lButtonCommand Or GetDlgCtrlID(m_lNoHandle)
        SendMessage m_lMsgHandle, WM_COMMAND, lButtonCommand, m_lNoHandle
    End If

    m_lMsgHandle = 0
    m_lNoHandle = 0
    bTimedOut = True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub MessageBoxTimerUpdateEvent
'! Description (Описание)  :   [THIS IS CALLBACK procedure. Will called by timer procedure. This function is called when time out occurs by the timer]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub MessageBoxTimerUpdateEvent()

    Dim lRet As Long
    Dim sStr As String

    If Not (m_lMsgHandle = 0) Then
        m_TimeMsgBox = m_TimeMsgBox - 1

        If LenB(sMsgText) = 0 Then
            sStr = Space$(255)
            lRet = GetWindowText(m_lMsgHandle, sStr, 255)
            sStr = Left$(sStr, lRet)
            sMsgText = sStr
        End If

        sStr = sMsgText & " (Time left: " & m_TimeMsgBox & " seconds)"
        SetWindowText m_lMsgHandle, sStr
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function MsgBoxEx
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sMsgText (String)
'                              dwWait (Long)
'                              Buttons (VbMsgBoxStyle = vbOKOnly)
'                              sTitle (String = "Timed MessageBox Demo")
'!--------------------------------------------------------------------------------
Public Function MsgBoxEx(sMsgText As String, dwWait As Long, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional sTitle As String = "Timed MessageBox") As VbMsgBoxResult

    Dim lTimer       As Long
    Dim lTimerUpdate As Long

    m_TimeMsgBox = dwWait
    ' SET CBT hook
    m_lhHook = SetWindowsHookEx(WH_CBT, AddressOf GetMessageBoxHandle, App.hInstance, GetCurrentThreadId())
    ' set the timer
    lTimer = SetTimer(0, 0, dwWait * 1000, AddressOf MessageBoxTimerEvent)
    ' Set timer
    lTimerUpdate = SetTimer(0, 0, 1 * 1000, AddressOf MessageBoxTimerUpdateEvent)
    ' Set the flag to false
    bTimedOut = False
    ' Display the message Box
    MsgBoxEx = MsgBox(sMsgText, Buttons, sTitle)
    ' Kill the timer
    KillTimer 0, lTimer
    KillTimer 0, lTimerUpdate
    ' Return ZERO so that caller routine will decide what to do
    sMsgText = vbNullString
    m_TimeMsgBox = 0

    If bTimedOut Then
        MsgBoxEx = 0
    End If

End Function
