Attribute VB_Name = "mCommonHash"
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
Private Const MODULE_NAME    As String = "mCommonHash"
Private Const KB_32          As Long = 32768

' ***************************************************************************
' Global constants
' ***************************************************************************
Private Const IDOK           As Long = 1       ' one button return value

' ***************************************************************************
' Module Constants
' ***************************************************************************
Private Const MB_OK          As Long = &H0&    ' one button
Private Const MB_YESNO       As Long = &H4&    ' two buttons
Private Const MB_YESNOCANCEL As Long = &H3&    ' three buttons

' ***************************************************************************
' Type structures
' ***************************************************************************
' UDT for passing data through the hook
Private Type MSGBOX_HOOK_PARAMS
    hWndOwner                               As Long
    hHook                               As Long
End Type

' ***************************************************************************
' Global Variables
'
' Variable name:     gblnStopProcessing
' Naming standard:   g bln StopProcessing
'                    - --- -------------
'                    |  |    |______ Variable subname
'                    |  |___________ Data type (Boolean)
'                    |______________ Global level designator
'
' ***************************************************************************
Public gblnStopProcessing As Boolean

' **************************************************************************
' Routine:       CalcProgress
'
' Description:   Calculates current amount of completion
'
' Parameters:    curCurrAmt - current value
'                curMaxAmount - maximum value to be MAX_PERCENT%
'
' Returns:       percentage of progression
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CalcProgress
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   curCurrAmt (Currency)
'                              curMaxAmount (Currency)
'!--------------------------------------------------------------------------------
Public Function CalcProgress(ByVal curCurrAmt As Currency, ByVal curMaxAmount As Currency) As Long

    Dim lngPercent    As Long
    Dim curAmtLeft    As Currency

    Const MAX_PERCENT As Long = 100

    ' percentage to be calcuated
    ' difference between current and max amount
    ' Reset progress bar
    DoEvents

    If (curCurrAmt <= 0@) Or (curMaxAmount <= 0@) Then
        CalcProgress = 0

        Exit Function

    End If

    ' Make sure current value does
    ' not exceed maximum value
    If curCurrAmt >= curMaxAmount Then
        curCurrAmt = curMaxAmount
    End If

    ' Calculate percentage based on current
    ' value and maximum allowable value
    On Error Resume Next

    curAmtLeft = curMaxAmount - curCurrAmt
    lngPercent = MAX_PERCENT - CLng((curAmtLeft * MAX_PERCENT) / curMaxAmount)

    On Error GoTo 0

    ' Nullify this error trap
    ' Validate percentage so we
    ' do not exceed our bounds
    Select Case lngPercent

        Case Is < 0
            lngPercent = 0

        Case Is > MAX_PERCENT
            lngPercent = MAX_PERCENT
    End Select

    CalcProgress = lngPercent
    ' Return calculated percent
End Function

' ***************************************************************************
'  Routine:     ErrorMsg
'
'  Description: Displays a standard VB MsgBox formatted to display severe
'               (Usually application-type) error messages.
'
'  Parameters:  strModule - The module where the error occurred
'               strRoutine - The routine where the error occurred
'               strMsg - The error message
'               strCaption - The MsgBox caption  (optional)
'
'  Returns:     None
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ErrorMsg
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strModule (String)
'                              strRoutine (String)
'                              strMsg (String)
'                              strCaption (String = vbNullString)
'!--------------------------------------------------------------------------------
Public Sub ErrorMsg(ByVal strModule As String, ByVal strRoutine As String, ByVal strMsg As String, Optional ByVal strCaption As String = vbNullString)

    Dim strNewCaption As String
    Dim strFullMsg    As String

    ' Formatted MsgBox caption
    ' Formatted message
    ' Make sure strModule is populated
    If LenB(Trim$(strModule)) = 0 Then
        strModule = "Unknown"
    End If

    ' Make sure strRoutine is populated
    If LenB(Trim$(strRoutine)) = 0 Then
        strRoutine = "Unknown"
    End If

    ' Make sure strMsg is populated
    If LenB(Trim$(strMsg)) = 0 Then
        strMsg = "Unknown"
    End If

    ' Format the MsgBox caption
    strNewCaption = strFormatCaption(strCaption, True)
    ' Format the message
    strFullMsg = "Module: " & vbTab & strModule & vbCr & "Routine:" & vbTab & strRoutine & vbCr & "Error:  " & vbTab & strMsg
    ' the MsgBox routine
    MsgBox strFullMsg, vbCritical Or vbOKOnly, strNewCaption
End Sub

' **************************************************************************
' Routine:       GetBlockSize
'
' Description:   Determines the size of the data to be processed. This
'                process has been speeded up by 50% or more by adjusting
'                the record length based on amount of data left to process.
'
' Parameters:    curAmtLeft - Amount of data left
'
' Returns:       New record size as a long integer
'
' ***************************************************************************
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetBlockSize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   curAmtLeft (Currency)
'!--------------------------------------------------------------------------------
Public Function GetBlockSize(ByVal curAmtLeft As Currency) As Long

    ' Determine the amount of data to process
    Select Case curAmtLeft

        Case Is >= KB_32
            GetBlockSize = KB_32

        Case Else
            GetBlockSize = CLng(curAmtLeft)
    End Select

End Function

' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************
' ***************************************************************************
'  Routine:     InfoMsg
'
'  Description: Displays a VB MsgBox with no return values.  It is designed to
'               be used where no response from the user is expected other than
'               "OK".
'
'  Parameters:  strMsg - The message text
'               strCaption - The MsgBox caption (optional)
'
'  Returns:     None
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub InfoMsg
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strMsg (String)
'                              strCaption (String = vbNullString)
'!--------------------------------------------------------------------------------
Public Sub InfoMsg(ByVal strMsg As String, Optional ByVal strCaption As String = vbNullString)

    Dim strNewCaption As String

    ' Formatted MsgBox caption
    ' Format the MsgBox caption
    strNewCaption = strFormatCaption(strCaption)
    ' the MsgBox routine
    MsgBox strMsg, vbInformation Or vbOKOnly, strNewCaption
End Sub

' ***************************************************************************
'  Routine:     FormatCaption
'
'  Description: Formats the caption text to use the application title as
'               default
'
'  Parameters:  strCaption - The input caption which may be appended to the
'                            application title.
'               bError - Add "Error" to the caption
'
'  Returns:     Formatted string to be used as a msgbox caption
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function strFormatCaption
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strCaption (String)
'                              bError (Boolean = False)
'!--------------------------------------------------------------------------------
Private Function strFormatCaption(ByVal strCaption As String, Optional ByVal bError As Boolean = False) As String

    Dim strNewCaption As String

    ' The formatted caption
    ' Set the caption to either input parm or the application name
    If LenB(Trim$(strCaption)) > 0 Then
        strNewCaption = Trim$(strCaption)
    Else
        ' Set the caption default
        strNewCaption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    End If

    ' Optionally, add error text
    If bError Then
        strNewCaption = strNewCaption & " Error"
    End If

    ' Return the new caption
    strFormatCaption = strNewCaption
End Function
