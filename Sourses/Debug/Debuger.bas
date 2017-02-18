Attribute VB_Name = "Debuger"
Option Explicit

Global giCount_             As Long
Global giAction_            As Long
Global Proc_Nam_(1 To 2000) As String
Global Proc_Lin_(1 To 2000) As Long
Global Proc_Ptr_            As Long
Global SysLog_Name_         As String

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Enter_Proc_
'! Description (Описание)  :   [Вход в процедуру - запоминаем строку]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Enter_Proc_(NamP As String)
       Proc_Ptr_ = Proc_Ptr_ + 1
       Proc_Nam_(Proc_Ptr_) = NamP
       Proc_Lin_(Proc_Ptr_) = giCount_
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Exit_Proc_
'! Description (Описание)  :   [Выход из процедуры]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Exit_Proc_()
       If (Proc_Ptr_ > 0) Then
          Proc_Nam_(Proc_Ptr_) = vbNullString
          Proc_Lin_(Proc_Ptr_) = 0
       End If
       Proc_Ptr_ = Proc_Ptr_ - 1
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Show_Err_
'! Description (Описание)  :   [Показать ошибку - вывести форму]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub Show_Err_(errNum As Long, errDescr As String)
Dim ptrRow As Long
Dim ii     As Long
Dim PP     As Long
Dim iii    As Long
Dim flgLog As Boolean
Dim SysNum As Integer

       flgLog = False
       If (SysLog_Name_ <> vbNullString) Then
          Err.Clear
          On Error Resume Next
          SysNum = FreeFile
          Open SysLog_Name_ For Append As #SysNum
          If Err.Number = 0 Then flgLog = True
       End If
       
       If (flgLog) Then
          Print #SysNum, "************* "; Date$; " "; Time$
          Print #SysNum, "      Error: "; CStr(errNum)
          Print #SysNum, " Description: "; errDescr
          Print #SysNum, "   Procedure: "; Proc_Nam_(Proc_Ptr_)
          Print #SysNum, "    Operator: "; CStr(giCount_)
          If (Proc_Ptr_ > 1) Then
             Print #SysNum, "List of procedure:"
             For ii = Proc_Ptr_ To 1 Step -1
                 PP = Proc_Lin_(ii)
                 If (PP <> 0) Then
                    Print #SysNum, "Operator,Name: "; Format$(PP, "00000"); " "; Proc_Nam_(ii)
                 End If
             Next ii
          End If
          Close #SysNum
       End If
       
       With frmError
            .lblErrCode.Caption = CStr(errNum)
            .lblErrDescr.Text = errDescr
            .lblProc.Caption = Proc_Nam_(Proc_Ptr_)
            .lblStmt.Caption = CStr(giCount_)
            
            If (Proc_Ptr_ = 1) Then
               .MSFlexGrid1.Visible = False
               .Label2.Visible = False
            Else
               .MSFlexGrid1.Visible = True
               .Label2.Visible = True
               
               With .MSFlexGrid1
                    ptrRow = 1
                    For ii = Proc_Ptr_ To 1 Step -1
                        PP = Proc_Lin_(ii)
                        If (PP <> 0) Then
                           .TextMatrix(ptrRow, 0) = CStr(PP)
                        End If
                        .TextMatrix(ptrRow, 1) = Proc_Nam_(ii)
                        ptrRow = ptrRow + 1
                        .Rows = .Rows + 1
                    Next ii
                    iii = .Rows - 1
                    If (.TextMatrix(iii, 0) = vbNullString) And (.TextMatrix(iii, 1) = vbNullString) Then
                       .Rows = .Rows - 1
                    End If
               End With
               
            End If
            
            .Show 1
       End With
       
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub App_Terminate
'! Description (Описание)  :   [Закрытие приложения]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub App_Terminate()
Dim objCurrFrm As Object

       For Each objCurrFrm In Forms
           Unload objCurrFrm
       Next
       
       End
End Sub
