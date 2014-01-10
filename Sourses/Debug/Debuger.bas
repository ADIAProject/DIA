Attribute VB_Name = "Debuger"
Option Explicit

Global giCount_             As Long
Global giAction_            As Integer
Global Proc_Nam_(1 To 2000) As String
Global Proc_Lin_(1 To 2000) As Long
Global Proc_Ptr_            As Integer
Global SysLog_Name_         As String

Public Sub Enter_Proc_(NamP As String)
       Proc_Ptr_ = Proc_Ptr_ + 1
       Proc_Nam_(Proc_Ptr_) = NamP
       Proc_Lin_(Proc_Ptr_) = giCount_
End Sub

Public Sub Exit_Proc_()
       If (Proc_Ptr_ > 0) Then
          Proc_Nam_(Proc_Ptr_) = vbNullString
          Proc_Lin_(Proc_Ptr_) = 0
       End If
       Proc_Ptr_ = Proc_Ptr_ - 1
End Sub

Public Sub Show_Err_(errNum As Long, errDescr As String)
Dim ptrRow As Long
Dim i      As Integer
Dim PP     As Long
Dim ii     As Long
Dim flgLog As Boolean
Dim SysNum As Integer
       flgLog = False
       If (SysLog_Name_ <> vbNullString) Then
          err.Clear
          On Error Resume Next
          SysNum = FreeFile
          Open SysLog_Name_ For Append As #SysNum
          If err.Number = 0 Then flgLog = True
       End If
       If (flgLog) Then
          Print #SysNum, "************* "; Date$; " "; Time$
          Print #SysNum, "      Error: "; CStr(errNum)
          Print #SysNum, " Description: "; errDescr
          Print #SysNum, "   Procedure: "; Proc_Nam_(Proc_Ptr_)
          Print #SysNum, "    Operator: "; CStr(giCount_)
          If (Proc_Ptr_ > 1) Then
             Print #SysNum, "List of procedure:"
             For i = Proc_Ptr_ To 1 Step -1
                 PP = Proc_Lin_(i)
                 If (PP <> 0) Then
                    Print #SysNum, "Operator,Name: "; Format$(PP, "00000"); " "; Proc_Nam_(i)
                 End If
             Next i
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
                    For i = Proc_Ptr_ To 1 Step -1
                        PP = Proc_Lin_(i)
                        If (PP <> 0) Then
                           .TextMatrix(ptrRow, 0) = CStr(PP)
                        End If
                        .TextMatrix(ptrRow, 1) = Proc_Nam_(i)
                        ptrRow = ptrRow + 1
                        .Rows = .Rows + 1
                    Next i
                    ii = .Rows - 1
                    If (.TextMatrix(ii, 0) = vbNullString) And (.TextMatrix(ii, 1) = vbNullString) Then
                       .Rows = .Rows - 1
                    End If
               End With
            End If
            .Show 1
       End With
End Sub

Public Sub App_Terminate()
Dim currFrm As Object
       For Each currFrm In Forms
           Unload currFrm
       Next
       End
End Sub
