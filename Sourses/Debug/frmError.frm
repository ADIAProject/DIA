VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warning, Error in Application!"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmError.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue running program (Not recommended)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   5400
      Width           =   2655
   End
   Begin VB.CommandButton cmdCreateFile 
      Caption         =   "Create Error.log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   75
      TabIndex        =   6
      ToolTipText     =   "Создать файл с описанием ошибки"
      Top             =   5400
      Width           =   1380
   End
   Begin VB.CommandButton cmdEmail 
      Caption         =   "E-mail to author (do not forget to attach Error.log)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   "Не забудьте прикрепить к письму созданный файл error.log"
      Top             =   5400
      Width           =   2580
   End
   Begin VB.Frame Frame1 
      Caption         =   "Description of error"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   75
      TabIndex        =   8
      Top             =   840
      Width           =   8500
      Begin VB.TextBox lblErrDescr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   600
         Width           =   8325
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number of application error:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   11
         Top             =   240
         Width           =   3705
      End
      Begin VB.Label lblErrCode 
         AutoSize        =   -1  'True
         Caption         =   "XXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close program"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2450
      Left            =   75
      TabIndex        =   0
      Top             =   2895
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   4313
      _Version        =   393216
      FixedCols       =   0
      FillStyle       =   1
      ScrollBars      =   2
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label label3 
      AutoSize        =   -1  'True
      Caption         =   "Line with error in code of the program:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   75
      TabIndex        =   13
      Top             =   75
      Width           =   5355
   End
   Begin VB.Label lblStmt 
      Alignment       =   1  'Right Justify
      Caption         =   "XXXXXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   5520
      TabIndex        =   12
      Top             =   75
      Width           =   3015
   End
   Begin VB.Label lblProc 
      Alignment       =   1  'Right Justify
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   4800
      TabIndex        =   3
      Top             =   450
      Width           =   3735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Error has occurred in procedure:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   75
      TabIndex        =   2
      Top             =   450
      Width           =   4635
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Stack of executing procedure:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   75
      TabIndex        =   1
      Top             =   2520
      Width           =   8415
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sFile As String

Private Sub cmdContinue_Click()
        Me.Hide
        giAction_ = -1
End Sub

Private Sub Form_Load()
Dim FileName As String

    FileName = "error.log"
    
    sFile = App.Path + vbBackslash + FileName
    If mbIsDriveCDRoom Then sFile = "c:\error.log"
    
        With Me.MSFlexGrid1
             .ColWidth(0) = 1200
             .ColWidth(1) = 3000
             .TextMatrix(0, 0) = "Operator"
             .TextMatrix(0, 1) = "Name of procedure"
             .Refresh
        End With
End Sub

'Создание файла error.log
Private Sub cmdCreateFile_Click()
Dim iFile As Integer
Dim ErrText As String
'On Error GoTo err
    
    ErrText = "Description error in program " & App.ProductName & vbCrLf & _
              "====================================" & vbCrLf & vbCrLf & _
              "DateTime:" & vbTab & CStr(Date + time) & vbCrLf & _
              "Error Number:" & vbTab & lblErrCode.Caption & vbCrLf & _
              "Procedure:" & vbTab & lblProc.Caption & vbCrLf & _
              "Row with error:" & vbTab & lblStmt.Caption & vbCrLf & _
              "Description error:" & vbTab & lblErrDescr.Text & vbCrLf & vbCrLf & _
              "Listing executing procedure" & vbCrLf & _
              "---------------------" & vbCrLf & _
              StackText & _
              "Extended information" & vbCrLf & _
              "---------------------" & vbCrLf & _
              "Version of program:" & vbTab & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
              "Work path:" & vbTab & App.Path & vbCrLf & _
              "Name OS:" & vbTab & OSInfo.Name & vbCrLf & _
              "Version OS:" & vbTab & OSInfo.VerFull & vbCrLf & _
              "Build OS:" & vbTab & OSInfo.BuildNumber & vbCrLf & _
              "Other:" & vbTab & OSInfo.ServicePack & vbCrLf & _
              "====================================" & vbCrLf & vbCrLf

    iFile = FreeFile
    
    If PathFileExists(sFile) = 0 Then
        Open sFile For Output As #iFile
    Else
        Open sFile For Append As #iFile
    End If
    
    Print #iFile, ErrText
    Close #iFile
       
    MsgBox "Error.log saved: " & vbNewLine & sFile _
            & vbNewLine & vbNewLine & "Send error.log to author!"

Exit Sub
err:
    MsgBox "Error on create logfile: №" & err.Number & vbCrLf _
                & "Description: " & err.Description
End Sub

'Чтение стека процедур из таблицы
Private Function StackText() As String
Dim i As Integer, ii As Integer
i = 0
ii = 0
StackText = vbNullString
    With Me.MSFlexGrid1
        
        ii = .Rows - 1
        If (.TextMatrix(ii, 0) = vbNullString) And (.TextMatrix(ii, 1) = vbNullString) Then
            .Rows = .Rows - 1
        End If
        
        If .Rows = 1 Then
            StackText = vbCrLf
            Exit Function
        End If
        
        For i = 1 To .Rows - 1
            StackText = StackText & .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbCrLf
        Next i
        
        StackText = StackText & vbCrLf
    End With
End Function
    
Private Sub cmdEmail_Click()
    If PathFileExists(sFile) = 0 Then
        cmdCreateFile_Click
    End If
    Call ShellExecute(0, "Open", "mailto:Roman<roman-novosib@ngs.ru>" & _
                                "?Subject=Error_" & Replace$(App.ProductName," ", vbNullString) & "_" & App.Major & "." & App.Minor & "." & App.Revision, vbNullString, vbNullString, 1)
End Sub

Private Sub cmdExit_Click()
    Me.Hide
    giAction_ = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmError = Nothing
End Sub
