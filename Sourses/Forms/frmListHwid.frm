VERSION 5.00
Begin VB.Form frmListHwid 
   Caption         =   "������ ��������� ���������"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListHwid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   11760
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.ListView lvFolders 
      Height          =   2895
      Left            =   60
      TabIndex        =   8
      Top             =   1080
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   5106
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icons           =   "frmListHwid.frx":000C
      SmallIcons      =   "frmListHwid.frx":0038
      ColumnHeaderIcons=   "frmListHwid.frx":0064
      View            =   3
      Arrange         =   1
      AllowColumnReorder=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      LabelEdit       =   2
      HideSelection   =   0   'False
      ShowLabelTips   =   -1  'True
      TextBackground  =   1
   End
   Begin prjDIADBS.ctlJCFrames frGroup 
      Height          =   990
      Left            =   60
      Top             =   60
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   1746
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15783104
      FillColor       =   15783104
      TextBoxColor    =   11595760
      Style           =   3
      RoundedCorner   =   0   'False
      Caption         =   "������ ���������:"
      Alignment       =   0
      HeaderStyle     =   1
      Begin prjDIADBS.CheckBoxW chkGrp1 
         Height          =   450
         Left            =   60
         TabIndex        =   0
         Top             =   400
         Width           =   2600
         _ExtentX        =   4577
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Caption         =   "frmListHwid.frx":0090
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkGrp4 
         Height          =   450
         Left            =   8880
         TabIndex        =   1
         Top             =   400
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Caption         =   "frmListHwid.frx":00EE
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkGrp3 
         Height          =   450
         Left            =   5940
         TabIndex        =   2
         Top             =   400
         Width           =   2600
         _ExtentX        =   4577
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmListHwid.frx":0140
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkGrp2 
         Height          =   450
         Left            =   2940
         TabIndex        =   3
         Top             =   400
         Width           =   2600
         _ExtentX        =   4577
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Caption         =   "frmListHwid.frx":017A
         Transparent     =   -1  'True
      End
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   750
      Left            =   9840
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   8
      BackColor       =   12244692
      Caption         =   "��"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Height          =   750
      Left            =   7920
      TabIndex        =   5
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   8
      BackColor       =   12244692
      Caption         =   "�����"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdCheckAll 
      Height          =   360
      Left            =   60
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   8
      BackColor       =   12244692
      Caption         =   "�������� ��"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdUnCheckAll 
      Height          =   360
      Left            =   60
      TabIndex        =   7
      Top             =   4500
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   8
      BackColor       =   12244692
      Caption         =   "����� ���������"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.LabelW lblInformation 
      Height          =   675
      Left            =   2310
      TabIndex        =   9
      Top             =   4155
      Visible         =   0   'False
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   $"frmListHwid.frx":01C0
   End
End
Attribute VB_Name = "frmListHwid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strMeCaptionView    As String
Private strMeCaptionInstall As String
Private strCmdOKCaption1    As String
Private strCmdOKCaption2    As String
Private strCmdOKCaption3    As String
Private mbGrp1              As Boolean
Private mbGrp2              As Boolean
Private mbGrp3              As Boolean
Private mbGrp4              As Boolean
Private miCurrentListCount  As Long

' ����������� ������� �����
Private lngFormWidthMin     As Long
Private lngFormHeightMin    As Long
Private strFormName         As String

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub FontCharsetChange
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub FontCharsetChange()

    ' ���������� �����
    With Me.Font
        .Name = strFontOtherForm_Name
        .Size = lngFontOtherForm_Size
        .Charset = lngFont_Charset
    End With

    frGroup.Font.Charset = lngFont_Charset
    SetBtnFontProperties cmdExit
    SetBtnFontProperties cmdOK
    SetBtnFontProperties cmdCheckAll
    SetBtnFontProperties cmdUnCheckAll
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkGrp1_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkGrp1_Click()
    mbGrp1 = chkGrp1.Value
    miCurrentListCount = 0
    LoadListbyMode
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkGrp2_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkGrp2_Click()
    mbGrp2 = chkGrp2.Value
    miCurrentListCount = 0
    LoadListbyMode
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkGrp3_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkGrp3_Click()
    mbGrp3 = chkGrp3.Value
    miCurrentListCount = 0
    LoadListbyMode
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub chkGrp4_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub chkGrp4_Click()
    mbGrp4 = chkGrp4.Value
    miCurrentListCount = 0
    LoadListbyMode
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdCheckAll_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdCheckAll_Click()

    Dim i As Integer

    With lvFolders.ListItems

        For i = 1 To .Count

            If Not .Item(i).Checked Then
                .Item(i).Checked = True
            End If

        Next

    End With

    FindCheckCountList
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdExit_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    mbooSelectInstall = False
    Me.Hide
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdOK_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()

    If mbooSelectInstall Then
        If FindCheckCountList > 0 Then
            If mbooSelectInstall Then
                strPathDRPList = GetPathList
                mbCheckDRVOk = True
            End If

        Else

            If mbooSelectInstall Then
                MsgBox "Not Selected. Window will be closed...", vbInformation + vbApplicationModal, strProductName
            End If
        End If
    End If

    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdUnCheckAll_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdUnCheckAll_Click()

    Dim i As Integer

    With lvFolders.ListItems

        For i = 1 To .Count

            If .Item(i).Checked Then
                .Item(i).Checked = False
            End If

        Next

    End With

    FindCheckCountList
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function CollectModeString
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Function CollectModeString() As String

    Dim strCmdStringDPInstTemp As String

    If mbGrp1 Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "0"
    End If

    If mbGrp2 Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & ">"
    End If

    If mbGrp3 Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "=?"
    End If

    If mbGrp4 Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "<"
    End If

    ' �������������� ������
    CollectModeString = strCmdStringDPInstTemp
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function FindCheckCountList
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Function FindCheckCountList() As Long

    Dim i       As Integer
    Dim miCount As Integer

    For i = 1 To lvFolders.ListItems.Count

        If lvFolders.ListItems.Item(i).Checked Then
            miCount = miCount + 1
        End If

    Next

    If miCount > 0 Then

        With cmdOK

            If Not .Enabled Then
                .Enabled = True
                '.Refresh
            End If

        End With

        'CMDOK
    Else

        With cmdOK

            If mbooSelectInstall Then
                If .Enabled Then
                    .Enabled = False
                    '.Refresh
                End If
            End If

        End With

        'CMDOK
    End If

    FindCheckCountList = miCount
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_KeyDown
'! Description (��������)  :   [��������� ������� ������ ����������]
'! Parameters  (����������):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        cmdExit_Click
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_Load
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmListHwid", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
        lngFormWidthMin = .Width
        lngFormHeightMin = .Height
    End With

    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork
    LoadIconImage2BtnJC cmdCheckAll, "BTN_CHECKMARK", strPathImageMainWork
    LoadIconImage2BtnJC cmdUnCheckAll, "BTN_UNCHECKMARK", strPathImageMainWork
    ' ��� ��������� ���������
    FormLoadAction
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub FormLoadDefaultParam
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub FormLoadDefaultParam()
    miCurrentListCount = 0

    If Not (lvFolders Is Nothing) Then
        lvFolders.ColumnHeaders.Clear
        lvFolders.ListItems.Clear
    End If

    chkGrp1.Value = Checked
    chkGrp2.Value = Checked
    chkGrp3.Value = Unchecked
    chkGrp4.Value = Checked
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub FormLoadAction
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub FormLoadAction()

    Dim i As Integer

    miCurrentListCount = 0

    ' ����������� ����������
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' ���������� �����
        FontCharsetChange
    End If

    mbGrp1 = chkGrp1.Value
    mbGrp2 = chkGrp2.Value
    mbGrp3 = chkGrp3.Value
    mbGrp4 = chkGrp4.Value
    mbCheckDRVOk = False

    If mbooSelectInstall Then
        If mbGroupTask Then

            For i = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)
                miCurrentListCount = miCurrentListCount + LoadList_Folders(CLng(arrCheckDP(0, i)), False, CollectModeString)
            Next

        Else
            miCurrentListCount = LoadList_Folders(CurrentSelButtonIndex, False, CollectModeString)
        End If

        cmdCheckAll_Click
        lblInformation.Visible = True
        cmdCheckAll.Visible = True
        cmdUnCheckAll.Visible = True

        If mbOnlyUnpackDP Then
            cmdOK.Caption = strCmdOKCaption3
        Else
            cmdOK.Caption = strCmdOKCaption2
        End If

        Me.Caption = strMeCaptionView & " " & lvFolders.ListItems.Count & " " & strMessages(124) & " " & miCurrentListCount & ")"
    Else
        miCurrentListCount = LoadList_Folders(CurrentSelButtonIndex, True, CollectModeString)
        cmdExit.Visible = False
        cmdOK.Caption = strCmdOKCaption2
        lblInformation.Visible = False
        cmdCheckAll.Visible = False
        cmdUnCheckAll.Visible = False
        Me.Caption = strMeCaptionInstall & " " & lvFolders.ListItems.Count & " " & strMessages(124) & " " & miCurrentListCount & ")"
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function GetPathList
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Function GetPathList() As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim strDevPathList  As String
    Dim strDevPathShort As String
    Dim strDevDPName    As String

    strDevPathList = vbNullString

    ' ���� ������� ���� ��� � ������, �� ���������
    With lvFolders.ListItems

        For i = 1 To .Count

            If .Item(i).Checked Then
                strDevPathShort = .Item(i).SubItems(1)

                If mbGroupTask Then
                    strDevDPName = .Item(i).SubItems(8)

                    For ii = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)
                        strDevPathList = arrCheckDP(1, ii)

                        If StrComp(strDevDPName, frmMain.acmdPackFiles(arrCheckDP(0, ii)).Tag, vbTextCompare) = 0 Then
                            If InStr(1, strDevPathList, strDevPathShort, vbTextCompare) = 0 Then
                                strDevPathList = AppendStr(strDevPathList, strDevPathShort, " ")
                            End If
                        End If

                        arrCheckDP(1, ii) = strDevPathList
                    Next

                Else

                    If InStr(1, strDevPathList, strDevPathShort, vbTextCompare) = 0 Then
                        strDevPathList = AppendStr(strDevPathList, strDevPathShort, " ")
                    End If
                End If
            End If

        Next

    End With

    'LVFOLDERS
    GetPathList = strDevPathList
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function LoadList_Folders
'! Description (��������)  :   [���������� ���c�� ��]
'! Parameters  (����������):   lngButtIndex (Long)
'                              mbViewed (Boolean = True)
'                              strMode (String = vbNullString)
'!--------------------------------------------------------------------------------
Private Function LoadList_Folders(lngButtIndex As Long, Optional ByVal mbViewed As Boolean = True, Optional ByVal strMode As String = vbNullString) As Long

    Dim strDevHwid          As String
    Dim strDevFolder        As String
    Dim strDevInf           As String
    Dim strDevDriverDB      As String
    Dim strDevDriverPrizn   As String
    Dim strDevDriverLocal   As String
    Dim strDevStatus        As String
    Dim strDevName          As String
    Dim strSection          As String
    Dim miPreviousListCount As Long
    Dim miThisListCount     As Long
    Dim lngLVTop            As Long
    Dim lngLVHeight         As Long
    Dim lngLVWidht          As Long
    Dim lngNumRow           As Long

    With lvFolders
        .Redraw = False

        If mbViewed Then
            .Checkboxes = False
        Else
            .Checkboxes = True
        End If

        If .ColumnHeaders.Count = 0 Then
            .ColumnHeaders.Add 1, , strTableHwidHeader1, 165 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 2, , strTableHwidHeader2, 100 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 3, , strTableHwidHeader3, 75 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 4, , strTableHwidHeader4, 90 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 5, , strTableHwidHeader9, 20 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 6, , strTableHwidHeader5, 90 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 7, , strTableHwidHeader6, 30 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 8, , strTableHwidHeader7, 350 * Screen.TwipsPerPixelX

            If mbGroupTask Then
                .ColumnHeaders.Add 9, , strTableHwidHeader8, 200 * Screen.TwipsPerPixelX
            End If
        End If

    End With

    lngNumRow = lvFolders.ListItems.Count
    miPreviousListCount = lvFolders.ListItems.Count

    Dim strTemp_x()     As String
    Dim strTempLine_x() As String
    Dim i_arr           As Long

    Debug.Print arrTTip(lngButtIndex)
    strTemp_x = Split(arrTTip(lngButtIndex), vbNewLine)

    For i_arr = LBound(strTemp_x) To UBound(strTemp_x)
        strTempLine_x = Split(strTemp_x(i_arr), " | ")

        If LenB(Trim$(strTemp_x(i_arr))) Then
            miThisListCount = miThisListCount + 1
            strDevHwid = Trim$(strTempLine_x(0))
            strDevFolder = Trim$(strTempLine_x(1))
            strDevInf = Trim$(strTempLine_x(2))
            strDevDriverDB = Trim$(strTempLine_x(3))
            strDevDriverPrizn = Trim$(strTempLine_x(4))
            strDevDriverLocal = Trim$(strTempLine_x(5))
            strDevStatus = Trim$(strTempLine_x(6))
            strDevName = Trim$(strTempLine_x(7))

            '*************************************************************
            If LenB(strMode) = 0 Then

                With lvFolders.ListItems.Add(, , strDevHwid)
                    .SubItems(1) = strDevFolder
                    .SubItems(2) = strDevInf
                    .SubItems(3) = strDevDriverDB
                    .SubItems(4) = strDevDriverPrizn
                    .SubItems(5) = strDevDriverLocal
                    .SubItems(6) = strDevStatus
                    .SubItems(7) = strDevName

                    If mbGroupTask Then
                        .SubItems(8) = frmMain.acmdPackFiles(lngButtIndex).Tag
                    End If

                End With

                lngNumRow = lngNumRow + 1
            Else

                '> - ����������
                If InStr(strMode, ">") Then
                    If InStr(strDevDriverPrizn, ">") Then

                        With lvFolders.ListItems.Add(, , strDevHwid)
                            .SubItems(1) = strDevFolder
                            .SubItems(2) = strDevInf
                            .SubItems(3) = strDevDriverDB
                            .SubItems(4) = strDevDriverPrizn
                            .SubItems(5) = strDevDriverLocal
                            .SubItems(6) = strDevStatus
                            .SubItems(7) = strDevName

                            If mbGroupTask Then
                                .SubItems(8) = frmMain.acmdPackFiles(lngButtIndex).Tag
                            End If

                        End With

                        lngNumRow = lngNumRow + 1
                        GoTo NextListElement
                    End If
                End If

                '0 - ���������������
                If InStr(strMode, "0") Then
                    If InStr(strDevStatus, "0") Then

                        With lvFolders.ListItems.Add(, , strDevHwid)
                            .SubItems(1) = strDevFolder
                            .SubItems(2) = strDevInf
                            .SubItems(3) = strDevDriverDB
                            .SubItems(4) = strDevDriverPrizn
                            .SubItems(5) = strDevDriverLocal
                            .SubItems(6) = strDevStatus
                            .SubItems(7) = strDevName

                            If mbGroupTask Then
                                .SubItems(8) = frmMain.acmdPackFiles(lngButtIndex).Tag
                            End If

                        End With

                        lngNumRow = lngNumRow + 1
                        GoTo NextListElement
                    End If
                End If

                '=? - �������������
                If InStr(strMode, "=") Or InStr(strMode, "?") Then
                    If InStr(strDevDriverPrizn, "=") Or InStr(strDevDriverPrizn, "?") Then

                        With lvFolders.ListItems.Add(, , strDevHwid)
                            .SubItems(1) = strDevFolder
                            .SubItems(2) = strDevInf
                            .SubItems(3) = strDevDriverDB
                            .SubItems(4) = strDevDriverPrizn
                            .SubItems(5) = strDevDriverLocal
                            .SubItems(6) = strDevStatus
                            .SubItems(7) = strDevName

                            If mbGroupTask Then
                                .SubItems(8) = frmMain.acmdPackFiles(lngButtIndex).Tag
                            End If

                        End With

                        lngNumRow = lngNumRow + 1
                        GoTo NextListElement
                    End If
                End If

                '< - ������
                If InStr(strMode, "<") Then
                    If InStr(strDevDriverPrizn, "<") Then

                        With lvFolders.ListItems.Add(, , strDevHwid)
                            .SubItems(1) = strDevFolder
                            .SubItems(2) = strDevInf
                            .SubItems(3) = strDevDriverDB
                            .SubItems(4) = strDevDriverPrizn
                            .SubItems(5) = strDevDriverLocal
                            .SubItems(6) = strDevStatus
                            .SubItems(7) = strDevName

                            If mbGroupTask Then
                                .SubItems(8) = frmMain.acmdPackFiles(lngButtIndex).Tag
                            End If

                        End With

                        lngNumRow = lngNumRow + 1
                    End If
                End If
            End If

NextListElement:
            '*************************************************************
        End If

    Next i_arr

    lvFolders.Sorted = True
    lvFolders.Redraw = True
    LoadList_Folders = miThisListCount
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadListbyMode
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub LoadListbyMode()

    Dim i As Long

    If Not (lvFolders Is Nothing) Then
        lvFolders.ListItems.Clear
    End If

    lvFolders.Visible = False

    If mbooSelectInstall Then
        If mbGroupTask Then

            For i = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)
                miCurrentListCount = miCurrentListCount + LoadList_Folders(CLng(arrCheckDP(0, i)), False, CollectModeString)
            Next

        Else
            miCurrentListCount = LoadList_Folders(CurrentSelButtonIndex, False, CollectModeString)
        End If

        cmdCheckAll_Click
        Me.Caption = strMeCaptionView & " " & lvFolders.ListItems.Count & " " & strMessages(124) & " " & miCurrentListCount & ")"
    Else
        miCurrentListCount = LoadList_Folders(CurrentSelButtonIndex, True, CollectModeString)
        Me.Caption = strMeCaptionInstall & " " & lvFolders.ListItems.Count & " " & strMessages(124) & " " & miCurrentListCount & ")"
    End If

    FindCheckCountList
    lvFolders.Visible = True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Localise
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal StrPathFile As String)
    ' ���������� ����� ��������� (��������� ������ �� �� ��� ������� �� �������������� ������)
    FontCharsetChange
    ' �������� �����
    strMeCaptionView = LocaliseString(StrPathFile, strFormName, "frmListHwidView", Me.Caption)
    strMeCaptionInstall = LocaliseString(StrPathFile, strFormName, "frmListHwidInstall", Me.Caption)
    lblInformation.Caption = LocaliseString(StrPathFile, strFormName, "lblInformation", lblInformation.Caption)
    '������
    cmdCheckAll.Caption = LocaliseString(StrPathFile, strFormName, "cmdCheckAll", cmdCheckAll.Caption)
    cmdUnCheckAll.Caption = LocaliseString(StrPathFile, strFormName, "cmdUnCheckAll", cmdUnCheckAll.Caption)
    strCmdOKCaption1 = LocaliseString(StrPathFile, strFormName, "cmdOKCaption1", "����������")
    strCmdOKCaption2 = LocaliseString(StrPathFile, strFormName, "cmdOKCaption2", "OK")
    strCmdOKCaption3 = LocaliseString(StrPathFile, strFormName, "cmdOKCaption3", "�����������")
    cmdExit.Caption = LocaliseString(StrPathFile, strFormName, "cmdExit", cmdExit.Caption)
    frGroup.Caption = LocaliseString(StrPathFile, strFormName, "frGroup", frGroup.Caption)
    chkGrp1.Caption = LocaliseString(StrPathFile, strFormName, "chkGrp1", chkGrp1.Caption)
    chkGrp2.Caption = LocaliseString(StrPathFile, strFormName, "chkGrp2", chkGrp2.Caption)
    chkGrp3.Caption = LocaliseString(StrPathFile, strFormName, "chkGrp3", chkGrp3.Caption)
    chkGrp4.Caption = LocaliseString(StrPathFile, strFormName, "chkGrp4", chkGrp4.Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_QueryUnload
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   Cancel (Integer)
'                              UnloadMode (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' ��������� �� ������ ����� � ������ ����������
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
    Else
        Set frmListHwid = Nothing
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_Resize
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub Form_Resize()

    Dim miDeltaFrm  As Long
    Dim lngLVHeight As Long
    Dim lngLVWidht  As Long
    Dim lngLVTop    As Long

    On Error Resume Next

    With Me

        If .WindowState <> vbMinimized Then
            If OsCurrVersionStruct.VerFull >= "6.0" Then
                miDeltaFrm = 125
            Else

                If mbAppThemed Then
                    miDeltaFrm = 0
                Else
                    miDeltaFrm = 0
                End If
            End If

            If .Width < lngFormWidthMin Then
                .Width = lngFormWidthMin
                .Enabled = False
                .Enabled = True

                Exit Sub

            End If

            If .Height < lngFormHeightMin Then
                .Height = lngFormHeightMin
                .Enabled = False
                .Enabled = True

                Exit Sub

            End If

            cmdOK.Left = .Width - cmdOK.Width - 200 - miDeltaFrm
            cmdOK.Top = .Height - cmdOK.Height - 600 - miDeltaFrm
            cmdExit.Left = cmdOK.Left - cmdExit.Width - 110
            cmdExit.Top = cmdOK.Top
            lngLVTop = (frGroup.Top + frGroup.Height) + 5 * Screen.TwipsPerPixelX
            lngLVHeight = ((cmdExit.Top - miDeltaFrm - 100)) - lngLVTop
            lngLVWidht = ((.Width - miDeltaFrm)) - 18 * Screen.TwipsPerPixelX

            If Not (lvFolders Is Nothing) Then
                lvFolders.Move 60, lngLVTop, lngLVWidht, lngLVHeight
                lvFolders.Refresh
            End If

            cmdCheckAll.Top = cmdExit.Top
            cmdUnCheckAll.Top = cmdCheckAll.Top + cmdCheckAll.Height + 50
            cmdCheckAll.Left = miDeltaFrm + 60
            cmdUnCheckAll.Left = cmdCheckAll.Left
            lblInformation.Top = cmdExit.Top
            lblInformation.Width = cmdExit.Left - cmdCheckAll.Left - cmdCheckAll.Width - 200
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub lvFolders_ColumnClick
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   ColumnHeader (LvwColumnHeader)
'!--------------------------------------------------------------------------------
Private Sub lvFolders_ColumnClick(ByVal ColumnHeader As LvwColumnHeader)

    Dim i As Long

    lvFolders.Sorted = False
    lvFolders.SortKey = ColumnHeader.Index - 1

    If ComCtlsSupportLevel() >= 1 Then

        For i = 1 To lvFolders.ColumnHeaders.Count

            If i <> ColumnHeader.Index Then
                lvFolders.ColumnHeaders(i).SortArrow = LvwColumnHeaderSortArrowNone
            Else

                If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowNone Then
                    ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown
                Else

                    If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown Then
                        ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp
                    ElseIf ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp Then
                        ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown
                    End If
                End If
            End If

        Next i

        Select Case ColumnHeader.SortArrow

            Case LvwColumnHeaderSortArrowDown, LvwColumnHeaderSortArrowNone
                lvFolders.SortOrder = LvwSortOrderAscending

            Case LvwColumnHeaderSortArrowUp
                lvFolders.SortOrder = LvwSortOrderDescending
        End Select

        lvFolders.SelectedColumn = ColumnHeader
    Else

        For i = 1 To lvFolders.ColumnHeaders.Count

            If i <> ColumnHeader.Index Then
                lvFolders.ColumnHeaders(i).Icon = 0
            Else

                If ColumnHeader.Icon = 0 Then
                    ColumnHeader.Icon = 1
                Else

                    If ColumnHeader.Icon = 2 Then
                        ColumnHeader.Icon = 1
                    ElseIf ColumnHeader.Icon = 1 Then
                        ColumnHeader.Icon = 2
                    End If
                End If
            End If

        Next i

        Select Case ColumnHeader.Icon

            Case 1, 0
                lvFolders.SortOrder = LvwSortOrderAscending

            Case 2
                lvFolders.SortOrder = LvwSortOrderDescending
        End Select

    End If

    lvFolders.Sorted = True

    If Not lvFolders.SelectedItem Is Nothing Then lvFolders.SelectedItem.EnsureVisible
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub lvFolders_ItemCheck
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   Item (LvwListItem)
'                              Checked (Boolean)
'!--------------------------------------------------------------------------------
Private Sub lvFolders_ItemCheck(ByVal Item As LvwListItem, ByVal Checked As Boolean)

    Dim i As Integer

    If mbooSelectInstall Then

        With lvFolders.ListItems

            If Item.Checked Then

                For i = 1 To .Count

                    If StrComp(.Item(i).SubItems(1), Item.SubItems(1), vbTextCompare) = 0 Then
                        .Item(i).Checked = True
                    End If

                Next

            Else

                For i = 1 To .Count

                    If StrComp(.Item(i).SubItems(1), Item.SubItems(1), vbTextCompare) = 0 Then
                        .Item(i).Checked = False
                    End If

                Next

            End If

        End With

    End If

    FindCheckCountList
End Sub
