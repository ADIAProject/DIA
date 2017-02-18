VERSION 5.00
Begin VB.Form frmListHwid 
   Caption         =   "Список доступных драйверов"
   ClientHeight    =   5040
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
   ScaleHeight     =   5040
   ScaleWidth      =   11760
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.ListView lvFolders 
      Height          =   2895
      Left            =   60
      TabIndex        =   4
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
      VisualTheme     =   1
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
      Caption         =   "Группы драйверов:"
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
         BackColor       =   15783104
         Value           =   1
         Caption         =   "frmListHwid.frx":000C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkGrp4 
         Height          =   450
         Left            =   8880
         TabIndex        =   3
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
         BackColor       =   15783104
         Value           =   1
         Caption         =   "frmListHwid.frx":006A
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
         BackColor       =   15783104
         Caption         =   "frmListHwid.frx":00BC
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkGrp2 
         Height          =   450
         Left            =   2940
         TabIndex        =   1
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
         BackColor       =   15783104
         Value           =   1
         Caption         =   "frmListHwid.frx":00F6
         Transparent     =   -1  'True
      End
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   645
      Left            =   9840
      TabIndex        =   9
      Top             =   4300
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   8
      BackColor       =   12244692
      Caption         =   "ОК"
      CaptionEffects  =   0
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Default         =   -1  'True
      Height          =   645
      Left            =   7920
      TabIndex        =   8
      Top             =   4300
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   8
      BackColor       =   12244692
      Caption         =   "Выход"
      CaptionEffects  =   0
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdCheckAll 
      Height          =   435
      Left            =   60
      TabIndex        =   6
      Top             =   4500
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   8
      BackColor       =   12244692
      Caption         =   "Выделить всё"
      CaptionEffects  =   0
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdUnCheckAll 
      Height          =   435
      Left            =   2220
      TabIndex        =   7
      Top             =   4500
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   8
      BackColor       =   12244692
      Caption         =   "Снять выделение"
      CaptionEffects  =   0
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.LabelW lblInformation 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4000
      Visible         =   0   'False
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   661
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
      Caption         =   $"frmListHwid.frx":013C
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

' Минимальные размеры формы
Private lngFormWidthMin     As Long
Private lngFormHeightMin    As Long
Private strFormName         As String

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Get CaptionW
'! Description (Описание)  :   [Получение Caption-формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get CaptionW() As String
    Dim lngLenStr As Long
    
    lngLenStr = DefWindowProc(Me.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    CaptionW = Space$(lngLenStr)
    DefWindowProc Me.hWnd, WM_GETTEXT, Len(CaptionW) + 1, ByVal StrPtr(CaptionW)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Let CaptionW
'! Description (Описание)  :   [Изменение Caption-формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Let CaptionW(ByVal NewValue As String)
    DefWindowProc Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue & vbNullChar)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkGrp1_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkGrp1_Click()
    mbGrp1 = chkGrp1.Value
    miCurrentListCount = 0
    LoadListbyMode
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkGrp2_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkGrp2_Click()
    mbGrp2 = chkGrp2.Value
    miCurrentListCount = 0
    LoadListbyMode
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkGrp3_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkGrp3_Click()
    mbGrp3 = chkGrp3.Value
    miCurrentListCount = 0
    LoadListbyMode
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkGrp4_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkGrp4_Click()
    mbGrp4 = chkGrp4.Value
    miCurrentListCount = 0
    LoadListbyMode
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdCheckAll_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdCheckAll_Click()

    Dim ii As Integer

    With lvFolders.ListItems

        For ii = 1 To .count

            If Not .item(ii).Checked Then
                .item(ii).Checked = True
            End If

        Next

    End With

    FindCheckCountList
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdExit_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    mbSelectInstall = False
    Me.Hide
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdOK_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()

    If mbSelectInstall Then
        If FindCheckCountList Then
            If mbSelectInstall Then
                strPathDRPList = GetPathList
                mbCheckDRVOk = True
            End If

        Else

            If mbSelectInstall Then
                MsgBox "Not Selected. Window will be closed...", vbInformation + vbApplicationModal, strProductName
            End If
        End If
    End If

    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdUnCheckAll_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdUnCheckAll_Click()

    Dim ii As Integer

    With lvFolders.ListItems

        For ii = 1 To .count

            If .item(ii).Checked Then
                .item(ii).Checked = False
            End If

        Next

    End With

    FindCheckCountList
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CollectModeString
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
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

    ' Результирующая строка
    CollectModeString = strCmdStringDPInstTemp
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FindCheckCountList
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function FindCheckCountList() As Long

    Dim ii      As Integer
    Dim miCount As Integer

    For ii = 1 To lvFolders.ListItems.count

        If lvFolders.ListItems.item(ii).Checked Then
            miCount = miCount + 1
        End If

    Next

    With cmdOK
        If miCount Then
            If Not .Enabled Then
                .Enabled = True
            End If
        Else
            If mbSelectInstall Then
                If .Enabled Then
                    .Enabled = False
                End If
            End If
        End If
    End With

    FindCheckCountList = miCount
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FontCharsetChange
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub FontCharsetChange()

    ' Выставляем шрифт
    With Me.Font
        .Name = strFontOtherForm_Name
        .Size = lngFontOtherForm_Size
        .Charset = lngFont_Charset
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FormLoadAction
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub FormLoadAction()

    Dim ii As Integer

    miCurrentListCount = 0

    ' Локализация приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    mbGrp1 = chkGrp1.Value
    mbGrp2 = chkGrp2.Value
    mbGrp3 = chkGrp3.Value
    mbGrp4 = chkGrp4.Value
    mbCheckDRVOk = False

    If mbSelectInstall Then
        If mbGroupTask Then

            For ii = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)
                miCurrentListCount = miCurrentListCount + LoadList_Folders(CLng(arrCheckDP(0, ii)), False, CollectModeString)
            Next

        Else
            miCurrentListCount = LoadList_Folders(lngCurrentBtnIndex, False, CollectModeString)
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

        Me.CaptionW = strMeCaptionView & strSpace & lvFolders.ListItems.count & strSpace & strMessages(124) & strSpace & miCurrentListCount & ")"
    Else
        miCurrentListCount = LoadList_Folders(lngCurrentBtnIndex, True, CollectModeString)
        cmdExit.Visible = False
        cmdOK.Caption = strCmdOKCaption2
        lblInformation.Visible = False
        cmdCheckAll.Visible = False
        cmdUnCheckAll.Visible = False
        Me.CaptionW = strMeCaptionInstall & strSpace & lvFolders.ListItems.count & strSpace & strMessages(124) & strSpace & miCurrentListCount & ")"
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FormLoadDefaultParam
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
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
'! Procedure   (Функция)   :   Sub Form_Activate
'! Description (Описание)  :   []
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()
    lvFolders.SetFocus
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_KeyDown
'! Description (Описание)  :   [обработка нажатий клавиш клавиатуры]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        cmdExit_Click
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Load
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, strFormName, False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
        lngFormWidthMin = .Width
        lngFormHeightMin = .Height
    End With

    LoadIconImage2Object cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2Object cmdExit, "BTN_EXIT", strPathImageMainWork
    LoadIconImage2Object cmdCheckAll, "BTN_CHECKMARK", strPathImageMainWork
    LoadIconImage2Object cmdUnCheckAll, "BTN_UNCHECKMARK", strPathImageMainWork
    ' все остальные процедуры
    FormLoadAction
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_QueryUnload
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Cancel (Integer)
'                              UnloadMode (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' Выгружаем из памяти форму и другие компоненты
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
    Else
        Set frmListHwid = Nothing
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Resize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Resize()

    Dim miDeltaFrm  As Long
    Dim lngLVHeight As Long
    Dim lngLVWidht  As Long
    Dim lngLVTop    As Long

    On Error Resume Next

    With Me

        If .WindowState <> vbMinimized Then
            If IsWinVistaOrLater Then
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
            If mbSelectInstall Then
                lngLVHeight = ((cmdExit.Top - miDeltaFrm - 300)) - lngLVTop
            Else
                lngLVHeight = ((cmdExit.Top - miDeltaFrm - 100)) - lngLVTop
            End If
            
            lngLVWidht = ((.Width - miDeltaFrm)) - 18 * Screen.TwipsPerPixelX

            If Not (lvFolders Is Nothing) Then
                lvFolders.Move 60, lngLVTop, lngLVWidht, lngLVHeight
                lvFolders.Refresh
            End If

            cmdCheckAll.Top = cmdExit.Top + 200
            cmdUnCheckAll.Top = cmdCheckAll.Top
            cmdCheckAll.Left = miDeltaFrm + 60
            cmdUnCheckAll.Left = cmdCheckAll.Left + cmdCheckAll.Width + 200
            lblInformation.Top = cmdCheckAll.Top - 500
            lblInformation.Left = cmdCheckAll.Left
            lblInformation.Width = lngLVWidht
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetPathList
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function GetPathList() As String

    Dim ii              As Integer
    Dim iii             As Integer
    Dim strDevPathList  As String
    Dim strDevPathShort As String
    Dim strDevDPName    As String

    strDevPathList = vbNullString

    ' Если данного пути нет в списке, то добавляем
    With lvFolders.ListItems

        For ii = 1 To .count

            If .item(ii).Checked Then
                strDevPathShort = GetPathNameFromPath(.item(ii).SubItems(1))

                If mbGroupTask Then
                    strDevDPName = .item(ii).SubItems(8)

                    For iii = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)
                        strDevPathList = arrCheckDP(1, iii)

                        If StrComp(strDevDPName, frmMain.acmdPackFiles(arrCheckDP(0, iii)).Tag, vbTextCompare) = 0 Then
                            If InStr(1, strDevPathList, strDevPathShort, vbTextCompare) = 0 Then
                                AppendStr strDevPathList, strDevPathShort, strSpace
                            End If
                        End If

                        arrCheckDP(1, iii) = strDevPathList
                    Next

                Else

                    If InStr(1, strDevPathList, strDevPathShort, vbTextCompare) = 0 Then
                        AppendStr strDevPathList, strDevPathShort, strSpace
                    End If
                End If
            End If

        Next

    End With

    GetPathList = strDevPathList
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadListbyMode
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadListbyMode()

    Dim ii As Long

    If Not (lvFolders Is Nothing) Then
        lvFolders.ListItems.Clear
    End If

    lvFolders.Visible = False

    If mbSelectInstall Then
        If mbGroupTask Then

            For ii = LBound(arrCheckDP, 2) To UBound(arrCheckDP, 2)
                miCurrentListCount = miCurrentListCount + LoadList_Folders(CLng(arrCheckDP(0, ii)), False, CollectModeString)
            Next

        Else
            miCurrentListCount = LoadList_Folders(lngCurrentBtnIndex, False, CollectModeString)
        End If

        cmdCheckAll_Click
        Me.CaptionW = strMeCaptionView & strSpace & lvFolders.ListItems.count & strSpace & strMessages(124) & strSpace & miCurrentListCount & ")"
    Else
        miCurrentListCount = LoadList_Folders(lngCurrentBtnIndex, True, CollectModeString)
        Me.CaptionW = strMeCaptionInstall & strSpace & lvFolders.ListItems.count & strSpace & strMessages(124) & strSpace & miCurrentListCount & ")"
    End If

    FindCheckCountList
    lvFolders.Visible = True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function LoadList_Folders
'! Description (Описание)  :   [Построение спиcка ОС]
'! Parameters  (Переменные):   lngButtIndex (Long)
'                              mbViewed (Boolean = True)
'                              strMode (String = vbNullString)
'!--------------------------------------------------------------------------------
Private Function LoadList_Folders(ByVal lngButtIndex As Long, Optional ByVal mbViewed As Boolean = True, Optional ByVal strMode As String = vbNullString) As Long

    Dim strDevHwid          As String
    Dim strDevInfPath       As String
    Dim strDevDriverDB      As String
    Dim strDevDriverPrizn   As String
    Dim strDevDriverLocal   As String
    Dim strDevStatus        As String
    Dim strDevName          As String
    Dim miPreviousListCount As Long
    Dim miThisListCount     As Long
    Dim lngNumRow           As Long
    Dim strTemp_x()         As String
    Dim strTempLine_x()     As String
    Dim i_arr               As Long

    With lvFolders
        .Redraw = False
        .Checkboxes = Not mbViewed
        
        If .ColumnHeaders.count = 0 Then
            With .ColumnHeaders
                .Add 1, , strTableHwidHeader1, 165 * Screen.TwipsPerPixelX
                .Add 2, , strTableHwidHeader2, 100 * Screen.TwipsPerPixelX
                .Add 3, , strTableHwidHeader4, 90 * Screen.TwipsPerPixelX
                .Add 4, , strTableHwidHeader9, 20 * Screen.TwipsPerPixelX
                .Add 5, , strTableHwidHeader5, 90 * Screen.TwipsPerPixelX
                .Add 6, , strTableHwidHeader6, 30 * Screen.TwipsPerPixelX
                .Add 7, , strTableHwidHeader7, 350 * Screen.TwipsPerPixelX

                If mbGroupTask Then
                    .Add 8, , strTableHwidHeader8, 200 * Screen.TwipsPerPixelX
                End If
            End With
            
        End If

        lngNumRow = .ListItems.count
        miPreviousListCount = .ListItems.count
    End With

    strTemp_x = Split(arrTTip(lngButtIndex), vbNewLine)

    For i_arr = 0 To UBound(strTemp_x)
        If LenB(Trim$(strTemp_x(i_arr))) Then
            strTempLine_x = Split(strTemp_x(i_arr), " | ")
            
            If UBound(strTempLine_x) Then
                miThisListCount = miThisListCount + 1
                strDevHwid = Trim$(strTempLine_x(0))
                strDevInfPath = Trim$(strTempLine_x(1))
                strDevDriverDB = Trim$(strTempLine_x(2))
                strDevDriverPrizn = Trim$(strTempLine_x(3))
                strDevDriverLocal = Trim$(strTempLine_x(4))
                strDevStatus = Trim$(strTempLine_x(5))
                strDevName = Trim$(strTempLine_x(6))
    
                '*************************************************************
                If LenB(strMode) = 0 Then
    
                    With lvFolders.ListItems.Add(, , strDevHwid)
                        .SubItems(1) = strDevInfPath
                        .SubItems(2) = strDevDriverDB
                        .SubItems(3) = strDevDriverPrizn
                        .SubItems(4) = strDevDriverLocal
                        .SubItems(5) = strDevStatus
                        .SubItems(6) = strDevName
    
                        If mbGroupTask Then
                            .SubItems(7) = frmMain.acmdPackFiles(lngButtIndex).Tag
                        End If
    
                    End With
    
                    lngNumRow = lngNumRow + 1
                Else
                    '> - обновление
                    If InStr(strMode, ">") Then
                        If InStr(strDevDriverPrizn, ">") Then
        
                            With lvFolders.ListItems.Add(, , strDevHwid)
                                .SubItems(1) = strDevInfPath
                                .SubItems(2) = strDevDriverDB
                                .SubItems(3) = strDevDriverPrizn
                                .SubItems(4) = strDevDriverLocal
                                .SubItems(5) = strDevStatus
                                .SubItems(6) = strDevName
        
                                If mbGroupTask Then
                                    .SubItems(7) = frmMain.acmdPackFiles(lngButtIndex).Tag
                                End If
        
                            End With
        
                            lngNumRow = lngNumRow + 1
                        End If
                    End If
                    
                    '0 - неустановленные
                    If InStr(strMode, "0") Then
                        If InStr(strDevStatus, "0") Then
        
                            With lvFolders.ListItems.Add(, , strDevHwid)
                                .SubItems(1) = strDevInfPath
                                .SubItems(2) = strDevDriverDB
                                .SubItems(3) = strDevDriverPrizn
                                .SubItems(4) = strDevDriverLocal
                                .SubItems(5) = strDevStatus
                                .SubItems(6) = strDevName
        
                                If mbGroupTask Then
                                    .SubItems(7) = frmMain.acmdPackFiles(lngButtIndex).Tag
                                End If
        
                            End With
        
                            lngNumRow = lngNumRow + 1
                        End If
                    End If
                
                    '=? - установленные
                    If InStr(strMode, strRavno) Or InStr(strMode, strVopros) Then
                        If InStr(strDevDriverPrizn, strRavno) Or InStr(strDevDriverPrizn, strVopros) Then
        
                            With lvFolders.ListItems.Add(, , strDevHwid)
                                .SubItems(1) = strDevInfPath
                                .SubItems(2) = strDevDriverDB
                                .SubItems(3) = strDevDriverPrizn
                                .SubItems(4) = strDevDriverLocal
                                .SubItems(5) = strDevStatus
                                .SubItems(6) = strDevName
        
                                If mbGroupTask Then
                                    .SubItems(7) = frmMain.acmdPackFiles(lngButtIndex).Tag
                                End If
        
                            End With
        
                            lngNumRow = lngNumRow + 1
                        End If
                    End If
                    
                    '< - старее
                    If InStr(strMode, "<") Then
                        If InStr(strDevDriverPrizn, "<") Then
        
                            With lvFolders.ListItems.Add(, , strDevHwid)
                                .SubItems(1) = strDevInfPath
                                .SubItems(2) = strDevDriverDB
                                .SubItems(3) = strDevDriverPrizn
                                .SubItems(4) = strDevDriverLocal
                                .SubItems(5) = strDevStatus
                                .SubItems(6) = strDevName
        
                                If mbGroupTask Then
                                    .SubItems(7) = frmMain.acmdPackFiles(lngButtIndex).Tag
                                End If
        
                            End With
        
                            lngNumRow = lngNumRow + 1
                        End If
                    End If
                End If
            End If
            '*************************************************************
        End If

    Next i_arr
    
    With lvFolders.ColumnHeaders
        If .count Then
            If lvFolders.ListItems.count Then
                .item(1).AutoSize LvwColumnHeaderAutoSizeToItems
                .item(2).AutoSize LvwColumnHeaderAutoSizeToItems
                .item(3).AutoSize LvwColumnHeaderAutoSizeToItems
                .item(4).AutoSize LvwColumnHeaderAutoSizeToItems
                .item(5).AutoSize LvwColumnHeaderAutoSizeToItems
                .item(6).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(7).AutoSize LvwColumnHeaderAutoSizeToItems
        
                If mbGroupTask Then
                    .item(8).AutoSize LvwColumnHeaderAutoSizeToItems
                End If
            Else
                .item(1).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(2).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(3).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(4).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(5).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(6).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(7).AutoSize LvwColumnHeaderAutoSizeToHeader
        
                If mbGroupTask Then
                    .item(8).AutoSize LvwColumnHeaderAutoSizeToHeader
                End If
            
            End If
        End If
    End With
    
    lvFolders.Sorted = True
    lvFolders.Redraw = True
    LoadList_Folders = miThisListCount
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal strPathFile As String)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    strMeCaptionView = LocaliseString(strPathFile, strFormName, "frmListHwidView", Me.Caption)
    strMeCaptionInstall = LocaliseString(strPathFile, strFormName, "frmListHwidInstall", Me.Caption)
    lblInformation.Caption = LocaliseString(strPathFile, strFormName, "lblInformation", lblInformation.Caption)
    'Кнопки
    cmdCheckAll.Caption = LocaliseString(strPathFile, strFormName, "cmdCheckAll", cmdCheckAll.Caption)
    cmdUnCheckAll.Caption = LocaliseString(strPathFile, strFormName, "cmdUnCheckAll", cmdUnCheckAll.Caption)
    strCmdOKCaption1 = LocaliseString(strPathFile, strFormName, "cmdOKCaption1", "Установить")
    strCmdOKCaption2 = LocaliseString(strPathFile, strFormName, "cmdOKCaption2", "OK")
    strCmdOKCaption3 = LocaliseString(strPathFile, strFormName, "cmdOKCaption3", "Распаковать")
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
    frGroup.Caption = LocaliseString(strPathFile, strFormName, "frGroup", frGroup.Caption)
    chkGrp1.Caption = LocaliseString(strPathFile, strFormName, "chkGrp1", chkGrp1.Caption)
    chkGrp2.Caption = LocaliseString(strPathFile, strFormName, "chkGrp2", chkGrp2.Caption)
    chkGrp3.Caption = LocaliseString(strPathFile, strFormName, "chkGrp3", chkGrp3.Caption)
    chkGrp4.Caption = LocaliseString(strPathFile, strFormName, "chkGrp4", chkGrp4.Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvFolders_ColumnClick
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   ColumnHeader (LvwColumnHeader)
'!--------------------------------------------------------------------------------
Private Sub lvFolders_ColumnClick(ByVal ColumnHeader As LvwColumnHeader)

    Dim ii As Long

    lvFolders.Sorted = False
    lvFolders.SortKey = ColumnHeader.Index - 1

    If ComCtlsSupportLevel() >= 1 Then

        For ii = 1 To lvFolders.ColumnHeaders.count

            If ii <> ColumnHeader.Index Then
                lvFolders.ColumnHeaders(ii).SortArrow = LvwColumnHeaderSortArrowNone
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

        Next ii

        Select Case ColumnHeader.SortArrow

            Case LvwColumnHeaderSortArrowDown, LvwColumnHeaderSortArrowNone
                lvFolders.SortOrder = LvwSortOrderAscending

            Case LvwColumnHeaderSortArrowUp
                lvFolders.SortOrder = LvwSortOrderDescending
        End Select

        lvFolders.SelectedColumn = ColumnHeader
    Else

        For ii = 1 To lvFolders.ColumnHeaders.count

            If ii <> ColumnHeader.Index Then
                lvFolders.ColumnHeaders(ii).Icon = 0
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

        Next ii

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
'! Procedure   (Функция)   :   Sub lvFolders_ItemCheck
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Item (LvwListItem)
'                              Checked (Boolean)
'!--------------------------------------------------------------------------------
Private Sub lvFolders_ItemCheck(ByVal item As LvwListItem, ByVal Checked As Boolean)

    Dim ii As Integer

    If mbSelectInstall Then

        With lvFolders.ListItems

            If item.Checked Then

                For ii = 1 To .count

                    If StrComp(.item(ii).SubItems(1), item.SubItems(1), vbTextCompare) = 0 Then
                        .item(ii).Checked = True
                    End If

                Next

            Else

                For ii = 1 To .count

                    If StrComp(.item(ii).SubItems(1), item.SubItems(1), vbTextCompare) = 0 Then
                        .item(ii).Checked = False
                    End If

                Next

            End If

        End With

    End If

    FindCheckCountList
End Sub
