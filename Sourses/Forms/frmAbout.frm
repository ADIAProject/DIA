VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "� ���������..."
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9645
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.ctlJCbutton cmdHomePage 
      Height          =   650
      Left            =   7320
      TabIndex        =   4
      Top             =   5505
      Width           =   2205
      _ExtentX        =   3889
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
      Caption         =   "HomePage"
      CaptionEffects  =   0
      PictureAlign    =   0
      DropDownSymbol  =   6
      DropDownSeparator=   -1  'True
      DropDownEnable  =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdOsZoneNet 
      Height          =   650
      Left            =   4980
      TabIndex        =   3
      Top             =   5505
      Width           =   2205
      _ExtentX        =   3889
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
      Caption         =   "���������� �� OsZone.Net"
      CaptionEffects  =   0
      PictureAlign    =   0
   End
   Begin prjDIADBS.ctlJCbutton cmdCheckUpd 
      Height          =   650
      Left            =   1320
      TabIndex        =   5
      Top             =   6300
      Width           =   2205
      _ExtentX        =   3889
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
      Caption         =   "��������� ����������..."
      CaptionEffects  =   0
      PictureAlign    =   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdLicence 
      Height          =   650
      Left            =   2460
      TabIndex        =   2
      Top             =   5505
      Width           =   2205
      _ExtentX        =   3889
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
      Caption         =   "������������ ����������"
      CaptionEffects  =   0
      PictureAlign    =   0
   End
   Begin prjDIADBS.ctlJCbutton cmdDonate 
      Height          =   650
      Left            =   120
      TabIndex        =   1
      Top             =   5505
      Width           =   2200
      _ExtentX        =   3889
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
      Caption         =   "���������� ������"
      CaptionEffects  =   0
      PictureAlign    =   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Default         =   -1  'True
      Height          =   650
      Left            =   6120
      TabIndex        =   0
      Top             =   6300
      Width           =   2205
      _ExtentX        =   3889
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
      Caption         =   "�������"
      CaptionEffects  =   0
      PictureAlign    =   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton ctlAquaButton 
      Height          =   1995
      Left            =   75
      TabIndex        =   6
      Top             =   120
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   3519
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   10
      BackColor       =   16765357
      Caption         =   ""
      CaptionEffects  =   0
      PictureNormal   =   "frmAbout.frx":000C
      PictureShadow   =   -1  'True
   End
   Begin prjDIADBS.LabelW lblTranslator 
      Height          =   315
      Left            =   105
      TabIndex        =   10
      Top             =   2820
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   4
      BackStyle       =   0
      Caption         =   "������� ���������: �������� �����"
   End
   Begin prjDIADBS.LabelW lblThanks 
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   3180
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   4
      BackStyle       =   0
      Caption         =   "�������������:"
      WordWrap        =   0   'False
   End
   Begin prjDIADBS.LabelW lblAuthor 
      Height          =   375
      Left            =   105
      TabIndex        =   9
      Top             =   2520
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "����� ���������: �������� �����"
   End
   Begin prjDIADBS.LabelW lblInfo 
      Height          =   1095
      Left            =   2280
      TabIndex        =   8
      Top             =   1440
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   1931
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Caption         =   "�������� ���������"
   End
   Begin prjDIADBS.LabelW lblNameProg 
      Height          =   1305
      Left            =   2280
      TabIndex        =   7
      Top             =   45
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   2302
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Label1"
   End
   Begin prjDIADBS.LabelW lblMailTo 
      Height          =   240
      Left            =   105
      TabIndex        =   12
      Top             =   5160
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12582912
      MousePointer    =   4
      BackStyle       =   0
      Caption         =   "�������� E-mail ������ ���������"
   End
   Begin VB.Menu mnuContextMenu1 
      Caption         =   "����������� ���� 1"
      Begin VB.Menu mnuContextLink 
         Caption         =   "�������� ���� 1"
         Index           =   0
      End
      Begin VB.Menu mnuContextLink 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContextLink 
         Caption         =   "�������� ���� 2"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strTranslatorName As String
Private strTranslatorUrl  As String
Private strFormName       As String
Private strCreditList_x() As String
Private lngCurCredit      As Long

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Property Get CaptionW
'! Description (��������)  :   [��������� Caption-�����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Property Get CaptionW() As String
    Dim lngLenStr As Long
    
    lngLenStr = DefWindowProc(Me.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    CaptionW = Space$(lngLenStr)
    DefWindowProc Me.hWnd, WM_GETTEXT, Len(CaptionW) + 1, ByVal StrPtr(CaptionW)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Property Let CaptionW
'! Description (��������)  :   [��������� Caption-�����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Property Let CaptionW(ByVal NewValue As String)
    DefWindowProc Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue & vbNullChar)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdCheckUpd_Click
'! Description (��������)  :   [������ ����� �������� ����������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdCheckUpd_Click()
    CheckUpd False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdDonate_Click
'! Description (��������)  :   [������ ����� Donate]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdDonate_Click()
    frmDonate.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdExit_Click
'! Description (��������)  :   [����� �� �����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdHomePage_Click
'! Description (��������)  :   [������� �� �������� ��������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdHomePage_Click()
    RunUtilsShell strUrl_MainWWWSite, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdLicence_Click
'! Description (��������)  :   [����� ������������� ����������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdLicence_Click()
    frmLicence.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdOsZoneNet_Click
'! Description (��������)  :   [������� �� ����� OsZone.net]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdOsZoneNet_Click()
    RunUtilsShell strUrlOsZoneNetThread, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ctlAquaButton1_Click
'! Description (��������)  :   [������� �� ���� ���������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ctlAquaButton_Click()
    RunUtilsShell strUrl_MainWWWSite, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub FontCharsetChange
'! Description (��������)  :   [��������� ������ �����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub FontCharsetChange()

    ' ���������� �����
    With Me.Font
        .Name = strFontOtherForm_Name
        .Size = lngFontOtherForm_Size
        .Charset = lngFont_Charset
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_KeyDown
'! Description (��������)  :   [��������� ������� ������ ����������]
'! Parameters  (����������):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_Load
'! Description (��������)  :   [������� ���  �������� �����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, strFormName, False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    lblNameProg.Caption = strFrmMainCaptionTemp & vbNewLine & " v." & strProductVersion & vbNewLine & strFrmMainCaptionTempDate & strDateProgram & ")"

    LoadIconImage2Object cmdExit, "BTN_EXIT", strPathImageMainWork
    LoadIconImage2Object cmdDonate, "BTN_DONATE", strPathImageMainWork
    LoadIconImage2Object cmdCheckUpd, "BTN_UPDATE", strPathImageMainWork
    LoadIconImage2Object cmdHomePage, "BTN_HOME", strPathImageMainWork
    LoadIconImage2Object cmdOsZoneNet, "BTN_HOME", strPathImageMainWork
    LoadIconImage2Object cmdLicence, "BTN_LICENCE", strPathImageMainWork

    Select Case strPCLangCurrentID

        Case "0419"
            lblAuthor.Caption = "����� ���������: �������� ����� aka Romeo91"
            lblThanks(0).Caption = "��� �������������:"
        Case Else
            lblAuthor.Caption = "Author of the program: Goloveev Roman (Romeo91)"
            lblThanks(0).Caption = "My thanks:"
    End Select

    mnuContextMenu1.Enabled = False
    
    ' ����������� ����������
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' ���������� �����
        FontCharsetChange
    End If
    
    ' �������� ������ �������������� � ����� �� �����
    LoadThankYou
    
    ' ����������� ���� ��� ������
    cmdHomePage.SetPopupMenu mnuContextMenu1
    mnuContextMenu1.Enabled = True
    
    ' ��������� � �������� label
    With lblAuthor
        .MouseIcon = lblMailTo.MouseIcon
        .MousePointer = lblMailTo.MousePointer
        .ForeColor = lblMailTo.ForeColor
        .ToolTipText = strUrl_MainWWWSite
    End With
    lblMailTo.ToolTipText = "roman-novosib@ngs.ru"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_Unload
'! Description (��������)  :   [�������� �����]
'! Parameters  (����������):   Cancel (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    cmdHomePage.UnsetPopupMenu
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub GenerateThankyou
'! Description (��������)  :   [��������� ������ ������������� �� ������� �� ��������]
'!                              Idea from
'!                              Copyright �2001-2013 by Tanner Helland
'!                              http://www.tannerhelland.com/photodemon
'! Parameters  (����������):   thxText (String)
'                              creditURL (String = vbNullString)
'!--------------------------------------------------------------------------------
Private Sub GenerateThankyou(ByVal thxText As String, Optional ByVal creditURL As String = vbNullString)
    'Because I now have too many people to thank, it's necessary to split the list into multiple columns
    Dim columnLimit As Long
    Dim thxOffset   As Long

    'Generate a new label
    Load lblThanks(lngCurCredit)
    
    columnLimit = 7
    thxOffset = 750

    With lblThanks(lngCurCredit)

        If lngCurCredit = 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 30 + thxOffset
        ElseIf lngCurCredit < columnLimit Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 30 + thxOffset
        ElseIf lngCurCredit = columnLimit Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60 - (lblThanks(columnLimit - 1).Top - lblThanks(0).Top)
            .Left = lblThanks(0).Left + 2700 + thxOffset
        ElseIf lngCurCredit < columnLimit * 2 - 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 2700 + thxOffset
        ElseIf lngCurCredit = columnLimit * 2 - 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60 - (lblThanks(columnLimit * 2 - 2).Top - lblThanks(0).Top)
            .Left = lblThanks(0).Left + 5400 + thxOffset
        ElseIf lngCurCredit < columnLimit * 3 - 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 5400 + thxOffset
        ElseIf lngCurCredit = columnLimit * 3 - 1 Then
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60 - (lblThanks(columnLimit * 3 - 3).Top - lblThanks(0).Top)
            .Left = lblThanks(0).Left + 8100 + thxOffset
        Else
            .Top = lblThanks(lngCurCredit - 1).Top + lblThanks(lngCurCredit - 1).Height + 60
            .Left = lblThanks(0).Left + 8100 + thxOffset
        End If

        .Caption = thxText

        If LenB(creditURL) = 0 Then
            .MousePointer = vbDefault
        Else
            .Font.Underline = True
            .MouseIcon = lblMailTo.MouseIcon
            .MousePointer = lblMailTo.MousePointer
            .ForeColor = lblMailTo.ForeColor
            .ToolTipText = creditURL
        End If

        .Visible = True
    End With

    ReDim Preserve strCreditList_x(0 To lngCurCredit)

    strCreditList_x(lngCurCredit) = creditURL
    lngCurCredit = lngCurCredit + 1
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub lblAuthor_Click
'! Description (��������)  :   [������� �� ���� ������������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub lblAuthor_Click()
    RunUtilsShell strUrl_MainWWWSite, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub lblMailTo_MouseDown
'! Description (��������)  :   [������� ������ �� "��������� � �������������"]
'! Parameters  (����������):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub lblMailTo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim strSubject As String
    
    If Button = vbLeftButton Then
        strSubject = "My wishes for the program (" & App.ProductName & ")"
        ShellExecute Me.hWnd, vbNullString, "mailto:Romeo91<roman-novosib@ngs.ru>?Subject=" & Replace$(strSubject, strSpace, "%20"), vbNullString, "c:\", 1
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub lblThanks_Click
'! Description (��������)  :   [When a thank-you credit is clicked, launch the corresponding website]
'! Parameters  (����������):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub lblThanks_Click(Index As Integer)

    If LenB(strCreditList_x(Index)) Then
        RunUtilsShell strCreditList_x(Index), False
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub lblTranslator_MouseDown
'! Description (��������)  :   [������� �� ���� �����������, ��� �������� �����]
'! Parameters  (����������):   Button (Integer)
'                              Shift (Integer)
'                              X (Single)
'                              Y (Single)
'!--------------------------------------------------------------------------------
Private Sub lblTranslator_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If LenB(strTranslatorUrl) Then
        If Button = vbLeftButton Then
            RunUtilsShell strTranslatorUrl, False
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadThankYou
'! Description (��������)  :   [�������� ������ ��������������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub LoadThankYou()
    lngCurCredit = 1
    
    GenerateThankyou "SamLab", "http://driveroff.net/"
    GenerateThankyou "OSzone.net forum's users", "http://forum.oszone.net/forum-62.html"
    ' Replacement CommonControls (TextBoxW, ListView, ComboBoxW, ListBoxW, ProgressBar, ToolTip, ImageList, OptionButtonW,RichTextBox, CheckBoxW, LabelW, SpinBox)
    GenerateThankyou "Krool", "http://www.vbforums.com/showthread.php?698563-CommonControls-(Replacement-of-the-MS-common-controls)"
    'JCbutton
    GenerateThankyou "Juned Chhipa", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71482&lngWId=1"
    'jcFrames
    GenerateThankyou "Juan Carlos San Roman", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64261&lngWId=1"
    'clsmenuimage, ScrollControl
    GenerateThankyou "Leandro Ascierto", "http://leandroascierto.com/blog/clsmenuimage/"
    GenerateThankyou "VBnet and Randy Birch", "http://vbnet.mvps.org/"
    'cmdparsing
    GenerateThankyou "EliteXP Software Solutions", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=72018&lngWId=1"
    'ucPickBox' ucStatusBar
    GenerateThankyou "Paul R.Territos", "http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=63905&lngWId=1"
    '[VB6] Function Wait (non-freezing & non-CPU-intensive)
    GenerateThankyou "Bonnie West", "http://www.vbforums.com/showthread.php?700373-VB6-Shell-amp-Wait"
    'Team HomeWork
    ' Timed MessageBox
    GenerateThankyou "Anirudha Vengurlekar", "anirudhav@yahoo.com"
    'AnimateForm
    GenerateThankyou "Jim Jose", "jimjosev33@yahoo.com"
    'MD5
    GenerateThankyou "Marcin Kleczynski", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69092&lngWId=1"
    'HighlightActiveControl
    GenerateThankyou "Giorgio Brausi", "http://nuke.vbcorner.net/"
    'Unicode String Array Sorting Class - cBlizzard.cls
    GenerateThankyou "Rohan Edwards aka Rd", "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=72576&lngWId=1"
    'LoadThankYou and other idea
    GenerateThankyou "Tanner Helland", "http://photodemon.org/"
    ' SortDMArray
    GenerateThankyou "Ellis Dee"
    GenerateThankyou "Zhu JinYong"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadTranslator
'! Description (��������)  :   [�������� �������� � �����������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub LoadTranslator()

    Select Case strPCLangCurrentID

        Case "0419"
            lblTranslator.Caption = "������� ���������: " & strTranslatorName

        Case Else
            lblTranslator.Caption = "Translation of the program: " & strTranslatorName
    End Select

    If LenB(strTranslatorUrl) Then

        With lblTranslator
            .MouseIcon = lblMailTo.MouseIcon
            .MousePointer = lblMailTo.MousePointer
            .ForeColor = lblMailTo.ForeColor
            .ToolTipText = strTranslatorUrl
        End With

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Localise
'! Description (��������)  :   [�������� ������ ����������� ��� �����������]
'! Parameters  (����������):   strPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal strPathFile As String)
    ' ���������� ����� ��������� (��������� ������ �� �� ��� ������� �� �������������� ������)
    FontCharsetChange
    ' �������� �����
    Me.CaptionW = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    '  ����� �������� ������� ��� ������ Caption ���� � ���������� Unicode
    Call LocaliseMenu(strPathFile)
    '������
    cmdDonate.Caption = LocaliseString(strPathFile, strFormName, "cmdDonate", cmdDonate.Caption)
    cmdCheckUpd.Caption = LocaliseString(strPathFile, strFormName, "cmdCheckUpd", cmdCheckUpd.Caption)
    cmdLicence.Caption = LocaliseString(strPathFile, strFormName, "cmdLicence", cmdLicence.Caption)
    cmdHomePage.Caption = LocaliseString(strPathFile, strFormName, "cmdHomePage", cmdHomePage.Caption)
    cmdOsZoneNet.Caption = LocaliseString(strPathFile, strFormName, "cmdOsZoneNet", cmdOsZoneNet.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
    ' ������
    lblMailTo.Caption = LocaliseString(strPathFile, strFormName, "lblMailTo", lblMailTo.Caption)
    lblInfo.Caption = LocaliseString(strPathFile, strFormName, "lblInfo", lblInfo.Caption)
    ' ������� ���������
    strTranslatorName = LocaliseString(strPathFile, "Lang", "TranslatorName", lblTranslator.Caption)
    strTranslatorUrl = LocaliseString(strPathFile, "Lang", "TranslatorUrl", vbNullString)
    LoadTranslator
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LocaliseMenu
'! Description (��������)  :   [�������� ������ ����������� ��� ����]
'! Parameters  (����������):   strPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub LocaliseMenu(ByVal strPathFile As String)
    SetUniMenu -1, 0, -1, mnuContextMenu1, LocaliseString(strPathFile, strFormName, "cmdHomePage", cmdHomePage.Caption)
    SetUniMenu 0, 0, -1, mnuContextMenu1, LocaliseString(strPathFile, strFormName, "mnuContextLink1", mnuContextLink(0).Caption)
    SetUniMenu 0, 2, -1, mnuContextMenu1, LocaliseString(strPathFile, strFormName, "mnuContextLink2", mnuContextLink(2).Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub mnuContextLink_Click
'! Description (��������)  :   [������ �������� ��� ����������� ���� ������]
'! Parameters  (����������):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuContextLink_Click(Index As Integer)

    Select Case Index

        Case 0
            RunUtilsShell strUrl_MainWWWSite, False

        Case 2
            RunUtilsShell strUrl_MainWWWForum, False
    End Select
    
End Sub
