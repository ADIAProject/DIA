VERSION 5.00
Begin VB.Form frmEmulate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Режим эмуляции работы программы для другого ПК"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmulate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin prjDIADBS.ctlJCFrames frFile 
      Height          =   1395
      Left            =   60
      Top             =   60
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2461
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
      Caption         =   "Файл для эмуляции"
      TextBoxHeight   =   20
      GradientHeaderStyle=   1
      Begin prjDIADBS.ctlUcPickBox ucFilePath 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   900
         Width           =   7900
         _ExtentX        =   13944
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files *hwids*.txt|*hwids*.txt|All Files (*.*)"
         UseDialogText   =   0   'False
      End
      Begin prjDIADBS.LabelW lblInfo 
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   7900
         _ExtentX        =   13944
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Выберите файл для загрузки и укажите для какой операционной системы произвести эмуляцию работы программы"
      End
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   650
      Left            =   4380
      TabIndex        =   7
      Top             =   4320
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
      Enabled         =   0   'False
      BackColor       =   12244692
      Caption         =   "Загрузить файл"
      CaptionEffects  =   0
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Default         =   -1  'True
      Height          =   650
      Left            =   6420
      TabIndex        =   8
      Top             =   4320
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
      Caption         =   "Отмена"
      CaptionEffects  =   0
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCFrames frOS 
      Height          =   2715
      Left            =   60
      Top             =   1500
      Width           =   8175
      _ExtentX        =   14949
      _ExtentY        =   2461
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
      Caption         =   "Операционная система / Модель компьютера"
      TextBoxHeight   =   20
      GradientHeaderStyle=   1
      Begin prjDIADBS.TextBoxW txtPCModel 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "frmEmulate.frx":000C
         Locked          =   -1  'True
         CueBanner       =   "frmEmulate.frx":004C
      End
      Begin prjDIADBS.ComboBoxW cmbOS 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1140
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Text            =   "frmEmulate.frx":006C
         CueBanner       =   "frmEmulate.frx":008C
      End
      Begin prjDIADBS.CheckBoxW chk64bit 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   7900
         _ExtentX        =   13944
         _ExtentY        =   450
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
         Caption         =   "frmEmulate.frx":00AC
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkIsNotebook 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1860
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   450
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
         Caption         =   "frmEmulate.frx":00EA
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblOSInfo 
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   7900
         _ExtentX        =   13944
         _ExtentY        =   1296
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BackStyle       =   0
         Caption         =   $"frmEmulate.frx":011E
      End
   End
End
Attribute VB_Name = "frmEmulate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFilePath As String
Private strFormName As String

Public Property Get CaptionW() As String
    Dim lngLenStr As Long
    
    lngLenStr = DefWindowProc(Me.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    CaptionW = Space$(lngLenStr)
    DefWindowProc Me.hWnd, WM_GETTEXT, Len(CaptionW) + 1, ByVal StrPtr(CaptionW)
End Property

Public Property Let CaptionW(ByVal NewValue As String)
    DefWindowProc Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue & vbNullChar)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub EnablerCmdOK
'! Description (Описание)  :   [Активизация кнопки OK]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub EnablerCmdOK()

    If Not PathIsAFolder(ucFilePath.Path) Then
        If PathExists(ucFilePath.Path) Then
            If cmbOS.ListIndex >= 0 Then
                cmdOK.Enabled = True
            End If
        End If
    End If

End Sub

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
'! Procedure   (Функция)   :   Sub LoadAndParseFile
'! Description (Описание)  :   [Загрузка файла и переопределение массива]
'! Parameters  (Переменные):   strFilePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadAndParseFile(ByVal strFilePath As String)

    Dim arrFileStrings()  As String
    Dim ColumnByStrings() As String
    Dim i                 As Long
    Dim strContentFile    As String

    FileReadData strFilePath, strContentFile
    arrFileStrings = Split(strContentFile, vbNewLine)

    'Переопределяем основной массив с данными об устройствах компьютера
    ReDim arrHwidsLocal(UBound(arrFileStrings))

    For i = 0 To UBound(arrFileStrings)
        ColumnByStrings = Split(arrFileStrings(i), vbTab)

        With arrHwidsLocal(i)
            .HWID = ColumnByStrings(0)
            .DevName = ColumnByStrings(1)
            .Status = ColumnByStrings(2)
            .VerLocal = ColumnByStrings(3)
            .HWIDOrig = ColumnByStrings(4)
            .Provider = ColumnByStrings(5)
            .HWIDCompat = ColumnByStrings(6)
            .Description = ColumnByStrings(7)
            .PriznakSravnenia = ColumnByStrings(8)
            .InfSection = ColumnByStrings(9)
            .HWIDCutting = ColumnByStrings(10)
            .HWIDMatches = ColumnByStrings(11)
            .InfName = ColumnByStrings(12)
            .DPsList = vbNullString
            .DRVScore = 0
        End With

    Next i

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadDefaultParam
'! Description (Описание)  :   [заполнение списка на выделение]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadDefaultParam()

    Dim i            As Long
    Dim strVerClient As String

    ' Выставляем текущую версию ОС, анализом из списка
    If Not OSCurrVersionStruct.ClientOrServer Then
        strVerClient = "*" & OSCurrVersionStruct.VerFull & "*" & "Server"
    Else
        strVerClient = "*" & OSCurrVersionStruct.VerFull & "*"
    End If

    For i = 0 To cmbOS.ListCount - 1

        If MatchSpec(cmbOS.List(i), strVerClient) Then
            cmbOS.ListIndex = i

            Exit For

        End If

    Next i

    ' Выставляем текущую разрядность ОС
    chk64bit.Value = CBool(mbIsWin64)
    chkIsNotebook.Value = CBool(mbIsNotebok)
    txtPCModel = strCompModel
    ' Выставляем стартовый каталог
    ucFilePath.Path = strAppPathBackSL
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadListOS
'! Description (Описание)  :   [заполнение списка на выделение]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadListOS()

    With cmbOS
        .Clear
        .AddItem "(5.0) Windows 2000", 0
        .AddItem "(5.1) Windows XP", 1
        .AddItem "(6.0) Windows Vista", 2
        .AddItem "(6.1) Windows 7", 3
        .AddItem "(6.2) Windows 8", 4
        .AddItem "(6.3) Windows 8.1", 5
        .AddItem "(5.2) Windows Server 2003", 6
        .AddItem "(6.0) Windows Server 2008", 7
        .AddItem "(6.1) Windows Server 2008 R2", 8
        .AddItem "(6.2) Windows Server 2012", 9
        .AddItem "(6.3) Windows Server 2012 R2", 10
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal strPathFile As String)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.CaptionW = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    
    ' Frames
    frFile.Caption = LocaliseString(strPathFile, strFormName, "frFile", frFile.Caption)
    frOS.Caption = LocaliseString(strPathFile, strFormName, "frOS", frOS.Caption)
    ' Labels
    lblInfo.Caption = LocaliseString(strPathFile, strFormName, "lblInfo", lblInfo.Caption)
    lblOSInfo.Caption = LocaliseString(strPathFile, strFormName, "lblOSInfo", lblOSInfo.Caption)
    ' CheckBoxes
    chk64bit.Caption = LocaliseString(strPathFile, strFormName, "chk64bit", chk64bit.Caption)
    chkIsNotebook.Caption = LocaliseString(strPathFile, strFormName, "chkIsNotebook", chkIsNotebook.Caption)
    ' Buttons
    cmdOK.Caption = LocaliseString(strPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdExit_Click
'! Description (Описание)  :   [нажали выход]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdOK_Click
'! Description (Описание)  :   [нажали ок]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()

    Dim strFilePath As String

    strFilePath = ucFilePath.Path

    If LenB(strFilePath) Then
        'Переопределение массива данных о дайверах эмулируемой системы
        LoadAndParseFile strFilePath
        
        'Переопределение версии и разрядности системы для режима эмуляции
        mbIsWin64 = CBool(chk64bit.Value)
        strOSCurrentVersion = Mid$(cmbOS.Text, 2, 3)
        
        'Переопределение модели компьютера
        mbIsNotebok = CBool(chkIsNotebook.Value)
        strCompModel = txtPCModel
        
        ' А теперь обновляем статус всех пакетов
        frmMain.UpdateStatusButtonAll
        
        ' Обновить список неизвестных дров и описание для кнопки
        frmMain.LoadCmdViewAllDeviceCaption
        ChangeStatusBarText strMessages(114)
        Unload Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_KeyDown
'! Description (Описание)  :   [обработка нажатий клавиш клавиатуры]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Load
'! Description (Описание)  :   [обработка при загрузке формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()
    ' Устанавливаем картинки кнопок и убираем описание кнопок
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmUtilsEdit", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    LoadIconImage2Object cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2Object cmdExit, "BTN_EXIT", strPathImageMainWork

    ' Локализация приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    ' Загружаем список операционных систем
    LoadListOS
    LoadDefaultParam

    cmbOS.Enabled = False
    chk64bit.Enabled = False
    chkIsNotebook.Enabled = False
    txtPCModel.Enabled = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ParseFileName
'! Description (Описание)  :   [Parsing filename snap of the OS, and get OS parametrs]
'! Parameters  (Переменные):   strFilePath (String)
'!--------------------------------------------------------------------------------
Private Function ParseFileName(ByVal strFilePath As String) As Boolean
    
    Dim strParse_x()    As String
    Dim strTemp         As String
    Dim i               As Long
    Dim ii              As Long
    Dim mbIsServer      As Boolean
    
    strParse_x = Split(GetFileNameFromPath(strFilePath), "_")
        
    If UBound(strParse_x) >= 3 Then
        For i = 1 To UBound(strParse_x)
            '"hwids_%PCMODEL%-Notebook_" & strOSCurrentVersion & "-Server_%OSBIT%"
            
            Select Case i
                '%PCMODEL%-Notebook
                Case 1
                    strTemp = strParse_x(1)
                    If InStr(1, LCase$(strTemp), "notebook") Then
                        chkIsNotebook.Value = 1
                        txtPCModel = Replace$(strTemp, "-notebook", vbNullString, , , vbTextCompare)
                    Else
                        chkIsNotebook.Value = 0
                        txtPCModel = strTemp
                    End If
                    
                'strOSCurrentVersion-Server
                Case 2
                    strTemp = strParse_x(2)
                    If InStr(1, LCase$(strTemp), "server") Then
                        strTemp = Replace$(strTemp, "-server", vbNullString, , , vbTextCompare)
                        mbIsServer = True
                    End If
                    For ii = 0 To cmbOS.ListCount - 1

                        If InStr(cmbOS.List(ii), strTemp) Then
                            If mbIsServer Then
                                If InStr(1, cmbOS.List(ii), "server", vbTextCompare) = 0 Then
                                    Exit For
                                End If
                            End If
                            
                            cmbOS.ListIndex = ii
                            Exit For
                
                        End If
                
                    Next ii
                
                '%OSBIT%
                Case 3
                    strTemp = strParse_x(3)
                    If InStr(1, LCase$(strTemp), "x64") Then
                        chk64bit.Value = 1
                    Else
                        chk64bit.Value = 0
                    End If
            End Select
        Next i
        ParseFileName = True
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucFilePath_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucFilePath_Click()

    If ucFilePath.FileCount Then
        strFilePath = ucFilePath.FileName
    End If

    If LenB(strFilePath) Then
    
        cmbOS.Enabled = True
        chk64bit.Enabled = True
        chkIsNotebook.Enabled = True
        txtPCModel.Enabled = True
        
        If FileExists(strFilePath) Then
            ucFilePath.Path = strFilePath
            ' Парсинг имени файла и определение параметров ОС и компьютера
            If Not ParseFileName(strFilePath) Then
                MsgBox strMessages(156), vbInformation, strProductName
            End If
            ' активация кнопки старт
            EnablerCmdOK
        End If
    End If

End Sub

