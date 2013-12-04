VERSION 5.00
Begin VB.Form frmEmulate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Режим эмуляции работы программы для другого ПК"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjDIADBS.ctlUcPickBox ucFilePath 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   556
      UseAutoForeColor=   0   'False
      Color           =   0
      DefaultExt      =   ""
      DialogType      =   3
      Enabled         =   0   'False
      FileFlags       =   2621446
      Filters         =   "Supported files *hwids*.txt|*hwids*.txt|All Files (*.*)"
      ToolTipText3    =   "Click Here to Locate File"
   End
   Begin prjDIADBS.ComboBoxW cmbOS 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7935
      _ExtentX        =   13996
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
      Text            =   "frmEmulate.frx":0000
      CueBanner       =   "frmEmulate.frx":0020
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   750
      Left            =   4260
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1323
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12244692
      Caption         =   "Загрузить файл"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Height          =   750
      Left            =   6240
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12244692
      Caption         =   "Отмена"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin prjDIADBS.CheckBoxW chk64bit 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   7935
      _ExtentX        =   13996
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
      Caption         =   "frmEmulate.frx":0040
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.LabelW lblInfo 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   7935
      _ExtentX        =   13996
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
Attribute VB_Name = "frmEmulate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFilePath                     As String
Private strFormName                     As String


Private Sub FontCharsetChange()
' Выставляем шрифт
    With Me.Font
        .Name = strOtherForm_FontName
        .Size = lngOtherForm_FontSize
        .Charset = lngDialog_Charset
    End With

End Sub

'! -----------------------------------------------------------
'!  Функция     :  cmdExit_Click
'!  Переменные  :
'!  Описание    :  нажали выход
'! -----------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'! -----------------------------------------------------------
'!  Функция     :  LoadAndParseFile
'!  Переменные  :  ByVal strFilePath As String
'!  Описание    :  Загрузка файла и переопределение массива
'! -----------------------------------------------------------
Private Sub LoadAndParseFile(ByVal strFilePath As String)
Dim arrFileStrings()            As String
Dim ColumnByStrings()           As String
Dim i                           As Long
Dim strContentFile              As String
Dim objLoadedfFile              As TextStream

    Set objLoadedfFile = objFSO.OpenTextFile(strFilePath, ForReading, False, TristateUseDefault)
    strContentFile = objLoadedfFile.ReadAll
    objLoadedfFile.Close

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
            .DRVExist = 0
            .DPsList = vbNullString
            .DRVScore = 0
        End With
    Next i

End Sub
'! -----------------------------------------------------------
'!  Функция     :  cmdOK_Click
'!  Переменные  :
'!  Описание    :  нажали ок
'! -----------------------------------------------------------
Private Sub cmdOK_Click()
Dim strFilePath As String

    strFilePath = ucFilePath.Path
    
    If LenB(strFilePath) Then
    
        LoadAndParseFile strFilePath
        
        'Переопределение версии и разрядности системы для режима эмуляции
        mbIsWin64 = CBool(chk64bit.Value)
        strOsCurrentVersion = Mid$(cmbOS.Text, 2, 3)
        
        ' А теперь Обновляем статус всех пакетов
        frmMain.UpdateStatusButtonAll
        ' Обновить список неизвестных дров и описание для кнопки
        frmMain.LoadCmdViewAllDeviceCaption
        ChangeStatusTextAndDebug strMessages(114)
        Unload Me
    End If

End Sub

Private Sub ucFilePath_Click()
    If ucFilePath.FileCount > 0 Then
        strFilePath = ucFilePath.FileName
    End If

    If LenB(strFilePath) > 0 Then
        ucFilePath.Path = strFilePath
        ' активация кнопки старт
        EnablerCmdOK
    End If
End Sub

Private Sub Localise(ByVal StrPathFile As String)
' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.Caption = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
    ' Лэйблы
    lblInfo.Caption = LocaliseString(StrPathFile, strFormName, "lblInfo", lblInfo.Caption)
    chk64bit.Caption = LocaliseString(StrPathFile, strFormName, "chk64bit", chk64bit.Caption)
    'Кнопки
    cmdOK.Caption = LocaliseString(StrPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(StrPathFile, strFormName, "cmdExit", cmdExit.Caption)

End Sub

'! -----------------------------------------------------------
'!  Функция     :  Form_KeyDown
'!  Переменные  :  KeyCode As Integer, Shift As Integer
'!  Описание    :  обработка нажатий клавиш клавиатуры
'! -----------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  Form_Load
'!  Переменные  :
'!  Описание    :  обработка при загрузке формы
'! -----------------------------------------------------------
Private Sub Form_Load()
' Устанавливаем картинки кнопок и убираем описание кнопок
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmUtilsEdit", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork

    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If
    
    ' Загружаем список операционных систем
    LoadListOS
    LoadDefaultParam
    
End Sub

' Активизация кнопки OK
Private Sub EnablerCmdOK()
    If Not IsPathAFolder(ucFilePath.Path) Then
        If PathFileExists(ucFilePath.Path) = 1 Then
            If cmbOS.ListIndex >= 0 Then
                cmdOK.Enabled = True
            End If
        End If
    End If
End Sub

'заполнение списка на выделение
Private Sub LoadDefaultParam()
Dim i As Long
Dim strVerClient As String

' Выставляем текущую версию ОС, анализом из списка
    If Not OsCurrVersionStruct.ClientOrServer Then
        strVerClient = "*" & OsCurrVersionStruct.VerFull & "*" & "Server"
    Else
        strVerClient = "*" & OsCurrVersionStruct.VerFull & "*"
    End If
    
    For i = 0 To cmbOS.ListCount - 1
        If MatchSpec(cmbOS.List(i), strVerClient) Then
            cmbOS.ListIndex = i
            Exit For
        End If
    Next i
    
' Выставляем текущую разрядность ОС
    chk64bit.Value = CBool(mbIsWin64)
    
' Выставляем стартовый каталог
    ucFilePath.Path = strAppPathBackSL
End Sub

'заполнение списка на выделение
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

