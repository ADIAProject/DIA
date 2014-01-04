VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If False Then
Private CdlCancel, CdlBufferTooSmall, CdlInvalidFileName, CdlSubclassFailure, CdlMaxLessThanMin, CdlNoFonts, CdlInitFailure
Private CdlOFNReadOnly, CdlOFNOverwritePrompt, CdlOFNHideReadOnly, CdlOFNNoChangeDir, CdlOFNShowHelp, CdlOFNNoValidate, CdlOFNAllowMultiSelect, CdlOFNExtensionDifferent, CdlOFNPathMustExist, CdlOFNFileMustExist, CdlOFNCreatePrompt, CdlOFNShareAware, CdlOFNNoReadOnlyReturn, CdlOFNNoNetworkButton, CdlOFNExplorer, CdlOFNNoDereferenceLinks, CdlOFNDontAddToRecent, CdlOFNForcesShowHidden
Private CdlCCRGBInit, CdlCCFullOpen, CdlCCPreventFullOpen, CdlCCShowHelp, CdlCCSolidColor, CdlCCAnyColor
Private CdlCFScreenFonts, CdlCFPrinterFonts, CdlCFShowHelp, CdlCFEffects, CdlCFScriptsOnly, CdlCFNoVectorFonts, CdlCFLimitSize, CdlCFFixedPitchOnly, CdlCFForceFontExist, CdlCFScalableOnly, CdlCFTTOnly, CdlCFNoFaceSel, CdlCFNoStyleSel, CdlCFNoSizeSel, CdlCFSelectScript, CdlCFNoScriptSel, CdlCFNoVertFonts
Private CdlPDAllPages, CdlPDSelection, CdlPDPageNums, CdlPDNoSelection, CdlPDNoPageNums, CdlPDCollate, CdlPDPrintToFile, CdlPDNoWarning, CdlPDReturnDC, CdlPDReturnIC, CdlPDReturnDefault, CdlPDUseDevModeCopiesAndCollate, CdlPDDisablePrintToFile, CdlPDCurrentPage, CdlPDHidePrintToFile, CdlPDNoCurrentPage
Private CdlPDResultCancel, CdlPDResultPrint, CdlPDResultApply
Private CdlBIFReturnOnlyFSDirs, CdlBIFDontGoBelowDomain, CdlBIFReturnFSAncestors, CdlBIFEditBox, CdlBIFNewDialogStyle, CdlBIFBrowseIncludeURLs, CdlBIFUseNewUI, CdlBIFUAHint, CdlBIFNoNewFolderButton, CdlBIFNoTranslateTargets, CdlBIFBrowseForComputer, CdlBIFBrowseForPrinter, CdlBIFBrowseIncludeFiles, CdlBIFShareable, CdlBIFBrowseFileJunctions
#End If
Private Const FNERR_BUFFERTOOSMALL As Long = &H3003
Private Const FNERR_INVALIDFILENAME As Long = &H3002
Private Const FNERR_SUBCLASSFAILURE As Long = &H3001
Private Const CFERR_MAXLESSTHANMIN As Long = &H2002
Private Const CFERR_NOFONTS As Long = &H2001
Public Enum CdlErrorConstants
CdlCancel = 32755
CdlBufferTooSmall = 20476
CdlInvalidFileName = 20477
CdlSubclassFailure = 20478
CdlMaxLessThanMin = 24573
CdlNoFonts = 24574
CdlInitFailure = 28665
End Enum
Private Const OFN_READONLY As Long = &H1
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_SHOWHELP As Long = &H10
Private Const OFN_NOVALIDATE As Long = &H100
Private Const OFN_ALLOWMULTISELECT As Long = &H200
Private Const OFN_EXTENSIONDIFFERENT As Long = &H400
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_CREATEPROMPT As Long = &H2000
Private Const OFN_SHAREAWARE As Long = &H4000
Private Const OFN_NOREADONLYRETURN As Long = &H8000&
Private Const OFN_NONETWORKBUTTON As Long = &H20000
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_NODEREFERENCELINKS As Long = &H100000
Private Const OFN_DONTADDTORECENT As Long = &H2000000
Private Const OFN_FORCESHOWHIDDEN As Long = &H10000000
Public Enum CdlOFNConstants
CdlOFNReadOnly = OFN_READONLY
CdlOFNOverwritePrompt = OFN_OVERWRITEPROMPT
CdlOFNHideReadOnly = OFN_HIDEREADONLY
CdlOFNNoChangeDir = OFN_NOCHANGEDIR
CdlOFNShowHelp = OFN_SHOWHELP
CdlOFNNoValidate = OFN_NOVALIDATE
CdlOFNAllowMultiSelect = OFN_ALLOWMULTISELECT
CdlOFNExtensionDifferent = OFN_EXTENSIONDIFFERENT
CdlOFNPathMustExist = OFN_PATHMUSTEXIST
CdlOFNFileMustExist = OFN_FILEMUSTEXIST
CdlOFNCreatePrompt = OFN_CREATEPROMPT
CdlOFNShareAware = OFN_SHAREAWARE
CdlOFNNoReadOnlyReturn = OFN_NOREADONLYRETURN
CdlOFNNoNetworkButton = OFN_NONETWORKBUTTON
CdlOFNExplorer = OFN_EXPLORER
CdlOFNNoDereferenceLinks = OFN_NODEREFERENCELINKS
CdlOFNDontAddToRecent = OFN_DONTADDTORECENT
CdlOFNForcesShowHidden = OFN_FORCESHOWHIDDEN
End Enum
Private Const CC_RGBINIT As Long = &H1
Private Const CC_FULLOPEN As Long = &H2
Private Const CC_PREVENTFULLOPEN As Long = &H4
Private Const CC_SHOWHELP As Long = &H8
Private Const CC_SOLIDCOLOR As Long = &H80
Private Const CC_ANYCOLOR As Long = &H100
Public Enum CdlCCConstants
CdlCCRGBInit = CC_RGBINIT
CdlCCFullOpen = CC_FULLOPEN
CdlCCPreventFullOpen = CC_PREVENTFULLOPEN
CdlCCShowHelp = CC_SHOWHELP
CdlCCSolidColor = CC_SOLIDCOLOR
CdlCCAnyColor = CC_ANYCOLOR
End Enum
Private Const CF_SCREENFONTS As Long = &H1
Private Const CF_PRINTERFONTS As Long = &H2
Private Const CF_SHOWHELP As Long = &H4
Private Const CF_EFFECTS As Long = &H100
Private Const CF_SCRIPTSONLY As Long = &H400
Private Const CF_NOVECTORFONTS As Long = &H800
Private Const CF_LIMITSIZE As Long = &H2000
Private Const CF_FIXEDPITCHONLY As Long = &H4000
Private Const CF_FORCEFONTEXIST As Long = &H10000
Private Const CF_SCALABLEONLY As Long = &H20000
Private Const CF_TTONLY As Long = &H40000
Private Const CF_NOFACESEL As Long = &H80000
Private Const CF_NOSTYLESEL As Long = &H100000
Private Const CF_NOSIZESEL As Long = &H200000
Private Const CF_SELECTSCRIPT As Long = &H400000
Private Const CF_NOSCRIPTSEL As Long = &H800000
Private Const CF_NOVERTFONTS As Long = &H1000000
Public Enum CdlCFConstants
CdlCFScreenFonts = CF_SCREENFONTS
CdlCFPrinterFonts = CF_PRINTERFONTS
CdlCFShowHelp = CF_SHOWHELP
CdlCFEffects = CF_EFFECTS
CdlCFScriptsOnly = CF_SCRIPTSONLY
CdlCFNoVectorFonts = CF_NOVECTORFONTS
CdlCFLimitSize = CF_LIMITSIZE
CdlCFFixedPitchOnly = CF_FIXEDPITCHONLY
CdlCFForceFontExist = CF_FORCEFONTEXIST
CdlCFScalableOnly = CF_SCALABLEONLY
CdlCFTTOnly = CF_TTONLY
CdlCFNoFaceSel = CF_NOFACESEL
CdlCFNoStyleSel = CF_NOSTYLESEL
CdlCFNoSizeSel = CF_NOSIZESEL
CdlCFSelectScript = CF_SELECTSCRIPT
CdlCFNoScriptSel = CF_NOSCRIPTSEL
CdlCFNoVertFonts = CF_NOVERTFONTS
End Enum
Private Const PD_ALLPAGES As Long = &H0
Private Const PD_SELECTION As Long = &H1
Private Const PD_PAGENUMS As Long = &H2
Private Const PD_NOSELECTION As Long = &H4
Private Const PD_NOPAGENUMS As Long = &H8
Private Const PD_COLLATE As Long = &H10
Private Const PD_PRINTTOFILE As Long = &H20
Private Const PD_NOWARNING As Long = &H80
Private Const PD_RETURNDC As Long = &H100
Private Const PD_RETURNIC As Long = &H200
Private Const PD_RETURNDEFAULT As Long = &H400
Private Const PD_USEDEVMODECOPIESANDCOLLATE As Long = &H40000
Private Const PD_DISABLEPRINTTOFILE As Long = &H80000
Private Const PD_CURRENTPAGE As Long = &H400000
Private Const PD_HIDEPRINTTOFILE As Long = &H100000
Private Const PD_NOCURRENTPAGE As Long = &H800000
Public Enum CdlPDConstants
CdlPDAllPages = PD_ALLPAGES
CdlPDSelection = PD_SELECTION
CdlPDPageNums = PD_PAGENUMS
CdlPDNoSelection = PD_NOSELECTION
CdlPDNoPageNums = PD_NOPAGENUMS
CdlPDCollate = PD_COLLATE
CdlPDPrintToFile = PD_PRINTTOFILE
CdlPDNoWarning = PD_NOWARNING
CdlPDReturnDC = PD_RETURNDC
CdlPDReturnIC = PD_RETURNIC
CdlPDReturnDefault = PD_RETURNDEFAULT
CdlPDUseDevModeCopiesAndCollate = PD_USEDEVMODECOPIESANDCOLLATE
CdlPDDisablePrintToFile = PD_DISABLEPRINTTOFILE
CdlPDCurrentPage = PD_CURRENTPAGE
CdlPDHidePrintToFile = PD_HIDEPRINTTOFILE
CdlPDNoCurrentPage = PD_NOCURRENTPAGE
End Enum
Private Const PD_RESULT_CANCEL As Long = &H0
Private Const PD_RESULT_PRINT As Long = &H1
Private Const PD_RESULT_APPLY As Long = &H2
Public Enum CdlPDResultConstants
CdlPDResultCancel = PD_RESULT_CANCEL
CdlPDResultPrint = PD_RESULT_PRINT
CdlPDResultApply = PD_RESULT_APPLY
End Enum
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_EDITBOX As Long = &H10
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_BROWSEINCLUDEURLS As Long = &H80
Private Const BIF_USENEWUI As Long = BIF_EDITBOX Or BIF_NEWDIALOGSTYLE
Private Const BIF_UAHINT As Long = &H100
Private Const BIF_NONEWFOLDERBUTTON As Long = &H200
Private Const BIF_NOTRANSLATETARGETS As Long = &H400
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const BIF_SHAREABLE As Long = &H8000&
Private Const BIF_BROWSEFILEJUNCTIONS As Long = &H10000
Public Enum CdlBIFConstants
CdlBIFReturnOnlyFSDirs = BIF_RETURNONLYFSDIRS
CdlBIFDontGoBelowDomain = BIF_DONTGOBELOWDOMAIN
CdlBIFReturnFSAncestors = BIF_RETURNFSANCESTORS
CdlBIFEditBox = BIF_EDITBOX
CdlBIFNewDialogStyle = BIF_NEWDIALOGSTYLE
CdlBIFBrowseIncludeURLs = BIF_BROWSEINCLUDEURLS
CdlBIFUseNewUI = BIF_USENEWUI
CdlBIFUAHint = BIF_UAHINT
CdlBIFNoNewFolderButton = BIF_NONEWFOLDERBUTTON
CdlBIFNoTranslateTargets = BIF_NOTRANSLATETARGETS
CdlBIFBrowseForComputer = BIF_BROWSEFORCOMPUTER
CdlBIFBrowseForPrinter = BIF_BROWSEFORPRINTER
CdlBIFBrowseIncludeFiles = BIF_BROWSEINCLUDEFILES
CdlBIFShareable = BIF_SHAREABLE
CdlBIFBrowseFileJunctions = BIF_BROWSEFILEJUNCTIONS
End Enum
Private Type OPENFILENAME
lStructSize As Long
hWndOwner As Long
hInstance As Long
lpstrFilter As Long
lpstrCustomFilter As Long
nMaxCustFilter As Long
nFilterIndex As Long
lpstrFile As Long
nMaxFile As Long
lpstrFileTitle As Long
nMaxFileTitle As Long
lpstrInitialDir As Long
lpstrTitle As Long
Flags As Long
nFileOffset As Integer
nFileExtension As Integer
lpstrDefExt As Long
lCustData As Long
lpfnHook As Long
lpTemplateName As Long
End Type
Private Type TCHOOSECOLOR
lStructSize As Long
hWndOwner As Long
hInstance As Long
RGBResult As Long
lpCustColors As Long
Flags As Long
lCustData As Long
lpfnHook As Long
lpTemplateName As Long
End Type
Private Type TCHOOSEFONT
lStructSize As Long
hWndOwner As Long
hDC As Long
lpLogFont As Long
iPointSize As Long
Flags As Long
RGBColor As Long
lCustData As Long
lpfnHook As Long
lpTemplateName As Long
hInstance As Long
lpszStyle As Long
nFontType As Integer
nSizeMin As Long
nSizeMax As Long
End Type
Private Const LF_FACESIZE As Long = 32
Private Const FW_NORMAL As Long = 400
Private Const FW_BOLD As Long = 700
Private Const DEFAULT_QUALITY As Long = 0
Private Type LOGFONT
LFHeight As Long
LFWidth As Long
LFEscapement As Long
LFOrientation As Long
LFWeight As Long
LFItalic As Byte
LFUnderline As Byte
LFStrikeOut As Byte
LFCharset As Byte
LFOutPrecision As Byte
LFClipPrecision As Byte
LFQuality As Byte
LFPitchAndFamily As Byte
LFFaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
End Type
Private Type PRINTPAGERANGE
nFromPage As Long
nToPage As Long
End Type
Private Type PRINTDLGEX
lStructSize As Long
hWndOwner As Long
hDevMode As Long
hDevNames As Long
hDC As Long
Flags As Long
Flags2 As Long
ExclusionFlags As Long
nPageRanges As Long
nMaxPageRanges As Long
lpPageRanges As Long
nMinPage As Long
nMaxPage As Long
nCopies As Long
hInstance As Long
lpPrintTemplateName As String
lpCallback As Long
nPropertyPages As Long
lphPropertyPages As Long
nStartPage As Long
dwResultAction As Long
End Type
Private Const CCHDEVNAMESEXTRA As Long = 100
Private Type DEVNAMES
wDriverOffset As Integer
wDeviceOffset As Integer
wOutputOffset As Integer
wDefault As Integer
wExtra As String * CCHDEVNAMESEXTRA
End Type
Private Const CCHDEVICENAME As Long = 32
Private Const CCHFORMNAME As Long = 32
Private Type DEVMODE
DMDeviceName As String * CCHDEVICENAME
DMSpecVersion As Integer
DMDriverVersion As Integer
DMSize As Integer
DMDriverExtra As Integer
DMFields As Long
DMOrientation As Integer
DMPaperSize As Integer
DMPaperLength As Integer
DMPaperWidth As Integer
DMScale As Integer
DMCopies As Integer
DMDefaultSource As Integer
DMPrintQuality As Integer
DMColor As Integer
DMDuplex As Integer
DMYResolution As Integer
DMTTOption As Integer
DMCollate As Integer
DMFormName As String * CCHFORMNAME
DMUnusedPadding As Integer
DMBitsPerPel As Integer
DMPelsWidth As Long
DMPelsHeight As Long
DMDisplayFlags As Long
DMDisplayFrequency As Long
End Type
Private Type BROWSEINFO
hWndOwner As Long
pIDLRoot As Long
pszDisplayName As Long
lpszTitle As Long
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Long
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameW" (ByRef pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameW" (ByRef pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32" Alias "ChooseColorW" (ByRef lpChooseColor As TCHOOSECOLOR) As Long
Private Declare Function ChooseFont Lib "comdlg32" Alias "ChooseFontW" (ByRef lpcf As TCHOOSEFONT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderW" (ByRef lpBI As BROWSEINFO) As Long
Private Declare Function SHGetFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByRef pIDList As Long) As Long
Private Declare Function ILCreateFromPath Lib "shell32" (ByVal lpszPath As Long) As Long
Private Declare Function ILCreateFromPath_W2K Lib "shell32" Alias "#157" (ByVal lpszPath As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" (ByVal pIDList As Long, ByVal lpBuffer As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PrintDialogEx Lib "comdlg32" Alias "PrintDlgExA" (ByRef pPD As PRINTDLGEX) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const MAX_PATH As Long = 260
Private PropCancelError As Boolean
Private PropTag As String
Private PropDC As Long
Private PropFlags As Long
Private PropDialogTitle As String
Private PropMaxFileSize As Long
Private PropFileName As String, PropFileTitle As String
Private PropFileOffset As Integer
Private PropFilter As String, PropFilterIndex As Long
Private PropInitDir As String
Private PropDefaultExt As String
Private PropColor As Long
Private PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropMin As Long, PropMax As Long
Private PropFromPage As Long, PropToPage As Long

Private Sub Class_Initialize()
PropMaxFileSize = MAX_PATH
Set PropFont = New StdFont
End Sub

Public Property Get CancelError() As Boolean
Attribute CancelError.VB_Description = "Indicates whether an error is generated when the user chooses the cancel button."
CancelError = PropCancelError
End Property

Public Property Let CancelError(ByVal Value As Boolean)
PropCancelError = Value
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = PropTag
End Property

Public Property Let Tag(ByVal Value As String)
PropTag = Value
End Property

Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle to the object's device context."
hDC = PropDC
End Property

Public Property Let hDC(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Flags() As Long
Attribute Flags.VB_Description = "Returns/sets the options for a dialog box."
Flags = PropFlags
End Property

Public Property Let Flags(ByVal Value As Long)
PropFlags = Value
End Property

Public Property Get DialogTitle() As String
Attribute DialogTitle.VB_Description = "Sets the string displayed in the title bar of the dialog box."
DialogTitle = PropDialogTitle
End Property

Public Property Let DialogTitle(ByVal Value As String)
PropDialogTitle = Value
End Property

Public Property Get MaxFileSize() As Long
Attribute MaxFileSize.VB_Description = "Returns/sets the maximum size of the file name opened."
MaxFileSize = PropMaxFileSize
End Property

Public Property Let MaxFileSize(ByVal Value As Long)
If Value < 1 Then Err.Raise 380
PropMaxFileSize = Value
End Property

Public Property Get FileName() As String
Attribute FileName.VB_Description = "Returns/sets the path and file name of a selected file."
FileName = PropFileName
End Property

Public Property Let FileName(ByVal Value As String)
PropFileName = Value
End Property

Public Property Get FileTitle() As String
Attribute FileTitle.VB_Description = "Returns the file name (without the path) of the file to open or save."
FileTitle = PropFileTitle
End Property

Public Property Let FileTitle(ByVal Value As String)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get FileOffset() As Long
Attribute FileOffset.VB_Description = "Returns a zero-based offset, in characters, from the beginning of the path to the file name."
FileOffset = CIntToUInt(PropFileOffset)
End Property

Public Property Let FileOffset(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Filter() As String
Attribute Filter.VB_Description = "Returns/sets the filters that are displayed in the type list box of a dialog box."
Filter = PropFilter
End Property

Public Property Let Filter(ByVal Value As String)
PropFilter = Value
End Property

Public Property Get FilterIndex() As Long
Attribute FilterIndex.VB_Description = "Returns/sets a default filter."
FilterIndex = PropFilterIndex
End Property

Public Property Let FilterIndex(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropFilterIndex = Value
End Property

Public Property Get InitDir() As String
Attribute InitDir.VB_Description = "Returns/sets the initial file directory."
InitDir = PropInitDir
End Property

Public Property Let InitDir(ByVal Value As String)
PropInitDir = Value
End Property

Public Property Get DefaultExt() As String
Attribute DefaultExt.VB_Description = "Returns/sets the default file name extension for the dialog box."
DefaultExt = PropDefaultExt
End Property

Public Property Let DefaultExt(ByVal Value As String)
PropDefaultExt = Value
End Property

Public Property Get Color() As Long
Attribute Color.VB_Description = "Returns/sets the selected color."
Color = PropColor
End Property

Public Property Let Color(ByVal Value As Long)
PropColor = Value
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Set Font = PropFont
End Property

Public Property Let Font(ByVal NewFont As StdFont)
Set Me.Font = NewFont
End Property

Public Property Set Font(ByVal NewFont As StdFont)
Set PropFont = NewFont
End Property

Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets the smallest allowable font size (font dialog) or print range (print dialog)."
Min = PropMin
End Property

Public Property Let Min(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropMin = Value
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum font size (font dialog) or print range (print dialog)."
Max = PropMax
End Property

Public Property Let Max(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropMax = Value
End Property

Public Property Get FromPage() As Long
Attribute FromPage.VB_Description = "Returns/sets the value for the first page to be printed."
FromPage = PropFromPage
End Property

Public Property Let FromPage(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropFromPage = Value
End Property

Public Property Get ToPage() As Long
Attribute ToPage.VB_Description = "Returns/sets the value for the first page to be printed."
ToPage = PropToPage
End Property

Public Property Let ToPage(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropToPage = Value
End Property

Public Property Get Action() As Integer
Attribute Action.VB_Description = "Sets the type of dialog box to be displayed."
Attribute Action.VB_UserMemId = 0
Err.Raise Number:=394, Description:="Property is write-only"
End Property

Public Property Let Action(ByVal Value As Integer)
Select Case Value
    Case 1
        Me.ShowOpen
    Case 2
        Me.ShowSave
    Case 3
        Me.ShowColor
    Case 4
        Me.ShowFont
    Case 5
        Me.ShowPrinter
    Case 6
        Shell "winhlp32.exe"
    Case 7
        Me.ShowFolder
    Case Else
        Err.Raise 380
End Select
End Property

Public Function ShowOpen() As Boolean
Attribute ShowOpen.VB_Description = "Displays the open dialog box."
Dim Buffer As String, Filter As String
Buffer = String(PropMaxFileSize, vbNullChar)
Dim OFN As OPENFILENAME
With OFN
.lStructSize = LenB(OFN)
If Not Screen.ActiveForm Is Nothing Then
    .hWndOwner = Screen.ActiveForm.hWnd
Else
    .hWndOwner = GetActiveWindow()
End If
.hInstance = App.hInstance
Filter = ProperFilter(PropFilter)
.lpstrFilter = StrPtr(Filter)
.nFilterIndex = PropFilterIndex
.lpstrFile = StrPtr(Buffer)
.nMaxFile = Len(Buffer)
.lpstrInitialDir = StrPtr(PropInitDir)
.lpstrTitle = StrPtr(PropDialogTitle)
.Flags = PropFlags
End With
If GetOpenFileName(OFN) <> 0 Then
    PropFlags = OFN.Flags
    If Mid$(Buffer, OFN.nFileOffset, 1) = vbNullChar Then
        PropFileName = Left$(Buffer, InStr(Buffer, vbNullChar & vbNullChar) - 1)
        PropFileTitle = vbNullString
    Else
        PropFileName = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
        PropFileTitle = Mid$(PropFileName, OFN.nFileOffset + 1)
    End If
    PropFileOffset = OFN.nFileOffset
    ShowOpen = True
Else
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case FNERR_BUFFERTOOSMALL
            Err.Raise Number:=CdlBufferTooSmall, Description:="The buffer at which the member LpstrFile points is too small."
        Case FNERR_INVALIDFILENAME
            Err.Raise Number:=CdlInvalidFileName, Description:="File name is invalid."
        Case FNERR_SUBCLASSFAILURE
            Err.Raise Number:=CdlSubclassFailure, Description:="An attempt to subclass a listbox failed due to insufficient memory."
        Case 0
            If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Public Function ShowSave() As Boolean
Attribute ShowSave.VB_Description = "Displays the save dialog box."
Dim Buffer As String, Filter As String, Length As Long
Buffer = String(PropMaxFileSize, vbNullChar)
Dim OFN As OPENFILENAME
With OFN
.lStructSize = LenB(OFN)
If Not Screen.ActiveForm Is Nothing Then
    .hWndOwner = Screen.ActiveForm.hWnd
Else
    .hWndOwner = GetActiveWindow()
End If
.hInstance = App.hInstance
Filter = ProperFilter(PropFilter)
.lpstrFilter = StrPtr(Filter)
.nFilterIndex = PropFilterIndex
If Not PropFileName = vbNullString Then
    If Len(PropFileName) > (PropMaxFileSize + 1) Then
        Length = PropMaxFileSize * 2
    Else
        Length = LenB(PropFileName)
    End If
    CopyMemory ByVal StrPtr(Buffer), ByVal StrPtr(PropFileName), Length
End If
.lpstrFile = StrPtr(Buffer)
.nMaxFile = Len(Buffer)
.lpstrInitialDir = StrPtr(PropInitDir)
.lpstrTitle = StrPtr(PropDialogTitle)
.Flags = PropFlags
.lpstrDefExt = StrPtr(PropDefaultExt)
End With
If GetSaveFileName(OFN) <> 0 Then
    PropFlags = OFN.Flags
    If Mid$(Buffer, OFN.nFileOffset, 1) = vbNullChar Then
        PropFileName = Left$(Buffer, InStr(Buffer, vbNullChar & vbNullChar) - 1)
        PropFileTitle = vbNullString
    Else
        PropFileName = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
        PropFileTitle = Mid$(PropFileName, OFN.nFileOffset + 1)
    End If
    PropFileOffset = OFN.nFileOffset
    ShowSave = True
Else
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case FNERR_BUFFERTOOSMALL
            Err.Raise Number:=CdlBufferTooSmall, Description:="The buffer at which the member LpstrFile points is too small."
        Case FNERR_INVALIDFILENAME
            Err.Raise Number:=CdlInvalidFileName, Description:="File name is invalid."
        Case FNERR_SUBCLASSFAILURE
            Err.Raise Number:=CdlSubclassFailure, Description:="An attempt to subclass a listbox failed due to insufficient memory."
        Case 0
            If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

' Example for Filter: "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

Private Function ProperFilter(ByVal Filter As String) As String
Dim i As Long, Sign As String, Temp As String
For i = 1 To Len(Filter)
    Sign = Mid$(Filter, i, 1)
    If Sign = "|" Then
        Temp = Temp & vbNullChar
    Else
        Temp = Temp & Sign
    End If
Next i
Do Until Right$(Temp, 2) = vbNullChar & vbNullChar
    Temp = Temp & vbNullChar
Loop
ProperFilter = Temp
End Function

Public Function ShowColor() As Boolean
Attribute ShowColor.VB_Description = "Displays the color dialog box."
Static CustomColors(0 To 15) As Long, CustomColorsInitialized As Boolean
Dim CHCLR As TCHOOSECOLOR
With CHCLR
.lStructSize = LenB(CHCLR)
If Not Screen.ActiveForm Is Nothing Then
    .hWndOwner = Screen.ActiveForm.hWnd
Else
    .hWndOwner = GetActiveWindow()
End If
.hInstance = App.hInstance
.RGBResult = WinColor(PropColor)
.Flags = PropFlags
If CustomColorsInitialized = False Then
    Dim i As Long
    For i = 0 To 15
        CustomColors(i) = GetSysColor(i)
    Next i
    CustomColorsInitialized = True
End If
.lpCustColors = VarPtr(CustomColors(0))
End With
If ChooseColor(CHCLR) <> 0 Then
    PropColor = CHCLR.RGBResult
    ShowColor = True
Else
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case 0
            If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Public Function ShowFont() As Boolean
Attribute ShowFont.VB_Description = "Displays the font dialog box."
Const CF_INITTOLOGFONTSTRUCT As Long = &H40, CF_ENABLEHOOK As Long = &H8
Dim CHFONT As TCHOOSEFONT, LF As LOGFONT
Call OLEFontToLogFont(PropFont, LF)
With CHFONT
.lStructSize = LenB(CHFONT)
If Not Screen.ActiveForm Is Nothing Then
    .hWndOwner = Screen.ActiveForm.hWnd
Else
    .hWndOwner = GetActiveWindow()
End If
.lpLogFont = VarPtr(LF)
.Flags = CF_INITTOLOGFONTSTRUCT Or PropFlags
.RGBColor = PropColor
.nSizeMin = PropMin
.nSizeMax = PropMax
End With
If ChooseFont(CHFONT) <> 0 Then
    PropFlags = CHFONT.Flags And Not CF_INITTOLOGFONTSTRUCT
    With PropFont
    If (CHFONT.Flags And CF_NOFACESEL) = 0 Then .Name = Left$(LF.LFFaceName(), InStr(LF.LFFaceName(), vbNullChar) - 1)
    If (CHFONT.Flags And CF_NOSTYLESEL) = 0 Then
        .Bold = CBool(LF.LFWeight = FW_BOLD)
        .Italic = CBool(LF.LFItalic <> 0)
    End If
    If (CHFONT.Flags And CF_NOSIZESEL) = 0 Then .Size = CSng(CHFONT.iPointSize / 10)
    If (CHFONT.Flags And CF_EFFECTS) <> 0 Then
        .Strikethrough = CBool(LF.LFStrikeOut <> 0)
        .Underline = CBool(LF.LFUnderline <> 0)
        PropColor = CHFONT.RGBColor
    End If
    If (CHFONT.Flags And CF_NOSCRIPTSEL) = 0 Then .Charset = CInt(LF.LFCharset)
    End With
    ShowFont = True
Else
    Dim ErrVal As Long
    ErrVal = CommDlgExtendedError()
    Select Case ErrVal
        Case CFERR_MAXLESSTHANMIN
            Err.Raise Number:=CdlMaxLessThanMin, Description:="The size specified in the nSizeMax member is less than the size specified in the nSizeMin member."
        Case CFERR_NOFONTS
            Err.Raise Number:=CdlNoFonts, Description:="No fonts exist."
        Case 0
            If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Private Sub OLEFontToLogFont(ByVal Font As StdFont, ByRef LF As LOGFONT)
Dim FontName As String
With LF
FontName = Left$(Font.Name, LF_FACESIZE)
CopyMemory .LFFaceName(0), ByVal StrPtr(FontName), LenB(FontName)
.LFHeight = -MulDiv(CLng(Font.Size), DPI_Y(), 72)
If Font.Bold = True Then .LFWeight = FW_BOLD Else .LFWeight = FW_NORMAL
.LFItalic = IIf(Font.Italic = True, 1, 0)
.LFStrikeOut = IIf(Font.Strikethrough = True, 1, 0)
.LFUnderline = IIf(Font.Underline = True, 1, 0)
.LFQuality = DEFAULT_QUALITY
.LFCharset = CByte(Font.Charset And &HFF)
End With
End Sub

Public Function ShowPrinter() As CdlPDResultConstants
Attribute ShowPrinter.VB_Description = "Displays the printer dialog box."
Const DM_ORIENTATION As Long = &H1, DM_PAPERSIZE As Long = &H2, DM_PAPERWIDTH As Long = &H8, DM_COPIES As Long = &H100, DM_PRINTQUALITY As Long = &H400, DM_COLOR As Long = &H800, DM_DUPLEX As Long = &H1000, DM_COLLATE As Long = &H8000&
Dim PDLGEX As PRINTDLGEX, PPAGERANGE As PRINTPAGERANGE, DMODE As DEVMODE, DNAMES As DEVNAMES
Dim lpDevMode As Long, lpDevNames As Long
Dim ObjPrinter As VB.Printer, NewPrinterName As String
PDLGEX.lStructSize = Len(PDLGEX)
If Not Screen.ActiveForm Is Nothing Then
    PDLGEX.hWndOwner = Screen.ActiveForm.hWnd
Else
    PDLGEX.hWndOwner = GetActiveWindow()
End If
PDLGEX.Flags = PropFlags
PDLGEX.nPageRanges = 1
PDLGEX.nMaxPageRanges = 1
PPAGERANGE.nFromPage = PropFromPage
PPAGERANGE.nToPage = PropToPage
PDLGEX.nMinPage = PropMin
PDLGEX.nMaxPage = PropMax
PDLGEX.lpPageRanges = VarPtr(PPAGERANGE)
Const START_PAGE_GENERAL As Long = &HFFFFFFFF
PDLGEX.nStartPage = START_PAGE_GENERAL
If VB.Printers.Count > 0 Then
    With VB.Printer
    DMODE.DMSize = Len(DMODE)
    DMODE.DMDeviceName = Left$(.DeviceName, CCHDEVICENAME)
    DMODE.DMFields = DM_ORIENTATION Or DM_PAPERSIZE Or DM_PAPERWIDTH Or DM_COPIES Or DM_PRINTQUALITY Or DM_COLOR Or DM_DUPLEX Or DM_COLLATE
    DMODE.DMOrientation = .Orientation
    DMODE.DMPaperSize = .PaperSize
    DMODE.DMPaperWidth = .Width
    DMODE.DMCopies = .Copies
    DMODE.DMDefaultSource = .PaperBin
    DMODE.DMPrintQuality = .PrintQuality
    DMODE.DMColor = .ColorMode
    DMODE.DMDuplex = .Duplex
    DMODE.DMCollate = IIf((PDLGEX.Flags And CdlPDCollate) <> 0, 1, 0)
    DNAMES.wDriverOffset = 8
    DNAMES.wDeviceOffset = DNAMES.wDriverOffset + Len(.DriverName) + 1
    DNAMES.wOutputOffset = DNAMES.wDeviceOffset + Len(.DeviceName) + 1
    DNAMES.wDefault = 0
    DNAMES.wExtra = Left$(.DriverName & vbNullChar & .DeviceName & vbNullChar & .Port & vbNullChar, CCHDEVNAMESEXTRA)
    PDLGEX.hDC = .hDC
    PDLGEX.nCopies = .Copies
    PDLGEX.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DMODE))
    lpDevMode = GlobalLock(PDLGEX.hDevMode)
    CopyMemory ByVal lpDevMode, DMODE, Len(DMODE)
    GlobalUnlock PDLGEX.hDevMode
    PDLGEX.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DNAMES))
    lpDevNames = GlobalLock(PDLGEX.hDevNames)
    CopyMemory ByVal lpDevNames, DNAMES, Len(DNAMES)
    GlobalUnlock PDLGEX.hDevNames
    End With
Else
    PDLGEX.hDevMode = 0
    PDLGEX.hDevNames = 0
End If
Dim ErrVal As Long
ErrVal = PrintDialogEx(PDLGEX)
Const S_OK As Long = &H0
If ErrVal = S_OK Then
    If PDLGEX.dwResultAction > 0 Then
        lpDevMode = GlobalLock(PDLGEX.hDevMode)
        CopyMemory DMODE, ByVal lpDevMode, Len(DMODE)
        GlobalUnlock PDLGEX.hDevMode
        GlobalFree PDLGEX.hDevMode
        lpDevNames = GlobalLock(PDLGEX.hDevNames)
        CopyMemory DNAMES, ByVal lpDevNames, Len(DNAMES)
        GlobalUnlock PDLGEX.hDevNames
        GlobalFree PDLGEX.hDevNames
        NewPrinterName = Left$(DMODE.DMDeviceName, InStr(DMODE.DMDeviceName, vbNullChar) - 1)
        If StrComp(VB.Printer.DeviceName, NewPrinterName, vbTextCompare) <> 0 Then
            For Each ObjPrinter In VB.Printers
                If StrComp(ObjPrinter.DeviceName, NewPrinterName, vbTextCompare) = 0 Then
                    Set VB.Printer = ObjPrinter
                End If
            Next ObjPrinter
        End If
        On Error Resume Next
        With VB.Printer
        .Copies = DMODE.DMCopies
        .Duplex = DMODE.DMDuplex
        .Orientation = DMODE.DMOrientation
        .PaperSize = DMODE.DMPaperSize
        .PrintQuality = DMODE.DMPrintQuality
        .ColorMode = DMODE.DMColor
        .PaperBin = DMODE.DMDefaultSource
        End With
        On Error GoTo 0
        PropFlags = PDLGEX.Flags
        If (PropFlags And CdlPDUseDevModeCopiesAndCollate) <> 0 Then
            If DMODE.DMCollate = 1 And (PropFlags And CdlPDCollate) = 0 Then PropFlags = PropFlags Or CdlPDCollate
        End If
        PropFromPage = PPAGERANGE.nFromPage
        PropToPage = PPAGERANGE.nToPage
        PropMin = PDLGEX.nMinPage
        PropMax = PDLGEX.nMaxPage
        If (PDLGEX.Flags And (CdlPDReturnDC Or CdlPDReturnIC)) <> 0 Then PropDC = PDLGEX.hDC
        ShowPrinter = PDLGEX.dwResultAction
    Else
        If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
    End If
Else
    If PDLGEX.hDevMode <> 0 Then GlobalFree PDLGEX.hDevMode
    If PDLGEX.hDevNames <> 0 Then GlobalFree PDLGEX.hDevNames
    Const E_OUTOFMEMORY As Long = &H8007000E, E_INVALIDARG As Long = &H80070057, E_POINTER As Long = &H80004003, E_HANDLE As Long = &H80070006, E_FAIL As Long = &H80004005
    Select Case ErrVal
        Case E_OUTOFMEMORY, E_INVALIDARG, E_POINTER, E_HANDLE, E_FAIL
            Err.Raise Number:=CdlInitFailure, Description:="The PrintDlgEx function failed during initialization."
        Case Else
            Err.Raise Number:=ErrVal, Description:="Unexpected error."
    End Select
End If
End Function

Public Function ShowFolder(Optional ByVal RootFolder As Variant) As Boolean
Attribute ShowFolder.VB_Description = "Displays the folder dialog box."
Dim BIF As BROWSEINFO, IDList As Long
With BIF
If Not Screen.ActiveForm Is Nothing Then
    .hWndOwner = Screen.ActiveForm.hWnd
Else
    .hWndOwner = GetActiveWindow()
End If
If IsMissing(RootFolder) Then
    .pIDLRoot = 0
Else
    Select Case VarType(RootFolder)
        Case vbLong, vbInteger, vbByte
            SHGetFolderLocation 0, RootFolder, 0, 0, .pIDLRoot
        Case vbString
            If ComCtlsW2KCompatibility() = False Then
                .pIDLRoot = ILCreateFromPath(StrPtr(Left$(RootFolder, MAX_PATH)))
            Else
                .pIDLRoot = ILCreateFromPath_W2K(StrPtr(Left$(RootFolder, MAX_PATH)))
            End If
        Case Else
            Err.Raise 380
    End Select
End If
.lpszTitle = StrPtr(PropDialogTitle)
.ulFlags = PropFlags
If Not PropInitDir = vbNullString Then
    .lpfnCallback = ProcPtr(AddressOf CdlShowFolderCallback)
    .lParam = StrPtr(PropInitDir)
End If
IDList = SHBrowseForFolder(BIF)
If .pIDLRoot <> 0 Then CoTaskMemFree .pIDLRoot
End With
If IDList <> 0 Then
    Dim Buffer As String, PathName As String
    Buffer = String(MAX_PATH, vbNullChar)
    SHGetPathFromIDList IDList, StrPtr(Buffer)
    If Left$(Buffer, 2) = vbNullChar & ":" Then Buffer = Mid$(Buffer, 3)
    PathName = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    CoTaskMemFree IDList
    On Error Resume Next
    Dim Attributes As VbFileAttribute
    Attributes = GetAttr(PathName)
    On Error GoTo 0
    If (Attributes And (vbDirectory Or vbVolume)) <> 0 Then
        If Not PathName = vbNullString Then PathName = PathName & IIf(Right$(PathName, 1) = "\", "", "\")
        PropFileTitle = vbNullString
    Else
        Dim i As Long, Pos As Long
        For i = 1 To Len(PathName)
            If Mid$(PathName, i, 1) = "\" Then Pos = i
        Next i
        PropFileTitle = Right$(PathName, Len(PathName) - Pos)
    End If
    PropFileName = PathName
    ShowFolder = True
Else
    If PropCancelError = True Then Err.Raise Number:=CdlCancel, Description:="Cancel was selected."
End If
End Function
