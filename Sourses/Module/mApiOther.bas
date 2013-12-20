Attribute VB_Name = "mApiOther"
Option Explicit

Public Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Public Type TRACKMOUSEEVENT_STRUCT
    cbSize                              As Long
    dwFlags                             As TRACKMOUSEEVENT_FLAGS
    hWndTrack                           As Long
    dwHoverTime                         As Long
End Type

Public Declare Function TrackMouseEvent Lib "user32.dll" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Public Declare Function TrackMouseEventComCtl Lib "Comctl32.dll" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function IsUserAnAdmin Lib "shell32.dll" () As Long
Public Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Const SND_ASYNC    As Long = &H1    'play asynchronously
Public Const SND_FILENAME As Long = &H20000    'sound is file name
Public Const EM_GETSEL    As Long = &HB0

' Character sets
Public Const ANSI_CHARSET = 0
Public Const DEFAULT_CHARSET = 1
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const HANGEUL_CHARSET = 129
Public Const HANGUL_CHARSET = 129
Public Const GB2312_CHARSET = 134
Public Const CHINESEBIG5_CHARSET = 136
Public Const OEM_CHARSET = 255
Public Const JOHAB_CHARSET = 130
Public Const HEBREW_CHARSET = 177
Public Const ARABIC_CHARSET = 178
Public Const GREEK_CHARSET = 161
Public Const TURKISH_CHARSET = 162
Public Const VIETNAMESE_CHARSET = 163
Public Const THAI_CHARSET = 222
Public Const EASTEUROPE_CHARSET = 238
Public Const RUSSIAN_CHARSET = 204
Public Const MAC_CHARSET = 77
Public Const BALTIC_CHARSET = 186
