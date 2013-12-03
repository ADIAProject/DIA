Attribute VB_Name = "mTrayIcon"
Option Explicit

'Как разместить иконку программы в TrayBar
Public Const NIM_ADD                    As Long = 0
Public Const NIM_DELETE                 As Long = 2

Private Const NIF_ICON                  As Long = 2
Private Const NIF_TIP                   As Long = 4

Type NOTIFYICONDATA
    cbSize                                  As Long
    hWnd                                As Long
    uId                                 As Long
    uFlags                              As Long
    uCallbackMessage                    As Long
    hIcon                               As Long
    szTip                               As String * 64

End Type

Private Declare Function Shell_NotifyIconA _
                          Lib "shell32.dll" (ByVal dwMessage As Long, _
                                             lpData As NOTIFYICONDATA) As Integer

'! -----------------------------------------------------------
'!  Функция     :  SetTrayIcon
'!  Переменные  :  Mode As Long, hwnd As Long, Icon As Long, tip As String
'!  Возвр. знач.:  As Long
'!  Описание    :  Установка значка в трей
'! -----------------------------------------------------------
Public Function SetTrayIcon(Mode As Long, _
                            ByVal lngHWnd As Long, _
                            ByVal lngIcon As Long, _
                            ByVal tip As String) As Long

Dim nidTemp                             As NOTIFYICONDATA

    With nidTemp
        .cbSize = Len(nidTemp)
        .hWnd = lngHWnd
        .uId = 0&
        .uFlags = NIF_ICON Or NIF_TIP
        .uCallbackMessage = 0&
        .hIcon = lngIcon
        .szTip = tip & vbNullChar

    End With

    'NIDTEMP
    SetTrayIcon = Shell_NotifyIconA(Mode, nidTemp)

End Function
