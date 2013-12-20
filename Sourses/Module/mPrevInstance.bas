Attribute VB_Name = "mPrevInstance"
'Активация ранее запущенной этой же программы
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function OpenIcon Lib "user32.dll" (ByVal hWnd As Long) As Long

'! -----------------------------------------------------------
'!  Функция     :  ShowPrevInstance
'!  Переменные  :
'!  Описание    :  Отобразить предыдущую копию программы, если программа запущена дважды
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ShowPrevInstance
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub ShowPrevInstance()

    Dim OldTitle     As String
    Dim hWnd         As Long

    Const SW_RESTORE As Long = 9

    OldTitle = App.Title
    App.Title = "This App Will Be Closed"
    hWnd = FindWindow(StrPtr("ThunderRT6FormDC"), StrPtr(OldTitle))

    If hWnd <> 0 Then
        ShowWindow hWnd, SW_RESTORE
        OpenIcon hWnd
        SetForegroundWindow hWnd
        AppActivate OldTitle

        End

    End If

End Sub
