Attribute VB_Name = "mTaskBar"
'Определение рабочей области Windows, то есть часть экрана, не затененного панелью задач или другими прикладными областями
Option Explicit

Public lngTopWorkArea        As Long
Public lngLeftWorkArea       As Long
Public lngRightWorkArea      As Long
Public lngBottomWorkArea     As Long

Public Const SPI_GETWORKAREA As Integer = 48

Public Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub GetWorkArea
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub GetWorkArea()

    Dim typRect As RECT

    SystemParametersInfo SPI_GETWORKAREA, vbNull, typRect, 0

    With typRect
        lngTopWorkArea = .Top * Screen.TwipsPerPixelY
        lngLeftWorkArea = .Left * Screen.TwipsPerPixelX
        lngRightWorkArea = .Right * Screen.TwipsPerPixelX
        lngBottomWorkArea = .Bottom * Screen.TwipsPerPixelY
    End With

End Sub

Public Function HPadding(ByVal Form As Form)

    Dim SaveMode As Integer

    With Form
        SaveMode = .ScaleMode
        .ScaleMode = vbTwips
        HPadding = .Width - .ScaleWidth
        .ScaleMode = SaveMode
        DoEvents
    End With
End Function

Public Function VPadding(ByVal Form As Form)

    Dim SaveMode As Integer

    With Form
        SaveMode = .ScaleMode
        .ScaleMode = vbTwips
        VPadding = .Height - .ScaleHeight
        .ScaleMode = SaveMode
        DoEvents
    End With
End Function

