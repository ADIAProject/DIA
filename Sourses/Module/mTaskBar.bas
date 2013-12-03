Attribute VB_Name = "mTaskBar"
'����������� ������� ������� Windows, �� ���� ����� ������, �� ����������� ������� ����� ��� ������� ����������� ���������
Option Explicit

Public lngTopWorkArea                   As Long
Public lngLeftWorkArea                  As Long
Public lngRightWorkArea                 As Long
Public lngBottomWorkArea                As Long

Public Const SPI_GETWORKAREA            As Integer = 48

Public Declare Function SystemParametersInfo _
                         Lib "user32.dll" _
                             Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                            ByVal uParam As Long, _
                                                            lpvParam As Any, _
                                                            ByVal fuWinIni As Long) As Long

Public Sub GetWorkArea()

Dim typRect                             As RECT

    SystemParametersInfo SPI_GETWORKAREA, vbNull, typRect, 0

    With typRect
        lngTopWorkArea = .Top * Screen.TwipsPerPixelY
        lngLeftWorkArea = .Left * Screen.TwipsPerPixelX
        lngRightWorkArea = .Right * Screen.TwipsPerPixelX
        lngBottomWorkArea = .Bottom * Screen.TwipsPerPixelY

    End With

End Sub
