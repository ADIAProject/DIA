Attribute VB_Name = "mAnime"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Source Code   : AnimateForm
' Auther        : Jim Jose
' eMail         : jimjosev33@yahoo.com
' Purpose       : Cooool flash style animations in Vb
' Comment       : Function contains 13 effects, each have
'               : reverse effect too. So total 26 animations in one function
'               : Completly error checked and free from memory leaks
' Copyright Jim Jose, Gtech Creations - 2005
'----------------------------------------------------------------------------------------------------------------------------------------------------------
'[Event Enum]
Public Enum AnimeEventEnum
    aUnload = 0
    aLoad = 1
End Enum

'[Effect Enum]
Public Enum AnimeEffectEnum
    eAppearFromLeft = 0
    eAppearFromRight = 1
    eAppearFromTop = 2
    eAppearFromBottom = 3
    eGenerateLeftTop = 4
    eGenerateLeftBottom = 5
    eGenerateRightTop = 6
    eGenerateRightBottom = 7
    eStrechHorizontally = 8
    eStrechVertically = 9
    eZoomOut = 10
    eFoldOut = 11
    eCurtonHorizontal = 12
    eCurtonVertical = 13
End Enum

'[Constants]
Private Const RGN_AND  As Long = 1
Private Const RGN_OR   As Long = 2
Private Const RGN_XOR  As Long = 3
Private Const RGN_COPY As Long = 5
Private Const RGN_DIFF As Long = 4
Private Const HWND_NOTOPMOST = -2
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const SWP_FLAGS = SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function AnimateForm
'! Description (Описание)  :   [Cooool flash style animations in Vb]
'! Parameters  (Переменные):   hwndObject (Object)
'                              aEvent (AnimeEventEnum)
'                              aEffect (AnimeEffectEnum = 11)
'                              FrameTime (Long = 1)
'                              FrameCount (Long = 33)
'!--------------------------------------------------------------------------------
Public Function AnimateForm(hwndObject As Object, ByVal aEvent As AnimeEventEnum, Optional ByVal aEffect As AnimeEffectEnum = 11, Optional ByVal FrameTime As Long = 1, Optional ByVal FrameCount As Long = 33) As Boolean

    On Error GoTo Handle

    Dim X1      As Long, Y1                  As Long
    Dim hRgn    As Long, tmpRgn            As Long
    Dim XValue  As Long, YValue          As Long
    Dim XIncr   As Double, YIncr          As Double
    Dim wHeight As Long, wWidth         As Long

    wWidth = hwndObject.Width / Screen.TwipsPerPixelX
    wHeight = hwndObject.Height / Screen.TwipsPerPixelY
    hwndObject.Visible = True

    Select Case aEffect

        Case eAppearFromLeft
            XIncr = wWidth / FrameCount

            For X1 = 0 To FrameCount
                ' Define the size of current frame/Create it
                XValue = X1 * XIncr
                hRgn = CreateRectRgn(0, 0, XValue, wHeight)

                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If

                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eAppearFromRight
            XIncr = wWidth / FrameCount

            For X1 = 0 To FrameCount
                ' Define the size of current frame/Create it
                XValue = wWidth - X1 * XIncr
                hRgn = CreateRectRgn(XValue, 0, wWidth, wHeight)

                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If

                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eAppearFromTop
            YIncr = wHeight / FrameCount

            For Y1 = 0 To FrameCount
                ' Define the size of current frame/Create it
                YValue = Y1 * YIncr
                hRgn = CreateRectRgn(0, 0, wWidth, YValue)

                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If

                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eAppearFromBottom
            YIncr = wHeight / FrameCount

            For Y1 = 0 To FrameCount
                ' Define the size of current frame/Create it
                YValue = wHeight - Y1 * YIncr
                hRgn = CreateRectRgn(0, YValue, wWidth, wHeight)

                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If

                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eGenerateLeftTop
            XIncr = wWidth / FrameCount
            YIncr = wHeight / FrameCount

            For X1 = 0 To FrameCount

                ' Define / Create Region
                If aEvent = aLoad Then
                    XValue = X1 * XIncr
                    YValue = X1 * YIncr
                Else
                    XValue = wWidth - X1 * XIncr
                    YValue = wHeight - X1 * YIncr
                End If

                hRgn = CreateRectRgn(0, 0, XValue, YValue)
                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eGenerateLeftBottom
            XIncr = wWidth / FrameCount
            YIncr = wHeight / FrameCount

            For X1 = 0 To FrameCount

                ' Define / Create Region
                If aEvent = aLoad Then
                    XValue = X1 * XIncr
                    YValue = wHeight - X1 * YIncr
                Else
                    XValue = wWidth - X1 * XIncr
                    YValue = X1 * YIncr
                End If

                hRgn = CreateRectRgn(0, wHeight, XValue, YValue)
                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eGenerateRightTop
            XIncr = wWidth / FrameCount
            YIncr = wHeight / FrameCount

            For X1 = 0 To FrameCount

                ' Define / Create Region
                If aEvent = aLoad Then
                    XValue = wWidth - X1 * XIncr
                    YValue = X1 * YIncr
                Else
                    XValue = X1 * XIncr
                    YValue = wHeight - X1 * YIncr
                End If

                hRgn = CreateRectRgn(XValue, YValue, wWidth, 0)
                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eGenerateRightBottom
            XIncr = wWidth / FrameCount
            YIncr = wHeight / FrameCount

            For X1 = 0 To FrameCount

                ' Define / Create Region
                If aEvent = aLoad Then
                    XValue = wWidth - X1 * XIncr
                    YValue = wHeight - X1 * YIncr
                Else
                    XValue = X1 * XIncr
                    YValue = X1 * YIncr
                End If

                hRgn = CreateRectRgn(XValue, YValue, wWidth, wHeight)
                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eStrechHorizontally
            XIncr = wWidth / FrameCount

            For X1 = 0 To FrameCount

                ' Define / Create Region
                If aEvent = aLoad Then
                    XValue = wWidth - X1 * XIncr
                Else
                    XValue = X1 * XIncr
                End If

                hRgn = CreateRectRgn(XValue / 2, 0, wWidth - XValue / 2, wHeight)
                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eStrechVertically
            YIncr = wHeight / FrameCount

            For Y1 = 0 To FrameCount

                ' Define / Create Region
                If aEvent = aLoad Then
                    YValue = Y1 * YIncr
                Else
                    YValue = wHeight - Y1 * YIncr
                End If

                hRgn = CreateRectRgn(0, wHeight / 2 - YValue / 2, wWidth, wHeight / 2 + YValue / 2)
                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eZoomOut
            XIncr = wWidth / FrameCount
            YIncr = wHeight / FrameCount

            For X1 = 0 To FrameCount

                ' Define / Create Region
                If aEvent = aLoad Then
                    XValue = X1 * XIncr
                    YValue = X1 * YIncr
                Else
                    XValue = wWidth - X1 * XIncr
                    YValue = wHeight - X1 * YIncr
                End If

                hRgn = CreateRectRgn((wWidth - XValue) / 2, (wHeight - YValue) / 2, (wWidth + XValue) / 2, (wHeight + YValue) / 2)
                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eFoldOut

            If hwndObject.Width >= hwndObject.Height Then
                XIncr = wWidth / FrameCount
                YIncr = wWidth / FrameCount
            Else
                XIncr = wHeight / FrameCount
                YIncr = wHeight / FrameCount
            End If

            For X1 = 0 To FrameCount

                ' Define / Create Region
                If aEvent = aLoad Then
                    XValue = X1 * XIncr
                    YValue = X1 * YIncr
                Else
                    XValue = wWidth - X1 * XIncr
                    YValue = wHeight - X1 * YIncr
                End If

                hRgn = CreateRectRgn((wWidth - XValue) / 2, (wHeight - YValue) / 2, (wWidth + XValue) / 2, (wHeight + YValue) / 2)
                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eCurtonHorizontal

            Dim ScanWidth As Long

            ScanWidth = FrameCount / 2

            For Y1 = 0 To FrameCount / 2
                ' Initiate region
                hRgn = CreateRectRgn(0, 0, 0, 0)

                For X1 = 0 To wHeight / FrameCount * 2
                    ' Create each curton region
                    tmpRgn = CreateRectRgn(0, X1 * ScanWidth, wWidth, X1 * ScanWidth + Y1)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_OR
                    DeleteObject tmpRgn
                Next

                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If

                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

        Case eCurtonVertical
            ScanWidth = FrameCount / 2

            For X1 = 0 To FrameCount / 2
                ' Initiate Region
                hRgn = CreateRectRgn(0, 0, 0, 0)

                For Y1 = 0 To wWidth / FrameCount * 2
                    ' Create each curton region
                    tmpRgn = CreateRectRgn(Y1 * ScanWidth, 0, Y1 * ScanWidth + X1, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_OR
                    DeleteObject tmpRgn
                Next

                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If

                ' Set the new region for the window
                SetWindowRgn hwndObject.hWnd, hRgn, True
                DoEvents
                Sleep FrameTime
            Next

    End Select

    AnimateForm = True

    Exit Function

Handle:
    AnimateForm = False
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub MakeTopMostNoFocus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   hWnd (Long)
'!--------------------------------------------------------------------------------
Public Sub MakeTopMostNoFocus(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS
End Sub
