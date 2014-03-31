Attribute VB_Name = "mSetAlphaIcon"
Option Explicit

Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&
Private Const IMAGE_ICON = 1
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpsz As Long, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetIcon
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   hWnd (Long)
'                              sIconResName (String)
'                              bSetAsAppIcon (Boolean = True)
'!--------------------------------------------------------------------------------
Public Sub SetIcon(ByVal hWnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)

    Dim lhWndTop   As Long
    Dim lhWnd      As Long
    Dim hIconLarge As Long
    Dim hIconSmall As Long

    If (bSetAsAppIcon) Then
        ' Find VB's hidden parent window:
        lhWnd = hWnd
        lhWndTop = lhWnd

        Do While Not (lhWnd = 0)
            lhWnd = GetWindow(lhWnd, GW_OWNER)

            If Not (lhWnd = 0) Then
                lhWndTop = lhWnd
            End If

        Loop

    End If

    hIconLarge = LoadImage(App.hInstance, StrPtr(sIconResName & vbNullChar), IMAGE_ICON, 32, 32, LR_DEFAULTCOLOR)

    If (bSetAsAppIcon) Then
        SendMessage lhWndTop, WM_SETICON, ICON_BIG, ByVal hIconLarge
    End If
    SendMessage hWnd, WM_SETICON, ICON_BIG, ByVal hIconLarge
    
    hIconSmall = LoadImage(App.hInstance, StrPtr(sIconResName & vbNullChar), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR)

    If (bSetAsAppIcon) Then
        SendMessage lhWndTop, WM_SETICON, ICON_SMALL, ByVal hIconSmall
    End If

    SendMessage hWnd, WM_SETICON, ICON_SMALL, ByVal hIconSmall
End Sub
