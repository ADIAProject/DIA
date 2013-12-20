Attribute VB_Name = "mMenu"
Option Explicit

'--- для контекстного меню
Private Declare Function GetMenu Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean

Private Type MENUITEMINFO
    cbSize                                  As Long
    fMask                               As Long
    fType                               As Long
    fState                              As Long
    wID                                 As Long
    hSubMenu                            As Long
    hbmpChecked                         As Long
    hbmpUnchecked                       As Long
    dwItemData                          As Long
    dwTypeData                          As String
    cch                                 As Long
End Type

'! -----------------------------------------------------------
'!  Функция     :  OpenContextMenu
'!  Переменные  :  FormName As Form, MenuName As Menu
'!  Описание    :  вызов контекстного меню
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub OpenContextMenu
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   FormName (Form)
'                              MenuName (Menu)
'!--------------------------------------------------------------------------------
Public Sub OpenContextMenu(FormName As Form, MenuName As Menu)
    ' Говорит системе, что пользователь щелкнул правой кнопкой мыши на форме
    'SendMessage FormName.hWnd, WM_RBUTTONDOWN, 0, 0&
    ' Показывает контекстное меню
    FormName.PopupMenu MenuName
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetMenuIcon
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   hWnd (Long)
'                              MenuIndex (Long)
'                              SubIndex (Long)
'                              Pic (Picture)
'!--------------------------------------------------------------------------------
Public Sub SetMenuIcon(hWnd As Long, MenuIndex As Long, SubIndex As Long, Pic As Picture)

    Dim hMenu As Long, hSubMenu As Long, hID As Long

    'Get the menuhandle of the form
    hMenu = GetMenu(hWnd)
    'Get the handle of the first submenu
    hSubMenu = GetSubMenu(hMenu, MenuIndex)
    'Get the menuId of the first entry
    hID = GetMenuItemID(hSubMenu, SubIndex)
    'Add the bitmap
    'SetMenuItemBitmaps hMenu, hID, MF_BITMAP, Pic, Pic
End Sub
