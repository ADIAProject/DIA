Attribute VB_Name = "mMenu"
Option Explicit

'--- для контекстного меню
Private Declare Function GetMenu Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu _
                          Lib "user32.dll" (ByVal hMenu As Long, _
                                            ByVal nPos As Long) As Long

Private Declare Function GetMenuItemID _
                          Lib "user32.dll" (ByVal hMenu As Long, _
                                            ByVal nPos As Long) As Long

Private Declare Function SetMenuItemBitmaps _
                          Lib "user32.dll" (ByVal hMenu As Long, _
                                            ByVal nPosition As Long, _
                                            ByVal wFlags As Long, _
                                            ByVal hBitmapUnchecked As Long, _
                                            ByVal hBitmapChecked As Long) As Long

Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo _
                          Lib "user32.dll" _
                              Alias "GetMenuItemInfoA" (ByVal hMenu As Long, _
                                                        ByVal un As Long, _
                                                        ByVal B As Boolean, _
                                                        lpMenuItemInfo As MENUITEMINFO) As Boolean

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
Public Sub OpenContextMenu(FormName As Form, MenuName As Menu)
' Говорит системе, что пользователь щелкнул правой кнопкой мыши на форме
'SendMessage FormName.hWnd, WM_RBUTTONDOWN, 0, 0&
' Показывает контекстное меню
    FormName.PopupMenu MenuName

End Sub

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
