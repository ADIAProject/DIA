Attribute VB_Name = "mMenu"
Option Explicit

'--- ��� ������������ ����
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
'!  �������     :  OpenContextMenu
'!  ����������  :  FormName As Form, MenuName As Menu
'!  ��������    :  ����� ������������ ����
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub OpenContextMenu
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   FormName (Form)
'                              MenuName (Menu)
'!--------------------------------------------------------------------------------
Public Sub OpenContextMenu(FormName As Form, MenuName As Menu)
    ' ������� �������, ��� ������������ ������� ������ ������� ���� �� �����
    'SendMessage FormName.hWnd, WM_RBUTTONDOWN, 0, 0&
    ' ���������� ����������� ����
    FormName.PopupMenu MenuName
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub SetMenuIcon
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   hWnd (Long)
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
