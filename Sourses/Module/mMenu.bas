Attribute VB_Name = "mMenu"
Option Explicit

Private Declare Sub SetMenuDefaultItem Lib "user32.dll" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Boolean)
Private Declare Function GetMenu Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpmii As MENUITEMINFO) As Boolean
Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpmii As MENUITEMINFO) As Boolean

Private Const MIIM_STATE              As Long = &H1
Private Const MIIM_ID                 As Long = &H2
Private Const MIIM_SUBMENU            As Long = &H4
Private Const MIIM_CHECKMARKS         As Long = &H8
Private Const MIIM_TYPE               As Long = &H10
Private Const MIIM_DATA               As Long = &H20
Private Const MIIM_STRING             As Long = &H40
Private Const MFT_RADIOCHECK          As Long = &H200

Private Const MF_BYCOMMAND            As Long = &H0&
Private Const MF_DISABLED             As Long = &H2&
Private Const MF_STRING               As Long = &H0&
Private Const MF_BITMAP               As Long = &H4&
Private Const MF_CHECKED              As Long = &H8&
Private Const MF_MENUBARBREAK         As Long = &H20&
Private Const MF_MENUBREAK            As Long = &H40&
Private Const MF_OWNERDRAW            As Long = &H100&
Private Const MF_RADIOCHECK           As Long = &H200&
Private Const MF_BYPOSITION           As Long = &H400&
Private Const MF_SEPARATOR            As Long = &H800&
Private Const MF_RIGHTORDER           As Long = &H2000&
Private Const MF_RIGHTJUSTIFY         As Long = &H4000&

'http://msdn.microsoft.com/en-us/library/windows/desktop/ms647578%28v=vs.85%29.aspx
'typedef struct tagMENUITEMINFO {
'  UINT      cbSize;
'  UINT      fMask;
'  UINT      fType;
'  UINT      fState;
'  UINT      wID;
'  HMENU     hSubMenu;
'  HBITMAP   hbmpChecked;
'  HBITMAP   hbmpUnchecked;
'  ULONG_PTR dwItemData;
'  LPTSTR    dwTypeData;
'  UINT      cch;
'  HBITMAP   hbmpItem;
'} MENUITEMINFO, *LPMENUITEMINFO;
Private Type MENUITEMINFO
    cbSize          As Long
    fMask           As Long
    fType           As Long
    fState          As Long
    wid             As Long
    hSubMenu        As Long
    hbmpChecked     As Long
    hbmpUnchecked   As Long
    dwItemData      As Long
    dwTypeData      As Long
    cch             As Long
    hbmpItem        As Long
End Type

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub OpenContextMenu
'! Description (Описание)  :   [вызов контекстного меню]
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
    SetMenuItemBitmaps hMenu, hID, MF_BITMAP, Pic, Pic
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetUniMenu
'! Description (Описание)  :   [Set Unicode Caption for menu]
'! Parameters  (Переменные):   sCaption (String)
'                              mnu (Menu)
'                              mnuItem (Long)
'                              mnuParentItem (Long)
'                              IsDefault (Boolean)
'!--------------------------------------------------------------------------------
Public Sub SetUniMenu(ByVal mnuParentItem As Long, ByVal mnuItem As Long, ByVal mnuSubItem As Long, ByVal mnu As Menu, ByVal sCaption As String, Optional IsDefault As Boolean = False, Optional strShortcut As String = vbNullString)

    Dim hMenu As Long
    Dim mInfo  As MENUITEMINFO
    
    If mnuParentItem = -1 Then
        hMenu = GetMenu(mnu.Parent.hWnd)
    Else
        hMenu = GetSubMenu(GetMenu(mnu.Parent.hWnd), mnuParentItem)
        'Shortcut to Menu
        If LenB(strShortcut) Then
            sCaption = sCaption & vbTab & strShortcut
        End If
    End If
    
    If hMenu <> 0 Then
        With mInfo
            If mnuSubItem <> -1 Then
                'DropDown Submenu Type with IdNumber
                .fMask = MIIM_SUBMENU Or MIIM_ID
                .dwTypeData = StrPtr(FillNullChar(255))
                .cch = 255
                .cbSize = Len(mInfo)
                ' MenuItem Number
                .wid = mnuSubItem
                'Get DropDown Submenu Info handle
                GetMenuItemInfo hMenu, mnuItem, True, mInfo
                'Get DropDown Submenu handle
                hMenu = .hSubMenu
            End If
            
            If hMenu <> 0 Then
                .cbSize = Len(mInfo)
                .fMask = MIIM_STRING
                'mnu
                .dwTypeData = StrPtr(sCaption)
                                
                If mnuSubItem = -1 Then
                ' Not DropDown Submenu
                    SetMenuItemInfo hMenu, mnuItem, True, mInfo
                    If IsDefault Then SetMenuDefaultItem hMenu, mnuItem, True
                Else
                    SetMenuItemInfo hMenu, mnuSubItem, True, mInfo
                    If IsDefault Then SetMenuDefaultItem hMenu, mnuSubItem, True
                End If
            Else
                mnu.Caption = sCaption
            End If
        End With
    Else
        mnu.Caption = sCaption
    End If
    
End Sub
