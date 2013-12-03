Attribute VB_Name = "mRTF_RTL"
'==========================================================
'           Copyright Information
'==========================================================
'Program Name: Azkary
'Program Author   : Elsheshtawy, Ahmed Amin
'Home Page        : http://www.islamware.com
'Copyrights © 2007 Islamware. All rights reserved.
'==========================================================
'Permission to use, copy, modify, and distribute this software and its
'documentation for any purpose and without fee is hereby granted.
'==========================================================
Option Explicit

Public Const EM_SETPARAFORMAT = WM_USER + 71
Public Const EM_SETBIDIOPTIONS = WM_USER + 200
Public Const PFM_DIR = &H10000    'Direction mask bit
Public Const PFE_RTLPAR = &H1    'RTL paragraph style bit
Public Const BOM_DEFPARADIR = &H1    'Default direction mask
Public Const BOE_RTLDIR = &H1    'Default RTL para style

Type vbParaFormat
    cbSize                                  As Long
    dwMask                              As Long
    wNumbering                          As Integer
    wEffects                            As Integer
    dxStrtIndent                        As Long
    DXRightIndent                       As Long
    DXOffset                            As Long
    wAlignment                          As Integer
    cTabCount                           As Integer
    rgxTabs(31)                         As Long

End Type

Public vbPF                             As vbParaFormat
Public Declare Function SendPFMessage _
                         Lib "user32.dll" _
                             Alias "SendMessageA" (ByVal hWnd As Long, _
                                                   ByVal wMsg As Long, _
                                                   ByVal wParam As Long, _
                                                   lParam As vbParaFormat) As Long

Type vbBiDiOptions
    cbSize                                  As Long
    wMask                               As Integer
    wEffects                            As Integer

End Type

Public vbBO                             As vbBiDiOptions
Public Declare Function SendBOMessage _
                         Lib "user32.dll" _
                             Alias "SendMessageA" (ByVal hWnd As Long, _
                                                   ByVal wMsg As Long, _
                                                   ByVal wParam As Long, _
                                                   lParam As vbBiDiOptions) As Long

Public Function SetParaDirection(hWndRTF As Long, Direction As Integer) As Boolean
    vbPF.cbSize = LenB(vbPF)
    'Size
    vbPF.dwMask = PFM_DIR
    'Attribute to set
    vbPF.wEffects = Direction
    'New direction
    SetParaDirection = SendPFMessage(hWndRTF, EM_SETPARAFORMAT, 0, vbPF)

End Function

