Attribute VB_Name = "mOLECommon"
'    Copyright 2010 Makarov Andrey
Option Explicit

Public Declare Sub OleInitialize Lib "ole32.dll" (pvReserved As Any)
Public Declare Sub OleUninitialize Lib "ole32.dll" ()
Public Declare Function CLSIDFromString _
                         Lib "ole32.dll" (ByVal lpsz As String, _
                                          pCLSID As Guid) As Long

Public Declare Function IIDFromString _
                         Lib "ole32.dll" (ByVal lpsz As String, _
                                          lpiid As Guid) As Long

Public Declare Function CoCreateInstance _
                         Lib "ole32.dll" (rclsid As Guid, _
                                          ByVal pUnkOuter As Long, _
                                          ByVal dwClsContext As Long, _
                                          riid As Guid, _
                                          ppv As Any) As Long

Private Declare Function CallWindowProc _
                          Lib "user32.dll" _
                              Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                       ByVal hWnd As Long, _
                                                       ByVal Msg As Long, _
                                                       ByVal wParam As Long, _
                                                       ByVal lParam As Long) As Long

Private Declare Function PutMem2 _
                          Lib "msvbvm60" (ByVal pWORDDst As Long, _
                                          ByVal NewValue As Long) As Long

Private Declare Function PutMem4 _
                          Lib "msvbvm60" (ByVal pDWORDDst As Long, _
                                          ByVal NewValue As Long) As Long

Private Declare Function GetMem4 _
                          Lib "msvbvm60" (ByVal pDWORDSrc As Long, _
                                          ByVal pDWORDDst As Long) As Long

Private Declare Function GlobalAlloc _
                          Lib "kernel32.dll" (ByVal wFlags As Long, _
                                              ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Private Const GMEM_FIXED                As Long = &H0
Private Const asmPUSH_imm32             As Byte = &H68
Private Const asmRET_imm16              As Byte = &HC2
Private Const asmCALL_rel32             As Byte = &HE8

Public Type Guid
    Data1                               As Long
    Data2                               As Integer
    Data3                               As Integer
    Data4(0 To 7)                       As Byte
End Type

Public Const TBPF_NOPROGRESS = 0
Public Const TBPF_INDETERMINATE = 1
Public Const TBPF_NORMAL = 2
Public Const TBPF_ERROR = 4
Public Const TBPF_PAUSED = 8
Public Const unk_QueryInterface         As Long = 0
Public Const unk_AddRef                 As Long = 1
Public Const unk_Release                As Long = 2

Public Function CallInterface(ByVal pInterface As Long, _
                              ByVal Member As Long, _
                              ByVal ParamsCount As Long, _
                              Optional ByVal P1 As Long = 0, _
                              Optional ByVal P2 As Long = 0, _
                              Optional ByVal P3 As Long = 0, _
                              Optional ByVal P4 As Long = 0, _
                              Optional ByVal p5 As Long = 0, _
                              Optional ByVal p6 As Long = 0, _
                              Optional ByVal p7 As Long = 0, _
                              Optional ByVal p8 As Long = 0, _
                              Optional ByVal p9 As Long = 0, _
                              Optional ByVal p10 As Long = 0) As Long

Dim i As Long, t                        As Long
Dim hGlobal As Long, hGlobalOffset      As Long

    If ParamsCount < 0 Then err.Raise 5

    'invalid call
    If pInterface = 0 Then err.Raise 5
    '5 байт для запихивания каждого параметра в стек
    '5 байт - PUSH this
    '5 байт - вызов мембера
    '3 байта - ret 0x0010, выпихивая при этом и параметры CallWindowProc
    '1 байт - выравнивание, поскольку последний PutMem4 требует 4 байта.
    hGlobal = GlobalAlloc(GMEM_FIXED, 5 * ParamsCount + 5 + 5 + 3 + 1)

    If hGlobal = 0 Then err.Raise 7
    'insuff. memory
    hGlobalOffset = hGlobal

    If ParamsCount > 0 Then
        t = VarPtr(P1)

        For i = ParamsCount - 1 To 0 Step -1
            PutMem2 hGlobalOffset, asmPUSH_imm32
            hGlobalOffset = hGlobalOffset + 1
            GetMem4 t + i * 4, hGlobalOffset
            hGlobalOffset = hGlobalOffset + 4
        Next

    End If

    'Первый параметр любого интерфейсного метода - this. Делаем...
    PutMem2 hGlobalOffset, asmPUSH_imm32
    hGlobalOffset = hGlobalOffset + 1
    PutMem4 hGlobalOffset, pInterface
    hGlobalOffset = hGlobalOffset + 4
    'Вызов мембера интерфейса
    PutMem2 hGlobalOffset, asmCALL_rel32
    hGlobalOffset = hGlobalOffset + 1
    GetMem4 pInterface, VarPtr(t)
    'дереференс: находим положение vTable
    GetMem4 t + Member * 4, VarPtr(t)
    'смещение по vTable, после чего дереференс оного
    PutMem4 hGlobalOffset, t - hGlobalOffset - 4
    hGlobalOffset = hGlobalOffset + 4
    'Интерфейсы stdcall. Поэтому не будем cdecl учитывать.
    PutMem4 hGlobalOffset, &H10C2&
    'ret 0x0010
    CallInterface = CallWindowProc(hGlobal, 0, 0, 0, 0)
    GlobalFree hGlobal

End Function
