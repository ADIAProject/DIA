Attribute VB_Name = "mApiOther"
Option Explicit

'*** Process ***
Private Const DONT_RESOLVE_DLL_REFERENCES As Long = &H1

Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExW" (ByVal lpLibFileName As Long, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function APIFunctionPresent
'! Description (Описание)  :   [Проверка на поддержку функции Api в текущей винде]
'! Parameters  (Переменные):   FunctionName (String)
'                              DLLName (String)
'!--------------------------------------------------------------------------------
Public Function APIFunctionPresent(ByVal FunctionName As String, ByVal DLLName As String) As Boolean

    Dim lHandle   As Long
    Dim lAddr     As Long
    Dim FreeLib   As Boolean
    Dim lngStrPtr As Long

    lngStrPtr = StrPtr(DLLName)
    lHandle = GetModuleHandle(lngStrPtr)

    If lHandle = 0 Then
        lHandle = LoadLibraryEx(lngStrPtr, 0&, DONT_RESOLVE_DLL_REFERENCES)
        FreeLib = True
    End If

    If lHandle <> 0 Then
        lAddr = GetProcAddress(lHandle, FunctionName)

        If FreeLib Then
            FreeLibrary lHandle
        End If
    End If

    APIFunctionPresent = (lAddr <> 0)
End Function
