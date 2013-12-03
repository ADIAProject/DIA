Attribute VB_Name = "mApiProcess"
Option Explicit

' необходимо для регистрации компонента
Public Const DONT_RESOLVE_DLL_REFERENCES As Long = &H1
Public Const GMEM_FIXED                 As Long = 0    'Fixed memory GlobalAlloc flag
Public Const PATCH_04                   As Long = 88                                   'Table B (before) address patch offset
Public Const PATCH_05                   As Long = 93                                   'Table B (before) entry count patch offset
Public Const PATCH_08                   As Long = 132                                  'Table A (after) address patch offset
Public Const PATCH_09                   As Long = 137                                  'Table A (after) entry count patch offset

Public Declare Sub RtlMoveMemory _
                    Lib "kernel32.dll" (Destination As Any, _
                                        Source As Any, _
                                        ByVal Length As Long)

Public Declare Sub CopyMemory _
                    Lib "kernel32.dll" _
                        Alias "RtlMoveMemory" (Destination As Any, _
                                               Source As Any, _
                                               ByVal Length As Long)

Public Declare Sub CopyMemoryLong _
                    Lib "kernel32.dll" _
                        Alias "RtlMoveMemory" (ByVal Destination As Long, _
                                               ByVal Source As Long, _
                                               ByVal Length As Long)

Public Declare Function GlobalAlloc _
                         Lib "kernel32.dll" (ByVal wFlags As Long, _
                                             ByVal dwBytes As Long) As Long

Public Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long

Public Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExW" (ByVal lpLibFileName As Long, _
                                                     ByVal hFile As Long, _
                                                     ByVal dwFlags As Long) As Long

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long

Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Public Declare Function CallWindowProc _
                         Lib "user32.dll" _
                             Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                      ByVal hWnd As Long, _
                                                      ByVal Msg As Long, _
                                                      ByVal wParam As Long, _
                                                      ByVal lParam As Long) As Long

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Public Declare Function OpenProcess _
                         Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, _
                                             ByVal bInheritHandle As Long, _
                                             ByVal dwProcessId As Long) As Long

Public Declare Function WriteProcessMemory _
                         Lib "kernel32.dll" (ByVal hProcess As Long, _
                                             lpBaseAddress As Any, _
                                             lpBuffer As Any, _
                                             ByVal nSize As Long, _
                                             Optional lpNumberOfBytesWritten As Long) As Long

Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long


Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long