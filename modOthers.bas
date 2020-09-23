Attribute VB_Name = "modOthers"
' Win32 Declarations for Cut, Copy, Paste and Delete
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE And KEY_ENUMERATE_SUB_KEYS And KEY_NOTIFY And KEY_CREATE_SUB_KEY And KEY_CREATE_LINK And KEY_SET_VALUE
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_OPTION_VOLATILE = 1
Private Const REG_SZ = 1

Private Const ERROR_SUCCESS = 0&

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

Public cEnumValues As Collection
Public cEnumKeys As Collection

Public Function RGGetKeyValue(hKey As Long, SubKey As String, ValueName As String, Optional Default As String = "")
    
    Dim lngRet As Long
    Dim lngResult As Long
    Dim sData As String
    
    
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then
        sData = String(128, vbNullChar)
        
        lngRet = RegQueryValueEx(lngResult, ValueName, 0, REG_SZ, ByVal sData, Len(sData))
        
        If Not lngRet = ERROR_SUCCESS Then RGGetKeyValue = Default: Exit Function
        RGGetKeyValue = Left(sData, InStr(1, sData, vbNullChar) - 1)
        RegCloseKey lngResult
    Else
        RGGetKeyValue = Default
    End If
End Function


Public Function RGBHexColor(lngColor As Long) As String
Dim sHex As String
sHex = Hex(lngColor)
sHex = Right$("000000" & sHex, 6)
sHex = Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)
RGBHexColor = sHex
End Function

Public Sub LoadText(lst As TextBox, file As String)
On Error GoTo error
Dim mystr As String
Dim fNum
fNum = FreeFile
Open file For Input As #fNum
Do While Not EOF(fNum)
Line Input #fNum, a$
texto$ = texto$ + a$ + Chr$(13) + Chr$(10)
Loop
lst = texto$
Close #fNum
Exit Sub
error:
x = MsgBox("File Not Found", vbOKOnly, "Error!!")
End Sub

Public Sub FileSave(Text As String, FilePath As String)
'Save a text file
On Error GoTo error
Dim Directory As String
Dim fNum
fNum = FreeFile
Directory$ = FilePath
On Error GoTo error
Open Directory$ For Output As #fNum
Print #fNum, Text
Close #fNum
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

