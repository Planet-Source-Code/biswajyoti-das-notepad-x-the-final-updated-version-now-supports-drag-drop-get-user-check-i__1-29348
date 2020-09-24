Attribute VB_Name = "RegCool"
Public Const HKEY_CURRENT_USER = &H80000001
Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type
Public GetSound As String, JKDir As String, vbNewLine As String, ColorUP As String, Light As Integer, Dark As Integer
Public WinDir As String
Public Type SHITEMID
cb As Long
abID As Byte
End Type
Public Type ITEMIDLIST
mkid As SHITEMID
End Type
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&
Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ
Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte
Const DisplayErrorMsg = False
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal lngHKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal lngHKey As Long, ByVal lpClass As String, _
ByVal lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, _
lpcbMaxSubKeyLen As Long, ByVal lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, _
ByVal lpcbMaxValueLen As Long, ByVal lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Const mcregOptionNonVolatile = 0
Private Const mcregErrorNone = 0
Private Const mcregErrorBadDB = 1
Private Const mcregErrorBadKey = 2
Private Const mcregErrorCantOpen = 3
Private Const mcregErrorCantRead = 4
Private Const mcregErrorCantWrite = 5
Private Const mcregErrorOutOfMemory = 6
Private Const mcregErrorInvalidParameter = 7
Private Const mcregErrorAccessDenied = 8
Private Const mcregErrorInvalidParameterS = 87
Private Const mcregErrorNoMoreItems = 259
Private Const mcregKeyAllAccess = &H3F
Private Const mcregKeyQueryValue = &H1
Function GetStringValue(SubKey As String, Entry As String)
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
If rtn = ERROR_SUCCESS Then
sBuffer = Space(255)
lBufferSize = Len(sBuffer)
rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize)
If rtn = ERROR_SUCCESS Then
rtn = RegCloseKey(hKey)
sBuffer = Trim(sBuffer)
GetStringValue = Left(sBuffer, Len(sBuffer) - 1)
Else
GetStringValue = "Error"
If DisplayErrorMsg = True Then
End If
End If
Else
GetStringValue = "Errore"
If DisplayErrorMsg = True Then
End If
End If
End If
End Function
Public Sub ParseKey(Keyname As String, Keyhandle As Long)
rtn = InStr(Keyname, "\")
If Left(Keyname, 5) <> "HKEY_" Or Right(Keyname, 1) = "\" Then
MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + Keyname
Exit Sub
ElseIf rtn = 0 Then
Keyhandle = GetMainKeyHandle(Keyname)
Keyname = ""
Else
Keyhandle = GetMainKeyHandle(Left(Keyname, rtn - 1))
Keyname = Right(Keyname, Len(Keyname) - rtn)
End If
End Sub
Function GetMainKeyHandle(MainKeyName As String) As Long
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
Select Case MainKeyName
Case "HKEY_CLASSES_ROOT"
GetMainKeyHandle = HKEY_CLASSES_ROOT
Case "HKEY_CURRENT_USER"
GetMainKeyHandle = HKEY_CURRENT_USER
Case "HKEY_LOCAL_MACHINE"
GetMainKeyHandle = HKEY_LOCAL_MACHINE
Case "HKEY_USERS"
GetMainKeyHandle = HKEY_USERS
Case "HKEY_PERFORMANCE_DATA"
GetMainKeyHandle = HKEY_PERFORMANCE_DATA
Case "HKEY_CURRENT_CONFIG"
GetMainKeyHandle = HKEY_CURRENT_CONFIG
Case "HKEY_DYN_DATA"
GetMainKeyHandle = HKEY_DYN_DATA
End Select
End Function
Function SetStringValue(SubKey As String, Entry As String, Value As String)
Call ParseKey(SubKey, MainKeyHandle)
If MainKeyHandle Then
rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)
If rtn = ERROR_SUCCESS Then
rtn = RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value))
If Not rtn = ERROR_SUCCESS Then
If DisplayErrorMsg = True Then
End If
End If
rtn = RegCloseKey(hKey)
Else
If DisplayErrorMsg = True Then
End If
End If
End If
End Function
Public Sub RegistryCreateNewKey(eRootKey As String, strKeyName As String)
  Dim lngRetVal As Long
  Dim lngHKey As Long
  On Error GoTo PROC_ERR
  lngRetVal = RegCreateKeyEx(eRootKey, strKeyName, 0&, vbNullString, mcregOptionNonVolatile, mcregKeyAllAccess, 0&, lngHKey, 0&)
  If lngRetVal = mcregErrorNone Then
    RegCloseKey (lngHKey)
  End If
PROC_EXIT:
  Exit Sub
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RegistryCreateNewKey"
  Resume PROC_EXIT
End Sub


