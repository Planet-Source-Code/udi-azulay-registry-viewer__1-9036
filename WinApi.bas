Attribute VB_Name = "WinApi"
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Public Const KEY_CREATE_LINK = &H20
    Public Const KEY_CREATE_SUB_KEY = &H4
    Public Const KEY_ENUMERATE_SUB_KEYS = &H8
    Public Const KEY_EVENT = &H1     '  Event contains key event record
    Public Const KEY_NOTIFY = &H10
    Public Const KEY_QUERY_VALUE = &H1
    Public Const KEY_SET_VALUE = &H2
    
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_CONFIG = &H80000005
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_DYN_DATA = &H80000006
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_PERFORMANCE_DATA = &H80000004
    Public Const HKEY_USERS = &H80000003

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long

Public Type ByteArray
  FirstByte As Byte
  ByteBuffer(100) As Byte
End Type

Public Const REG_NONE = 0                       ' No value type
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Public Const REG_LINK = 6                       ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Public Const REG_RESOURCE_LIST = 8              ' Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Public Const REG_CREATED_NEW_KEY = &H1                          ' New Registry Key created
Public Const REG_OPENED_EXISTING_KEY = &H2                      ' Existing Key opened
Public Const REG_WHOLE_HIVE_VOLATILE = &H1                      ' Restore whole hive volatile
Public Const REG_REFRESH_HIVE = &H2                             ' Unwind changes to last flush
Public Const REG_NOTIFY_CHANGE_NAME = &H1                       ' Create or delete (child)
Public Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Public Const REG_NOTIFY_CHANGE_LAST_SET = &H4                      ' Time stamp
Public Const REG_NOTIFY_CHANGE_SECURITY = &H8
Public Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)

