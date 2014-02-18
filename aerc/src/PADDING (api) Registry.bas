Option Compare Database
Option Explicit

Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const REG_OPTION_NON_VOLATILE = 0

Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_READ = &H20000
Const STANDARD_RIGHTS_WRITE = &H20000
Const STANDARD_RIGHTS_EXECUTE = &H20000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_ALL = &H1F0000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20

Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = &H3F


Private Declare Function api_RegCloseKey Lib "advapi32.dll" Alias "RegCloseKey" (ByVal HKey As Long) As Long
Private Declare Function api_RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function api_RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function api_RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Declare Function api_RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long

Private Declare Function api_RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

Private Declare Function api_RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long

Private Declare Function api_RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Private Declare Function api_RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function api_RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal CbName As Long) As Long

Private Declare Function api_RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKey As Long, ByVal lpValueName As String) As Long

Private Declare Function api_RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function api_ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

Public Function m_Reg_KeyDelete(HKey As Long, sKeyToDelete As String) As Boolean
'    Delete a Key and all its SubKeys.
'    Returns True if the operation succeeds.
    
       Dim lRetVal As Long      'result
    
       lRetVal = api_RegDeleteKey(HKey, sKeyToDelete)
    
       m_Reg_KeyDelete = (lRetVal = ERROR_NONE)
    
End Function

Public Function m_Reg_ValueDelete(HKey As Long, ValueToDelete As String) As Boolean
'    Delete a Value of a Key.
'    Returns True if the operation succeeds.
    
       Dim lRetVal As Long      'result
    
       lRetVal = api_RegDeleteValue(HKey, ValueToDelete)
    
       m_Reg_ValueDelete = (lRetVal = ERROR_NONE)
    
End Function
Public Function m_Reg_KeyCreate(PredefinedKey As String, KeyName As String) As Long
'    Create a Key (and all Key needed if not existing).
'    Ex : m_Reg_KeyCreate("LM", "Software\aaa\bbb\ccc")
    
       Dim HKey As Long            'handle to the new key
       Dim lRetVal As Long         'result of the RegCreateKeyEx function
    
'    If the key allready exists, there is no error and the key is open correctly on HKey handle
       lRetVal = api_RegCreateKeyEx(p_PredifinedKey(PredefinedKey), KeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, HKey, lRetVal)
    
      If (lRetVal = ERROR_NONE) Then
          m_Reg_KeyCreate = HKey
      Else
          m_Reg_KeyCreate = 0
      End If
    
End Function

Public Function m_Reg_ValueSet(HKey As Long, ValueName As String, ValueSetting As Variant, lValueType As Long) As Boolean
    
'    Set a Value for a key allready openned
    
       Dim lRetVal As Long      'result of the SetValueEx function
    
       lRetVal = p_SetValueEx(HKey, ValueName, lValueType, ValueSetting)
    
       m_Reg_ValueSet = (lRetVal = ERROR_NONE)
    
End Function

Public Function m_Reg_ValueSetQuick(PredefinedKey As String, KeyName As String, ValueName As String, ValueSetting As Variant, lValueType As Long) As Boolean
    
'    Set a Value
    
       Dim HKey As Long
    
       HKey = m_Reg_KeyOpen(PredefinedKey, KeyName, KEY_ALL_ACCESS)
    
       If (HKey <> 0) Then
           m_Reg_ValueSetQuick = m_Reg_ValueSet(HKey, ValueName, ValueSetting, lValueType)
          m_Reg_KeyClose HKey
      Else
          m_Reg_ValueSetQuick = False
      End If
    
End Function

Public Function m_Reg_ValueGet(HKey As Long, Optional ValueName As String) As Variant
'    Returns the value of the Key.
'    Returns Null if the Key doesn't exist.
'    Returns the Default Value id ValuName is "" or is missing.
    
       Dim lRetVal As Long         'result of the API functions
       Dim xValue As Variant       'setting of queried value
    
       lRetVal = p_QueryValueEx(HKey, ValueName, xValue)
    
      If lRetVal = ERROR_NONE Then
          m_Reg_ValueGet = xValue
      Else
          m_Reg_ValueGet = Null
      End If
    
End Function

Public Function m_Reg_ValueGetQuick(PredefinedKey As String, KeyName As String, Optional ValueName As String) As Variant
    
'    Returns the value of the Key.
'    Returns Null if the Key doesn't exist.
'    Returns the Default Value id ValuName is "" or is missing.
    
       Dim HKey As Long
    
       HKey = m_Reg_KeyOpen(PredefinedKey, KeyName, KEY_READ)
       If HKey <> 0 Then
          m_Reg_ValueGetQuick = m_Reg_ValueGet(HKey, ValueName)
          m_Reg_KeyClose HKey
      Else
          m_Reg_ValueGetQuick = Null
      End If
    
End Function

Private Function p_SetValueEx(ByVal HKey As Long, ValueName As String, lType As Long, xValue As Variant) As Long
'    Writes a Value (String or DWord).
    
       Dim lValue As Long
       Dim sValue As String
    
       Select Case lType
    Case REG_SZ, REG_EXPAND_SZ
           sValue = xValue & Chr$(0)
           p_SetValueEx = api_RegSetValueExString(HKey, ValueName, 0&, lType, sValue, Len(sValue))
      Case REG_DWORD
          lValue = xValue
          p_SetValueEx = api_RegSetValueExLong(HKey, ValueName, 0&, lType, lValue, 4)
      End Select
    
End Function

Private Function p_QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, xValue As Variant) As Long
'    Returns the data of a Value (for a given Kay).
    
       Dim cch     As Long
       Dim lrc     As Long
       Dim lType   As Long
       Dim lValue  As Long
       Dim sValue  As String
    
       On Error GoTo QueryValueExError
    
'    Determine the size and type of data to be read
      lrc = api_RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
      If lrc <> ERROR_NONE Then Error 5
    
      Select Case lType
'        For strings
    Case REG_SZ, REG_EXPAND_SZ
          sValue = String(cch, 0)
          lrc = api_RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
          If lrc = ERROR_NONE Then
              xValue = Left$(sValue, cch - 1)
              If lType = REG_EXPAND_SZ Then
                  xValue = m_Reg_ExpandString("" & xValue)
              End If
          End If
'        For DWORDS
      Case REG_DWORD
          lrc = api_RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
          If lrc = ERROR_NONE Then
              xValue = lValue
          End If
      Case Else
'        all other data types not supported
      End Select
    
QueryValueExExit:
      p_QueryValueEx = lrc
    Exit Function
    
QueryValueExError:
      Resume QueryValueExExit
    
End Function

Private Function p_PredifinedKey(PredifinedKey As String) As Long
'    Returns the Id of a predifined Key.
'    Both Nikename or Long name can be given.
    
       Dim x As Long
    
       Select Case PredifinedKey
    Case "CR", "HKEY_CLASSES_ROOT"
           x = HKEY_CLASSES_ROOT
       Case "CU", "HKEY_CURRENT_USER"
          x = HKEY_CURRENT_USER
      Case "LM", "HKEY_LOCAL_MACHINE"
          x = HKEY_LOCAL_MACHINE
      Case "U", "HKEY_USERS"
          x = HKEY_USERS
      Case Else
          x = HKEY_CURRENT_USER
      End Select
    
      p_PredifinedKey = x
    
End Function

Public Function m_Reg_ValueEnum(HKey As Long, Item As Long) As String
'    Returns tha Name of the Item# Value of the openned Key (better if the key is open with KEY_READ mode)..
'    The first Value is item 0.
'    The Default Value is never given (its name is "").
'    Returns "" if the item is over the number of Values in the Key.
    
       Dim lRetVal       As Long
       Dim ValueName    As String
       Dim lValueNameLen As Long
       Dim lData         As Long
      Dim lDataLen      As Long
      Dim lValueType    As Long
    
      lValueNameLen = 2000
      ValueName = String(lValueNameLen, 0)
      lDataLen = 2000
    
      lRetVal = api_RegEnumValue(HKey, Item, ByVal ValueName, lValueNameLen, 0&, lValueType, ByVal lData, lDataLen)
    
      If lRetVal = ERROR_NONE Then
          m_Reg_ValueEnum = Left(ValueName, lValueNameLen)
      Else
          m_Reg_ValueEnum = ""
      End If
    
End Function

Public Function m_Reg_KeyEnum(HKey As Long, Item As Long) As String
'    Returns the name of the Item# SubKey of the openned Key (better if the key is open with KEY_READ mode).
'    The first SubKey is item 0.
'    Returns "" if the item is over the number of SubKeys in the Key.
    
       Dim lRetVal       As Long
       Dim sSubKeyName    As String
       Dim lSubKeyNameLen As Long
    
       lSubKeyNameLen = 2000
      sSubKeyName = String(lSubKeyNameLen, 0)
    
      lRetVal = api_RegEnumKey(HKey, Item, ByVal sSubKeyName, lSubKeyNameLen)
    
      If lRetVal = ERROR_NONE Then
          m_Reg_KeyEnum = Left(sSubKeyName, lSubKeyNameLen)
      Else
          m_Reg_KeyEnum = vbNullString
      End If
    
End Function

Public Function m_Reg_KeyClose(HKey As Long) As Boolean
'    Close an open key, and returns True if succeed
    
       Dim lRetVal As Long
    
       lRetVal = api_RegCloseKey(HKey)
    
       If lRetVal = ERROR_NONE Then
           m_Reg_KeyClose = True
       End If
    
End Function

Public Function m_Reg_KeyOpen(PredefinedKey As String, KeyName As String, Optional AccessMode As Long) As Long
'    Open a key and returns the handle
    
       Dim HKey    As Long
       Dim lRetVal As Long
    
       If AccessMode = 0 Then
           AccessMode = KEY_ALL_ACCESS
       End If
    
      lRetVal = api_RegOpenKeyEx(p_PredifinedKey(PredefinedKey), KeyName, 0, AccessMode, HKey)
    
      If lRetVal = ERROR_NONE Then
          m_Reg_KeyOpen = HKey
      End If
    
End Function

Public Function m_Reg_ExpandString(Txt As String) As String
    
       Dim lrc     As Long
       Dim sValue  As String
       Dim Pos     As Long
    
       sValue = String(Len(Txt) + 400, Chr$(0))
       lrc = api_ExpandEnvironmentStrings(Txt, sValue, Len(sValue))
       If (lrc > 0) Then
           Pos = InStr(sValue, Chr$(0))
          sValue = Left$(sValue, Pos - 1)
      Else
          sValue = ""
      End If
    
      m_Reg_ExpandString = sValue
    
End Function