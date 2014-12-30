VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Security Mask constants
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or _
                                 KEY_QUERY_VALUE Or _
                                 KEY_SET_VALUE Or _
                                 KEY_CREATE_SUB_KEY Or _
                                 KEY_ENUMERATE_SUB_KEYS Or _
                                 KEY_NOTIFY Or _
                                 KEY_CREATE_LINK) And _
                                 (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
                           KEY_QUERY_VALUE Or _
                           KEY_ENUMERATE_SUB_KEYS Or _
                           KEY_NOTIFY) And _
                           (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or _
                           KEY_SET_VALUE Or _
                           KEY_CREATE_SUB_KEY) And _
                           (Not SYNCHRONIZE))

' Possible registry data types
Public Enum InTypes
   ValNull = 0
   ValString = 1
   ValXString = 2
   ValBinary = 3
   ValDWord = 4
   ValLink = 6
   ValMultiString = 7
   ValResList = 8
End Enum

' Registry value type definitions
Private Const REG_NONE As Long = 0
Private Const REG_SZ As Long = 1
Private Const REG_EXPAND_SZ As Long = 2
Private Const REG_BINARY As Long = 3
Private Const REG_DWORD As Long = 4
Private Const REG_LINK As Long = 6
Private Const REG_MULTI_SZ As Long = 7
Private Const REG_RESOURCE_LIST As Long = 8

' Registry section definitions
Public Enum RegGroup
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_PERFORMANCE_DATA = &H80000004
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_DYN_DATA = &H80000006
End Enum

'Private Const HKEY_CLASSES_ROOT = &H80000000
'Private Const HKEY_CURRENT_USER = &H80000001
'Private Const HKEY_LOCAL_MACHINE = &H80000002
'Private Const HKEY_USERS = &H80000003
'Private Const HKEY_PERFORMANCE_DATA = &H80000004
'Private Const HKEY_CURRENT_CONFIG = &H80000005
'Private Const HKEY_DYN_DATA = &H80000006

' Codes returned by Reg API calls
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

' Registry API functions used in this module (there are more of them)
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

' This routine allows you to get values from anywhere in the Registry, it currently
' only handles string, double word and binary values. Binary values are returned as
' hex strings.
'
' Example
' Text1.Text = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultUserName")
'
Public Function ReadRegistry(ByVal Group As RegGroup, _
                             ByVal Section As String, _
                             ByVal Key As String) As String
   On Error Resume Next
   Dim lResult As Long
   Dim lKeyValue As Long
   Dim lDataTypeValue As Long
   Dim lValueLength As Long
   Dim sValue As String
   Dim td As Double
   Dim tStr1 As String
   Dim tStr2 As String
   Dim i As Integer

   lResult = RegOpenKeyEx(Group, Section, 0, KEY_READ, lKeyValue)
   sValue = Space$(2048)
   lValueLength = Len(sValue)
   lResult = RegQueryValueEx(lKeyValue, Key, 0&, lDataTypeValue, sValue, lValueLength)
   
   If (lResult = 0) And (Err.Number = 0) Then
      If lDataTypeValue = REG_DWORD Then
         td = Asc(Mid$(sValue, 1, 1)) + &H100& * Asc(Mid$(sValue, 2, 1)) + &H10000 * Asc(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(sValue, 4, 1)))
         sValue = Format$(td, "000")
      End If
      
      If lDataTypeValue = REG_BINARY Then
'         Return a binary field as a hex string (2 chars per byte)
          tStr2 = ""
          For i = 1 To lValueLength
             tStr1 = Hex(Asc(Mid(sValue, i, 1)))
             If Len(tStr1) = 1 Then tStr1 = "0" & tStr1
             tStr2 = tStr2 + tStr1
          Next
          sValue = tStr2
      Else
         sValue = Left$(sValue, lValueLength - 1)
      End If
   Else
      sValue = "Not Found"
   End If
   
   lResult = RegCloseKey(lKeyValue)
   ReadRegistry = sValue
End Function

Public Sub EditRegistryKey(ByVal Group As RegGroup, _
                           ByVal Section As String, _
                           ByVal Key As String, _
                           ByVal KeyType As InTypes, _
                           ByVal KeyVal As Variant)
   Dim lNewVal As Long
   Dim sNewVal As Variant
   Dim inLen As Long
   Dim lResult As Long
   Dim keyHandle As Long
   
   lResult = RegOpenKeyEx(Group, Section, 0, KEY_WRITE, keyHandle)
   
   If lResult = ERROR_NONE Then
      If KeyType = ValDWord Then
         lNewVal = CLng(KeyVal)
         inLen = 4
         lResult = RegSetValueExLong(keyHandle, Key, 0&, KeyType, lNewVal, inLen)
      Else
         If KeyType = ValString Then
            KeyVal = KeyVal + Chr(0)
         End If
         
         sNewVal = KeyVal
         inLen = Len(sNewVal)
         lResult = RegSetValueExString(keyHandle, Key, 0&, KeyType, sNewVal, inLen)
      End If
      lResult = RegFlushKey(keyHandle)
      RegCloseKey keyHandle
   End If
End Sub
                           

' This routine allows you to write values into the entire Registry, it currently
' only handles string and double word values.
'
' Example
' WriteRegistry HKEY_CURRENT_USER, "SOFTWARE\My Name\My App\", "NewSubKey", ValString, "NewValueHere"
' WriteRegistry HKEY_CURRENT_USER, "SOFTWARE\My Name\My App\", "NewSubKey", ValDWord, "31"
'
Public Sub WriteRegistry(ByVal Group As RegGroup, _
                         ByVal Section As String, _
                         ByVal Key As String, _
                         ByVal ValType As InTypes, _
                         ByVal Value As Variant)
   On Error Resume Next
   Dim lResult As Long
   Dim lKeyValue As Long
   Dim inLen As Long
   Dim lNewVal As Long
   Dim sNewVal As String

   lResult = RegCreateKey(Group, Section, lKeyValue)
   
   If ValType = ValDWord Then
      lNewVal = CLng(Value)
      inLen = 4
      lResult = RegSetValueExLong(lKeyValue, Key, 0&, ValType, lNewVal, inLen)
   Else
'     Fixes empty string bug - spotted by Marcus Jansson
      If ValType = ValString Then
         Value = Value + Chr(0)
      End If
      
      sNewVal = Value
      inLen = Len(sNewVal)
      lResult = RegSetValueExString(lKeyValue, Key, 0&, 1&, sNewVal, inLen)
   End If
   
   lResult = RegFlushKey(lKeyValue)
   lResult = RegCloseKey(lKeyValue)
End Sub

' This routine enumerates the subkeys under any given key
' Call repeatedly until "Not Found" is returned - store values in array or something
'
' Example - this example just adds all the subkeys to a string - you will probably want to
' save then into an array or something.
'
' Dim Res As String
' Dim i As Long
' Res = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\", i)
' Do Until Res = "Not Found"
'   Text1.Text = Text1.Text & " " & Res
'   i = i + 1
'   Res = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\", i)
' Loop

Public Function ReadRegistryGetSubkey(ByVal Group As RegGroup, _
                                      ByVal Section As String, _
                                      Idx As Long) As String
On Error Resume Next
   Dim lResult As Long
   Dim lKeyValue As Long
   Dim lDataTypeValue As Long
   Dim lValueLength As Long
   Dim sValue As String
   Dim td As Double

   lResult = RegOpenKey(Group, Section, lKeyValue)
   sValue = Space$(2048)
   lValueLength = Len(sValue)
   lResult = RegEnumKey(lKeyValue, Idx, sValue, lValueLength)
   
   If (lResult = 0) And (Err.Number = 0) Then
      sValue = Left$(sValue, InStr(sValue, Chr(0)) - 1)
   Else
      sValue = "Not Found"
   End If
   
   lResult = RegCloseKey(lKeyValue)
   ReadRegistryGetSubkey = sValue
End Function

' This routine allows you to get all the values from anywhere in the Registry under any
' given subkey, it currently only returns string and double word values.
'
' Example - returns list of names/values to multiline text box
' Dim Res As Variant
' Dim i As Long
' Res = ReadRegistryGetAll(HKEY_CURRENT_USER, "Software\Microsoft\Notepad", i)
' Do Until Res(2) = "Not Found"
'    Text1.Text = Text1.Text & Chr(13) & Chr(10) & Res(1) & " " & Res(2)
'    i = i + 1
'    Res = ReadRegistryGetAll(HKEY_CURRENT_USER, "Software\Microsoft\Notepad", i)
' Loop
'
'Public Function ReadRegistryGetAll(ByVal Group As RegGroup, _
'                                   ByVal Section As String, _
'                                   Idx As Long) As Variant
'   On Error Resume Next
'   Dim lResult As Long
'   Dim lKeyValue As Long
'   Dim DataTypeValue As Long
'   Dim lValueLength As Long
'   Dim lValueNameLength As Long
'   Dim sValueName As String
'   Dim sValue As String
'   Dim td As Double
'
'   lResult = RegOpenKey(Group, Section, lKeyValue)
'   sValue = Space$(2048)
'   sValueName = Space$(2048)
'   lValueLength = Len(sValue)
'   lValueNameLength = Len(sValueName)
'   lResult = RegEnumValue(lKeyValue, Idx, sValueName, lValueNameLength, 0&, lDataTypeValue, sValue, lValueLength)
'
'   If (lResult = 0) And (Err.Number = 0) Then
'      If lDataTypeValue = REG_DWORD Then
'         td = Asc(Mid$(sValue, 1, 1)) + &H100& * Asc(Mid$(sValue, 2, 1)) + &H10000 * Asc(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(sValue, 4, 1)))
'         sValue = Format$(td, "000")
'      End If
'
'      sValue = Left$(sValue, lValueLength - 1)
'      sValueName = Left$(sValueName, lValueNameLength)
'   Else
'      sValue = "Not Found"
'   End If
'
'   lResult = RegCloseKey(lKeyValue)
'
''  Return the datatype, value name and value as an array
'   ReadRegistryGetAll = Array(lDataTypeValue, sValueName, sValue)
'End Function

' This routine deletes a specified key (and all its subkeys and values if on Win95) from the registry.
' Be very careful using this function.
'
' Example
' DeleteSubkey HKEY_CURRENT_USER, "Software\My Name\My App"
'
Public Function DeleteSubkey(ByVal Group As RegGroup, _
                             ByVal Section As String) As String
   On Error Resume Next
   Dim lResult As Long
   Dim lKeyValue As Long

   lResult = RegOpenKeyEx(Group, vbNullChar, 0&, KEY_ALL_ACCESS, lKeyValue)
   lResult = RegDeleteKey(lKeyValue, Section)
   lResult = RegCloseKey(lKeyValue)
End Function

' This routine deletes a specified value from below a specified subkey.
' Be very careful using this function.
'
' Example
' DeleteValue HKEY_CURRENT_USER, "Software\My Name\My App", "NewSubKey"
'
Public Function DeleteValue(ByVal Group As RegGroup, _
                            ByVal Section As String, _
                            ByVal Key As String) As String
   On Error Resume Next
   Dim lResult As Long
   Dim lKeyValue As Long

   lResult = RegOpenKey(Group, Section, lKeyValue)
   lResult = RegDeleteValue(lKeyValue, Key)
   lResult = RegCloseKey(lKeyValue)
End Function
