Attribute VB_Name = "Registry"
Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4
Public Const REG_EXPAND_SZ = 2

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259

Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_ALL_ACCESS = &H3F

Public Const REG_OPTION_NON_VOLATILE = 0

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
       (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
       ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
       ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
       (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
       ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
       (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
       lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
       (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
       lpType As Long, lpData As Long, lpcbData As Long) As Long
Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
       (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
       lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
       (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
       ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
       (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
       ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Public Sub CreateNewKey(strKey As String, lngHKey As Long)

  Dim hNewKey As Long
  Dim lngRC As Long

  lngRC = RegCreateKeyEx(lngHKey, strKey, 0&, vbNullString, _
          REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lngRC)
   
  RegCloseKey (hNewKey)
  
End Sub

Public Sub SetKeyValue(ByVal strKey As String, ByVal strValue As String, _
       strSetting As String, lngHKey As Long, Optional typ As Long = REG_SZ)
     
    Dim lngRC As Long       'result of the SetValueEx function
    Dim hNewKey As Long         'handle of open key
    
    'Open the key
    lngRC = RegOpenKeyEx(lngHKey, strKey, 0&, _
                                 KEY_SET_VALUE, hNewKey)
    
    'Put the value
    
    If typ = REG_SZ Then
    
    lngRC = RegSetValueExString(hNewKey, _
            strValue, 0&, REG_SZ, strSetting, Len(strSetting))
    
    Else
    Dim lsetting As Long
    lsetting = CLng(strSetting)
    lngRC = RegSetValueExLong(hNewKey, strValue, 0, REG_DWORD, lsetting, Len(lsetting))
    End If
    
    
    
    
    'Close the key
    RegCloseKey (hNewKey)
    
End Sub

Public Function SetValueEx(ByVal hKey As Long, ByVal _
                strValue As String, ByVal sValue As String) As Long
  
  sValue = vValue & Chr$(0)
  Dim ltype As Long
  SetValueEx = RegSetValueExString(hKey, strValue, 0&, ltype, sValue, Len(sValue))
  SetValueEx = ltype
End Function

Public Function QueryValue(strKey As String, _
                strValue As String, lPredefinedKey As Long, Optional typ As Long = REG_SZ) As String

  Dim lngRC As Long
  Dim hKey As Long
  Dim strSetting As String
  
  
  'Get the key handle
  lngRC = RegOpenKeyEx(lPredefinedKey, strKey, 0, _
    KEY_QUERY_VALUE, hKey)
    
  'Get the value
  lngRC = QueryValueEx(hKey, strValue, strSetting, typ)
  
  RegCloseKey (hKey)
  QueryValue = strSetting
  
End Function




Public Function QueryValueEx(ByVal hKey As Long, _
       ByVal strValue As String, ByRef strSetting As String, typ As Long) As Long
  
  Dim lngRC As Long
  Dim lngChData As Long
  
  'Get the length, zero if error
  lngRC = RegQueryValueExNULL(hKey, strValue, 0&, REG_SZ, 0&, lngChData)
  If lngRC <> ERROR_NONE Then
     Call PutError(lngRC)
     QueryValueEx = lngRC
     GoTo xt_QueryValueEx
     
  ElseIf typ = Registry.REG_SZ Then
    strSetting = Space$(lngChData)
    
    lngRC = RegQueryValueExString(hKey, strValue, 0&, REG_SZ, _
                strSetting, lngChData)
  ElseIf typ = Registry.REG_DWORD Then
    Dim longSetting As Long
    lngRC = RegQueryValueExLong(hKey, strValue, 0&, REG_SZ, _
                longSetting, lngChData)
    strSetting = longSetting
  End If
  
  If lngRC = ERROR_NONE Then
     
     If Len(strSetting) > 0 Then
     strSetting = Left$(strSetting, lngChData - 1)
     End If
     
  Else
     Call PutError(lngRC)
     strSetting = ""
  End If
  
xt_QueryValueEx:
  
  QueryValueEx = lngRC
End Function

Public Function GetError(ByVal lngErrorCode As Long) As String
  
  Select Case lngErrorCode
    Case ERROR_BADDB
      GetError = "Bad DB"
    Case ERROR_BADKEY
      GetError = "Bad Key"
    Case ERROR_CANTOPEN
      GetError = "Can't Open"
    Case ERROR_CANTREAD
      GetError = "Can't Read"
    Case ERROR_CANTWRITE
      GetError = "Can't Write"
    Case ERROR_OUTOFMEMORY
      GetError = "Out of Memory"
    Case ERROR_ARENA_TRASHED
      GetError = "Arena Trashed" 'Ooo that sounds like a bad one!
    Case ERROR_ACCESS_DENIED
      GetError = "Access Denied"
    Case ERROR_INVALID_PARAMETERS
      GetError = "Invalid Parameters"
    Case ERROR_NO_MORE_ITEMS
      GetError = "No more items"
    Case Else
      GetError = "Who knows what happened but it's bad!"
  End Select
  
End Function

Public Sub PutError(ByVal lngError As Long)

  Dim strMessage As String
  
  strMessage = "Error occurred - " & GetError(lngError) & vbCrLf & vbCrLf
  strMessage = strMessage & "Be sure to: " & vbCrLf
  strMessage = strMessage & " 1) Create the key" & vbCrLf
  strMessage = strMessage & " 2) Set the Key Value" & vbCrLf
  strMessage = strMessage & " 3) Then you can query the value."
  
  MsgBox strMessage
  
End Sub

