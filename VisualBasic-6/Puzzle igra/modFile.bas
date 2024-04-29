Attribute VB_Name = "modFile"
Option Explicit

'Module Purpose:
'***************
'   Helps you associate different file types with your
'   programs.

'Module provided by Sarfraz Ahmed Chanido

'Group:
'http://groups.msn.com/SindhiComputing


Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_NOTIFY = &H10
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_QUERY_VALUE = &H1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE And KEY_ENUMERATE_SUB_KEYS And KEY_NOTIFY And KEY_CREATE_SUB_KEY And KEY_CREATE_LINK And KEY_SET_VALUE
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1                        ' Unicode nul terminated string
Public Const REG_DWORD = 4

Public Function Associate(Program As String, Extension As String, Description As String, Optional Icon As String)
    '** Description:
    '** Associate file with Program
    
    '** Example
    '   Call Associate(App.Path & "\progname.exe", "txt", "Text Document", App.Path & "\iconname.ico")
    
    RGCreateKey HKEY_CLASSES_ROOT, "." & Extension
    RGSetKeyValue HKEY_CLASSES_ROOT, "." & Extension, "", Extension & "file"
    
    RGCreateKey HKEY_CLASSES_ROOT, Extension & "file"
    RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell"
    If LCase(Extension) = "bat" Then
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\edit"
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\edit\command"
        RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file\shell\edit\command", "", Program & " " & "%1" 'Set file path
    Else
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\open"
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\open\command"
        RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file\shell\open\command", "", Program & " " & "%1" 'Set file path
    End If
    RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\DefaultIcon"
    
    RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file", "", Description 'Set file description
    RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file\DefaultIcon", "", Icon 'Set file icon
End Function

Public Function RGCreateKey(hKey As Long, SubKey As String)
    '** Description:
    '** Create a new key
    
    Dim lngRet As Long
    Dim lngResult As Long
    Dim lngDis As Long
    
    ' Create a new key
    lngRet = RegCreateKeyEx(hKey, SubKey, 0&, 0&, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngResult, lngDis)
    lngRet = RegCloseKey(lngResult) 'Close key
End Function

Public Function RGSetKeyValue(hKey As Long, SubKey As String, ValueName As String, sValue As String)
    '** Description:
    '** Set key value
        
    Dim lngRet As Long
    Dim lngResult As Long
    
    ' Open key
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then 'If key exist
        ' Set key value
        RegSetValueEx lngResult, ValueName, 0, REG_SZ, ByVal sValue, Len(sValue)
        RegFlushKey lngResult 'Update registry
        RegCloseKey lngResult 'Close key
    End If
End Function

Public Sub DeleteKey(ByVal hKey As Long, ByVal Key As String)
Dim Temp As Long

'To delete association use these two lines:

'DeleteKey HKEY_CLASSES_ROOT, "." & "txt"
'DeleteKey HKEY_CLASSES_ROOT, "txt" & "file"

'-->Add your extention where "txt" appears above.

Temp = RegDeleteKey(hKey, Key)

End Sub

