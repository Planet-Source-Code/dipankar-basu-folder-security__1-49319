Attribute VB_Name = "meSetting"
Option Explicit
Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias _
"RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName _
As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, _
lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey _
As Long) As Long
' Read & Write Access to Windows Settings INI files and read windows registry values
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName _
As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal _
lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName _
As String) As Long
Public Function ReadINI(sSection As String, sKeyName As String, sINIFileName _
As String) As String
On Local Error Resume Next: Dim sRet As String: sRet = String(255, Chr(0))
ReadINI = Left(sRet, GetPrivateProfileString(sSection, _
ByVal sKeyName, vbNullString, sRet, Len(sRet), sINIFileName))
End Function
Public Function WriteINI(sSection As String, sKeyName As String, sValueData _
As String, sINIFileName As String) As Boolean
'On Local Error Resume Next
Call WritePrivateProfileString(sSection, sKeyName, sValueData, _
sINIFileName)
WriteINI = (Err.Number = 0)
End Function
Public Function ReadString(ByVal hKey As HKeyTypes, ByVal strPath As String, ByVal strValue As String) As String
Dim r As Long, varType As Long, varData As String, varLen As Long, KeyHand As Long: On Local Error Resume Next
varLen = 1024: varData = String$(1024, 0)
r = RegOpenKey(hKey, strPath, KeyHand)
r = RegQueryValueEx(KeyHand, strValue, 0, varType, ByVal varData, varLen)
r = RegCloseKey(KeyHand)
ReadString = Trim(Left(varData, InStr(1, varData, Chr(0)) - 1))
End Function
' (c) 2002 All Rights Reserved by Dipankar Basu

