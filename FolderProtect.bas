Attribute VB_Name = "DirectorySecure"
' Written : October 02, 2002 by Dipankar Basu, Modified Build :  March 18, 2003.
' URL : http://www.geocities.com/basudip_in/
' Copyright  (c) 2002  All rights reserved by Dipankar Basu
Option Explicit
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Declare Function SHGetPathFromIDList Lib "SHELL32.DLL" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "SHELL32.DLL" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Public Function BrowseFolder(hWnd As Long, szDialogTitle As String) As String
'   Syntax:  StrVar = BrowseFolder(hWnd, StrVar)
On Local Error Resume Next
    Dim X As Long, BI As BROWSEINFO, dwIList As Long, szPath As String, wPos As Integer
    BI.hOwner = hWnd
    BI.lpszTitle = szDialogTitle
    BI.ulFlags = BIF_RETURNONLYFSDIRS
    dwIList = SHBrowseForFolder(BI)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    If X Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Trim(Left$(szPath, wPos - 1))
    Else
        BrowseFolder = vbNullString
    End If
End Function
Public Sub SetFolderHide(folderspec)
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    If Not (f.Attributes And 2) Then f.Attributes = f.Attributes + 2
End Sub
Public Sub ClearFolderHide(folderspec)
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    If f.Attributes And 2 Then f.Attributes = f.Attributes - 2
End Sub
Public Function IsFolderExists(FolderName As String) As Boolean
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
IsFolderExists = fs.FolderExists(FolderName)
End Function
Public Function Encrypt(ByVal strInput As String, ByVal strKey As String) As String
Dim iCount As Long, ingPtr As Long, strOutput As String
strInput = strInput & strKey
For iCount = 1 To Len(strInput)
strOutput = strOutput + Hex(Asc(Chr((Asc(Mid(strInput, iCount, 1))) Xor Asc(Mid(strKey, ingPtr + 1, 1)))) + iCount * Len(strKey) Xor Len(strInput))
ingPtr = ((ingPtr + 1) Mod (Len(strKey)))
Next iCount
Encrypt = LCase(strOutput)
End Function
