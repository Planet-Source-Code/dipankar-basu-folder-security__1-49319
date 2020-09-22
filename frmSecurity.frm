VERSION 5.00
Begin VB.Form frmSecurity 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Security"
   ClientHeight    =   3570
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox UnProtectLst 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1620
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   120
      X2              =   5520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H00400000&
      Height          =   3375
      Left            =   120
      Top             =   120
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      Height          =   2415
      Left            =   240
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label lblSecureDir 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Browse to Secure Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "click here to Browse folders in the system"
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label lblPLUpdt 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Protected Folders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Click here to Update the List of Protected folders in the system"
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSecure 
         Caption         =   "&Secure"
      End
      Begin VB.Menu mnuUnSecure 
         Caption         =   "&Unsecure"
      End
      Begin VB.Menu mnusepf1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEBin 
         Caption         =   "&Empty Recycle Bin"
      End
      Begin VB.Menu mnusepf2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuPassword 
         Caption         =   "Change Login &Password"
      End
      Begin VB.Menu mnuseph1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About . . ."
      End
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Program : Secure & Hide Folders
'   Written by : Dipankar Basu       http://www.geocities.com/basudip_in/
'   Revision : March 19, 2003.
'   (c)2003 All rights reserved by Dipankar Basu
Option Explicit
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function shEmptyRecycleBin Lib "SHELL32.DLL" Alias _
    "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal RootPath As String, _
     ByVal dwFlags As Long) As Long
Dim WindowsDirectory As String
Private Const Ext = ".{645FF040-5081-101B-9F08-00AA002F954E}" ' RecycleBin
Private Function bFileExists(sFileName As String) As Boolean
On Error Resume Next
Dim I As Integer
I = Len(Dir$(sFileName))
bFileExists = IIf(Err Or I = 0, False, True)
End Function
Private Sub lblPLUpdt_Click()
Dim Counter As Integer, MaxFls As Integer, sFilePath As String, itemData As String
sFilePath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
MaxFls = Val(ReadINI("Protected Folders List", "MaxFiles", sFilePath & "secure.ini"))
UnProtectLst.Clear
If MaxFls < 1 Then Exit Sub
For Counter = 1 To MaxFls
itemData = ReadINI("Protected Folders List", "File" & CStr(Counter), sFilePath & "secure.ini")
If itemData <> vbNullString Then UnProtectLst.AddItem itemData
DoEvents
Next
End Sub
Private Sub lblPLUpdt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPLUpdt.Caption = "Refresh List"
End Sub
Private Sub lblPLUpdt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPLUpdt.Caption = "Protected Folders"
End Sub
Private Sub lblSecureDir_Click()
Dim sFileName As String
sFileName = BrowseFolder(Me.hWnd, "Select the folder to Protect")
If StrPtr(sFileName) = 0 Or Trim(sFileName) = vbNullString Then Exit Sub
Call SecureFolder(sFileName)
End Sub
Private Sub mnuContents_Click()
On Error Resume Next: Dim sFilePath As String
sFilePath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
If bFileExists(sFilePath & "secure.chm") Then
App.HelpFile = sFilePath & "secure.chm" ' Help File
SendKeys "{F1}", True
Else
Call MsgBox("Help files (secure.chm) do not exist, Please install the help file", vbCritical, "Folder Security pre-release Beta-test")
End If
End Sub
Private Sub mnuEBin_Click()
shEmptyRecycleBin 0, vbNullString, False
End Sub
Private Sub mnuExit_Click()
Unload Me
End Sub
Private Sub mnuPassword_Click()
Password.Show
End Sub
Private Sub mnuSecure_Click()
On Error GoTo Err
Dim Path As String
reEnterPath:
Path = InputBox("Enter the name of the folder to Protect" & vbCrLf & _
"Provide the full path and the name of the folder / directory", "Protect Folder")
If StrPtr(Path) = 0 Or Trim(Path) = vbNullString Then Exit Sub
If Not IsFolderExists(Trim(Path)) Then
MsgBox "The folder do not exist in the specified path or the specified folder is protected" _
, vbInformation, "Folder is Inaccessible"
GoTo reEnterPath:
End If
Call SecureFolder(Trim(Path))
Exit Sub
Err:
MsgBox Err.Description, , Err.Number
End Sub
Private Sub Unsecure(ByVal fPath As String)
On Error GoTo Err
    Dim Temp As String, Data As String, File As String, FileName As String, Path As String
reEnterPath:
If fPath = vbNullString Then
    Path = InputBox("Enter the name of the folder to unhide" & vbCrLf & _
    "Provide the full path and the name of the folder / directory", "Unprotect Folder")
 Else
 Path = fPath
 fPath = vbNullString
 End If
    If StrPtr(Path) = 0 Or Trim(Path) = vbNullString Then Exit Sub
    If Not IsFolderExists(Trim(Trim(Path) & Ext)) Then
    MsgBox "The folder do not exist in the specified path or the specified folder is not protected" _
    , vbInformation, "Folder is Inaccessible"
    GoTo reEnterPath:
    End If
    Path = Trim(Path) & Ext
    Temp = Mid$(Path, InStrRev(Path, "\") + 1, Len(Path))
    Data = Left$(Temp, InStr(Temp, ".{") - 1)
    File = Left$(Path, Len(Path) - Len(Temp))
    FileName = File & Data
    Name Path As FileName
    ClearFolderHide (FileName)
    DoEvents
       Dim sFilePath As String, n As Integer
       sFilePath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
       Temp = ReadINI("Protected Folders List", "MaxFiles", sFilePath & "secure.ini")
readAgain:
        n = n + 1: Path = vbNullString
        Path = ReadINI("Protected Folders List", "File" & n, sFilePath & "secure.ini")
        If n > Val(Temp) Then GoTo NoEntryFound:
        If Trim(UCase(Path)) <> Trim(UCase(FileName)) Then GoTo readAgain:
        Call WriteINI("Protected Folders List", "File" & n, vbNullString, sFilePath & "secure.ini")
        If Val(Temp) = n Then
deleteAgain:
            n = n - 1
            Temp = ReadINI("Protected Folders List", "File" & n, sFilePath & "secure.ini")
            If Temp = vbNullString And n > 0 Then GoTo deleteAgain:
            Call WriteINI("Protected Folders List", "MaxFiles", CStr(n), sFilePath & "secure.ini")
         End If
NoEntryFound:
MsgBox "The folder is unprotected" & vbCrLf & Path, vbApplicationModal + vbInformation, "Security removed"
Exit Sub
Err:
MsgBox Err.Description, , Err.Number
End Sub
Private Sub Form_Load()
 On Local Error Resume Next
    Dim ret As Long, buff As String
    buff = Space(255)
    ret = GetWindowsDirectory(buff, 255)
    WindowsDirectory = Left$(buff, InStr(buff, vbNullChar) - 1)
Dim Counter As Integer, MaxFls As Integer, sFilePath As String, itemData As String
sFilePath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
MaxFls = Val(ReadINI("Protected Folders List", "MaxFiles", sFilePath & "secure.ini"))
If MaxFls < 1 Then Exit Sub
For Counter = 1 To MaxFls
itemData = ReadINI("Protected Folders List", "File" & CStr(Counter), sFilePath & "secure.ini")
If itemData <> vbNullString Then UnProtectLst.AddItem itemData
DoEvents
Next
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
Private Sub mnuAbout_Click()
frmAbout.Show vbModal, Me
End Sub
Private Sub mnuUnSecure_Click()
Call Unsecure(vbNullString)
End Sub
Private Sub UnProtectLst_DblClick()
Call Unsecure(UnProtectLst.Text)
End Sub
Private Sub UnProtectLst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With UnProtectLst
If .Text = vbNullString Then
.ToolTipText = "double-click to unprotect folder"
Else
.ToolTipText = .Text
End If
End With
End Sub
Private Sub SecureFolder(ByVal sFolderName As String)
On Error GoTo Err
    Dim Path As String, Data As String, File As String, FileName As String
    Path = sFolderName
    Data = Mid$(Path, InStrRev(Path, "\") + 1, Len(Path))
    File = Left$(Path, Len(Path) - Len(Data))
    If Not UCase$(Path) = UCase$(WindowsDirectory) _
    And Not UCase$(Data) = UCase$("desktop") Then
            If InStr(UCase(Path), UCase(Ext)) Then
            MsgBox "Unable to protect the specified folder" & vbCrLf & _
            "The folder may be protected", , "Folder Security"
            Exit Sub
            End If
        FileName = File & Data & Ext
        Name sFolderName As FileName
        SetFolderHide (FileName)
        DoEvents
        Dim sFilePath As String, n As Integer
        sFilePath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
readAgain:
        n = n + 1: Path = vbNullString
        Path = ReadINI("Protected Folders List", "File" & n, sFilePath & "secure.ini")
        If Trim(Path) <> vbNullString Or Null Then GoTo readAgain:
        If Val(n) > Val(ReadINI("Protected Folders List", "MaxFiles", sFilePath & "secure.ini")) Then _
        Call WriteINI("Protected Folders List", "MaxFiles", CStr(n), sFilePath & "secure.ini")
        Call WriteINI("Protected Folders List", "File" & n, File & Data, sFilePath & "secure.ini")
MsgBox "The folder is protected" & vbCrLf & File & Data, vbApplicationModal + vbInformation, "Security applied"
    Else
        MsgBox "This folder cannot be protected", vbApplicationModal + vbInformation, "Security..."
    End If
Exit Sub
Err:
If Err.Number = 75 Then
MsgBox "some files in the specified folder may be opened," & vbCrLf & _
    "close the associated program and try to protect this folder", vbInformation, "Unable to protect the folder"
Else
MsgBox Err.Description, , Err.Number
End If
End Sub
'       Reference :-
' Extending Windows Explorer with Name Space Extensions
