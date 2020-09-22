VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   3
         ToolTipText     =   "Close"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblAuthor 
         BackStyle       =   0  'Transparent
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   1365
         Width           =   4335
      End
      Begin VB.Label lblmail 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   1875
         Width           =   4335
      End
      Begin VB.Label lblURL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Web:URL"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label lblSystem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "System :  "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label lblcopy 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (c) All Rights Reserved"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3000
         Width           =   4335
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   120
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label lblAppNameVer 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "App Name and Version"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const AppName = "Folder Security"
Private Sub cmdOK_Click()
Unload Me
End Sub
Private Sub Form_Load()
Dim OSystem As String
OSystem = ReadString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductName")
If Trim(OSystem) = vbNullString Then _
OSystem = ReadString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "ProductName")
lblSystem.Caption = "System :  " & OSystem
lblURL.Caption = "http://in.geocities.com/dipankarbasu_2k1/" 'url:web
lblmail.Caption = "basudip_in@hotmail.com" 'mailto:email
lblAuthor.Caption = AppName & " is developed by " & "Dipankar Basu"
lblcopy.Caption = "Copyright (c)2003  BasuDip Software Emporium"
lblAppNameVer.Caption = AppName & " Version :  " & App.Major & "." & App.Minor & "  Build :  " & App.Revision
Me.Caption = "About " & AppName
End Sub
Private Sub lblcopy_DblClick()
Warranty
End Sub
Private Sub lblmail_Click()
ShellExecute 0, "Open", "mailto:" & lblmail.Caption & "?Subject=" & App.Title, _
    vbNullString, vbNullString, vbNormal
End Sub
Private Sub lblmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmail.ForeColor = vbRed
End Sub
Private Sub lblmail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmail.ForeColor = vbBlue
End Sub
Private Sub lblURL_Click()
ShellExecute 0, "Open", lblURL.Caption, vbNullString, vbNullString, vbNormal
End Sub
Private Sub lblURL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblURL.ForeColor = vbRed
End Sub
Private Sub lblURL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblURL.ForeColor = vbBlue
End Sub
Private Sub Warranty()
Dim MsgW As String
MsgW = "This software is AS IS without warranty of any kind."
MsgW = MsgW + " While every possible care is taken, to ensure that the software is efficient and bug free."
MsgW = MsgW + " The developer does not hold himself responsible for any damage or data loss as a result of using"
MsgW = MsgW + " or distributing this software. In no event will Dipankar Basu be liable for any damages, however"
MsgW = MsgW + " caused and regardless of the theory of liability, arising out of the use of or inability to use the software."
Call MsgBox(MsgW, vbInformation + vbOKOnly, AppName)
End Sub
' (c) 2002 All Rights Reserved by Dipankar Basu
