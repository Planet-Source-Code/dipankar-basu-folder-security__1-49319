VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
End
End Sub
Private Sub cmdOK_Click()
 Dim PassCode As String, sFilePath As String
        sFilePath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
        PassCode = ReadINI("Folder Security", "Login", sFilePath & "secure.ini")
        If txtUserName.Text = vbNullString Or txtPassword.Text = vbNullString Then GoTo LoginFailed:
        If PassCode = Encrypt(txtUserName.Text, txtPassword.Text) Then
        GoTo LoginSuccess:
        Else
        GoTo LoginFailed:
        End If
LoginSuccess:
        Me.Hide
        frmSecurity.Show
        Unload Me
        Exit Sub
LoginFailed:
Static flsCount As Integer
Call MsgBox("Incorrect Login Information", vbInformation, "Login Failed")
        flsCount = flsCount + 1
        If flsCount = 4 Then Call cmdCancel_Click
        txtPassword.SetFocus
End Sub
Private Sub Form_Activate()
 Dim PassCode As String, sFilePath As String
 sFilePath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
 PassCode = ReadINI("Folder Security", "Login", sFilePath & "secure.ini")
If PassCode = vbNullString Then _
frmSecurity.Show: Unload Me
End Sub
Private Sub Form_Load()
If App.PrevInstance Then End
End Sub
Private Sub txtPassword_GotFocus()
With txtPassword
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub
