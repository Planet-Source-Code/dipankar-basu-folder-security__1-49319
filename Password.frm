VERSION 5.00
Begin VB.Form Password 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Login Password"
   ClientHeight    =   3075
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox n2Pass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox n1Pass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox uPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox uName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      MaxLength       =   25
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Retype Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "New Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CancelButton_Click()
Unload Me
End Sub
Private Sub n1Pass_GotFocus()
With n1Pass
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub
Private Sub n2Pass_GotFocus()
With n2Pass
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub
Private Sub OKButton_Click()
' The password do not care for the CAPS-Lock key
' that is if a user register a password like :-
' User : user       Password : login
' the user will be also able to login as :-
' User : USER       Password : LOGIN
If Trim(uName.Text) = vbNullString Then
uName.SetFocus
Call MsgBox("Please input User's Name", vbInformation, "Login Information")
Exit Sub
ElseIf Len(uName.Text) < 4 Then
uName.SetFocus
Call MsgBox("User's Name should be atleast 4 chars in length", vbInformation, "Login Information")
Exit Sub
ElseIf Len(n1Pass.Text) > 25 Then
Call MsgBox("Password should be maximum 25 chars in length", vbInformation, "Login Information")
Exit Sub
End If
Dim PassCode As String, sFilePath As String
sFilePath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
PassCode = ReadINI("Folder Security", "Login", sFilePath & "secure.ini")
If PassCode = vbNullString Then GoTo noVerify:
If uName.Text = vbNullString Or uPassword.Text = vbNullString Then
Call MsgBox("Please enter User's Name & Password", vbInformation, "Login Information")
Exit Sub
End If
If Encrypt(uName.Text, uPassword.Text) = PassCode Then
noVerify:
    If n1Pass.Text = n2Pass.Text Then
    If n1Pass.Text = vbNullString Then
    PassCode = vbNullString
    Else
    PassCode = Encrypt(uName.Text, n1Pass.Text)
    End If
    If WriteINI("Folder Security", "Login", PassCode, sFilePath & "secure.ini") Then
    If PassCode = vbNullString Then
    Call MsgBox("Password is successfully deleted", vbInformation, "Login Password Cleared")
    Else
    Call MsgBox("Password is successfully changed", vbInformation, "Login Password Changed")
    End If
    Call CancelButton_Click
    End If
    Else
    Call MsgBox("New Password do not match", vbInformation, "Password mismatch")
    n1Pass.SetFocus
    End If
Else
Call MsgBox("Username & Password Mismatch", vbInformation, "Incorrect Login Information")
uPassword.SetFocus
End If
End Sub
Private Sub uPassword_GotFocus()
With uPassword
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub
' (c) 2002 All Rights Reserved by Dipankar Basu
