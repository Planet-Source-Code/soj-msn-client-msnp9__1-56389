VERSION 5.00
Begin VB.Form frmAddAccount 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Account"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3735
   Icon            =   "frmAddAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPass 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remember Password"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox txtDisplay 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Display:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   220
      UseMnemonic     =   0   'False
      Width           =   3735
   End
   Begin VB.Image imgMenuBG 
      Height          =   615
      Left            =   0
      Picture         =   "frmAddAccount.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3765
   End
End
Attribute VB_Name = "frmAddAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkPass_Click()
If chkPass.Value = 1 Then
    txtPassword.Enabled = True
Else
    txtPassword.Enabled = False
End If
End Sub

Private Sub cmdAdd_Click()
Dim AccountTable As New ADODB.Recordset
AccountTable.Open "SELECT * FROM Accounts WHERE aEmail='" & EncryptMe(txtEmail.Text) & "'", db, 3, 3

If txtEmail.Text = "" Then MsgBox "You must provide an Email.", vbExclamation: Exit Sub
If txtDisplay.Text = "" Then MsgBox "You must provide an Account Display.", vbExclamation: Exit Sub
If chkPass.Value = 1 And txtPassword = "" Then MsgBox "You must provide a Password.", vbExclamation: Exit Sub

If AccountTable.EOF Then
    AccountTable.AddNew
    AccountTable("aDisplayPic") = "%app_path%\Display Data\image.bmp"
    AccountTable("aEmail") = EncryptMe(txtEmail.Text)
    AccountTable("aName") = txtDisplay.Text
    If chkPass.Value = 1 Then AccountTable("aPassword") = EncryptMe(EncryptMe(txtPassword.Text))
    AccountTable.Update
Else
    MsgBox "That email already exists in the list.", vbExclamation
End If

AccountTable.Close
frmManageAccounts.LoadListUp
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub
