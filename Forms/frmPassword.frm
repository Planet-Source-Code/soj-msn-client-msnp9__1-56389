VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3735
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
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
      TabIndex        =   3
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   3735
   End
   Begin VB.Image imgMenuBG 
      Height          =   615
      Left            =   0
      Picture         =   "frmPassword.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3765
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmMain.txtPassword.Text = ""
Unload Me
End Sub

Private Sub cmdOkay_Click()
frmMain.txtPassword.Text = txtPassword.Text
Unload Me
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOkay_Click
End Sub
