VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Messenger Debug"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7830
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTopMenu 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7830
      TabIndex        =   1
      Top             =   0
      Width           =   7830
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Messenger Debug"
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
         Width           =   7835
      End
      Begin VB.Image imgMenuBG 
         Height          =   615
         Left            =   0
         Picture         =   "frmDebug.frx":058A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "name"
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   3960
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   630
      Width           =   7815
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Me.Hide
End Sub
