VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvite 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invite Someone to this Conversation"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5775
   Icon            =   "frmInvite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "OK"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin MSComctlLib.TreeView trvContacts 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6588
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invite Someone to this Conversation"
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
      Top             =   225
      UseMnemonic     =   0   'False
      Width           =   5775
   End
   Begin VB.Image imgMenuBG 
      Height          =   615
      Left            =   0
      Picture         =   "frmInvite.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "frmInvite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOkay_Click()
Unload Me
End Sub

Private Sub Form_Load()
trvContacts.ImageList = frmMain.imgList
trvContacts.Nodes.Clear

'''''''''''''''''''''''''''''''
'Lists the WHOLE contact list.'
'  Dim itm As Variant
'  For Each itm In frmMain.contacts
'    trvContacts.Nodes.Add , , itm, MSNDecode(frmMain.contacts(itm)), 4
'  Next
'''''''''''''''''''''''''''''''

  Dim itm As Variant
  For Each itm In frmMain.online_contacts
    trvContacts.Nodes.Add , , itm, MSNDecode(frmMain.online_contacts(itm)), 4
  Next
End Sub
