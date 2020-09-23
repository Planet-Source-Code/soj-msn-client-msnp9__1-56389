VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageAccounts 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Management"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4695
   Icon            =   "frmManageAccounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlNull 
      Left            =   3480
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageAccounts.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlAccounts 
      Left            =   4080
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstAccounts 
      Height          =   4680
      Left            =   0
      TabIndex        =   0
      Top             =   615
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8255
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "imlAccounts"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      OLEDragMode     =   1
      NumItems        =   0
      Picture         =   "frmManageAccounts.frx":0A62
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Account Management"
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
      TabIndex        =   4
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   4735
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      Height          =   195
      Left            =   4230
      TabIndex        =   3
      Top             =   5355
      Width           =   390
   End
   Begin VB.Image imgClose 
      Height          =   195
      Left            =   3960
      Picture         =   "frmManageAccounts.frx":2864
      Top             =   5355
      Width           =   195
   End
   Begin VB.Label lblRemove 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remove Account"
      Height          =   195
      Left            =   2070
      TabIndex        =   2
      Top             =   5355
      Width           =   1245
   End
   Begin VB.Image imgRemove 
      Height          =   195
      Left            =   1800
      Picture         =   "frmManageAccounts.frx":2AAE
      Top             =   5355
      Width           =   195
   End
   Begin VB.Label lblAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Account"
      Height          =   195
      Left            =   330
      TabIndex        =   1
      Top             =   5355
      Width           =   930
   End
   Begin VB.Image imgAdd 
      Height          =   195
      Left            =   45
      Picture         =   "frmManageAccounts.frx":2CF8
      Top             =   5355
      Width           =   195
   End
   Begin VB.Image imgBMBG 
      Height          =   300
      Left            =   0
      Picture         =   "frmManageAccounts.frx":2F42
      Stretch         =   -1  'True
      Top             =   5300
      Width           =   4905
   End
   Begin VB.Image imgMenuBG 
      Height          =   615
      Left            =   0
      Picture         =   "frmManageAccounts.frx":3344
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4725
   End
End
Attribute VB_Name = "frmManageAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastdoneitem


Private Sub Form_Load()
LoadListUp
End Sub

Public Function LoadListUp()
Dim AccountTable As New ADODB.Recordset
AccountTable.Open "SELECT * FROM Accounts", db, adOpenDynamic, adLockOptimistic

doneone = False

lstAccounts.ListItems.Clear
lstAccounts.Icons = imlNull
imlAccounts.ListImages.Clear

Do Until AccountTable.EOF
    If doneone = False Then
        lastdoneitem = DecryptMe(AccountTable("aEmail"))
        doneone = True
    End If
    thedisplaypic = Replace(AccountTable("aDisplayPic"), "%app_path%", App.Path)
    imlAccounts.ListImages.Add , DecryptMe(AccountTable("aEmail")), LoadPicture(thedisplaypic)
    lstAccounts.Icons = imlAccounts
    lstAccounts.ListItems.Add , DecryptMe(AccountTable("aEmail")), AccountTable("aName"), DecryptMe(AccountTable("aEmail"))
    AccountTable.MoveNext
Loop

AccountTable.Close
frmSignIn.LoadAccounts
End Function

Private Sub imgAdd_Click()
lblAdd_Click
End Sub

Private Sub lblAdd_Click()
frmAddAccount.lblTitle.Caption = "Add Account"
frmAddAccount.Caption = "Add Account"
frmAddAccount.Show vbModal
End Sub

Private Sub imgRemove_Click()
lblRemove_Click
End Sub

Private Sub lblRemove_Click()
Dim AccountTable As New ADODB.Recordset
AccountTable.Open "SELECT * FROM Accounts WHERE aEmail='" & EncryptMe(lastdoneitem) & "'", db, adOpenDynamic, adLockOptimistic

theresult = MsgBox("Are you sure you want to delete this account:" & vbCrLf & lastdoneitem, vbExclamation + vbYesNo)

If theresult = vbYes Then
    AccountTable.Delete
    LoadListUp
End If

AccountTable.Close
End Sub

Private Sub imgClose_Click()
lblClose_Click
End Sub

Private Sub lblClose_Click()
Unload Me
End Sub


Private Sub lstAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)
lastdoneitem = Item.Key
End Sub

Private Sub lstAccounts_DblClick()
Dim AccountTable As New ADODB.Recordset
AccountTable.Open "SELECT * FROM Accounts WHERE aEmail='" & EncryptMe(lastdoneitem) & "'", db, adOpenDynamic, adLockOptimistic

frmAddAccount.lblTitle.Caption = "Edit Account"
frmAddAccount.Caption = "Edit Account"

frmAddAccount.txtDisplay.Text = AccountTable("aName")
frmAddAccount.txtEmail.Text = DecryptMe(AccountTable("aEmail"))
If Not AccountTable("aPassword") = "" Then
    frmAddAccount.txtPassword.Text = DecryptMe(DecryptMe(AccountTable("aPassword")))
    frmAddAccount.txtPassword.Enabled = True
    frmAddAccount.chkPass.Value = 1
Else
    frmAddAccount.txtPassword.Text = ""
    frmAddAccount.txtPassword.Enabled = False
    frmAddAccount.chkPass.Value = 0
End If

AccountTable.Close
frmAddAccount.Show vbModal
End Sub
