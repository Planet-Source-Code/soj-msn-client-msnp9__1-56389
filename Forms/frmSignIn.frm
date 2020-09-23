VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSignIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sign In"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4350
   Icon            =   "frmSignIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSignIn.frx":058A
   ScaleHeight     =   2565
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlAccounts 
      Left            =   3720
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   50
      ImageHeight     =   50
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstAccounts 
      Height          =   1305
      Left            =   600
      TabIndex        =   1
      Top             =   510
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2302
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
      Picture         =   "frmSignIn.frx":24C44
   End
   Begin MSComctlLib.ImageList imlNull 
      Left            =   3120
      Top             =   1920
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
            Picture         =   "frmSignIn.frx":26A46
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblManage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Accounts"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   2110
      Width           =   1335
   End
End
Attribute VB_Name = "frmSignIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastdoneitem
Dim LoggedIn As Boolean

Private Sub Form_Load()
LoadDB
LoadAccounts
End Sub

Private Sub Form_Unload(Cancel As Integer)
If LoggedIn = False Then End
End Sub

Private Sub lblManage_Click()
frmManageAccounts.Show vbModal
End Sub

Public Function LoadAccounts()
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
    lstAccounts.ListItems(DecryptMe(AccountTable("aEmail"))).Tag = AccountTable("aPassword") & "«:»" & thedisplaypic
    AccountTable.MoveNext
Loop

LoggedIn = False

AccountTable.Close
End Function

Private Sub lstAccounts_DblClick()
tagdata = Split(lstAccounts.ListItems(lastdoneitem).Tag, "«:»")
frmMain.txtSigninName.Text = lastdoneitem
If tagdata(0) = "" Then
    frmPassword.Show vbModal
Else
    frmMain.txtPassword.Text = DecryptMe(DecryptMe(tagdata(0)))
End If
dplocation = tagdata(1)
If Not frmMain.txtPassword.Text = "" Then LoggedIn = True: frmMain.Show: frmMain.SignInMSN: Unload Me
End Sub

Private Sub lstAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)
lastdoneitem = Item.Key
End Sub
