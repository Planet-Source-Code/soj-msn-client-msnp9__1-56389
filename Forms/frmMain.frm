VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Soj Messenger"
   ClientHeight    =   6810
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   5040
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":058A
   ScaleHeight     =   6810
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   3120
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4320
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2422
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3274
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":373A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBottomMenu 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5040
      TabIndex        =   5
      Top             =   6510
      Width           =   5040
      Begin VB.Label lblAddContact 
         BackStyle       =   0  'Transparent
         Caption         =   "Add a Contact"
         Height          =   255
         Left            =   315
         TabIndex        =   6
         Top             =   60
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   195
         Left            =   60
         Picture         =   "frmMain.frx":3C00
         Top             =   60
         Width           =   195
      End
      Begin VB.Image imgBMBG 
         Height          =   300
         Left            =   0
         Picture         =   "frmMain.frx":3E4A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5145
      End
   End
   Begin VB.PictureBox picTopMenu 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5040
      TabIndex        =   3
      Top             =   0
      Width           =   5040
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Online)"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   540
      End
      Begin VB.Image imgDude 
         Height          =   615
         Left            =   0
         Picture         =   "frmMain.frx":424C
         Top             =   0
         Width           =   495
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   465
      End
      Begin VB.Image imgMenuBG 
         Height          =   615
         Left            =   0
         Picture         =   "frmMain.frx":5292
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5085
      End
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtSigninName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSComctlLib.TreeView trvContacts 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6588
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   1
      ImageList       =   "imgList"
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
   Begin VB.Image imgAway 
      Height          =   615
      Left            =   2040
      Picture         =   "frmMain.frx":56AC
      Top             =   4560
      Width           =   495
   End
   Begin VB.Image imgBusy 
      Height          =   615
      Left            =   1440
      Picture         =   "frmMain.frx":66F2
      Top             =   4560
      Width           =   495
   End
   Begin VB.Image imgOnline 
      Height          =   615
      Left            =   240
      Picture         =   "frmMain.frx":7738
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgOffline 
      Height          =   615
      Left            =   840
      Picture         =   "frmMain.frx":877E
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileSignOut 
         Caption         =   "Sign Out"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMyStatus 
         Caption         =   "My Status"
         Begin VB.Menu mnuFileMyStatusOnline 
            Caption         =   "Online"
         End
         Begin VB.Menu mnuFileMyStatusBusy 
            Caption         =   "Busy"
         End
         Begin VB.Menu mnuFileMyStatusBRB 
            Caption         =   "Be Right Back"
         End
         Begin VB.Menu mnuFileMyStatusAway 
            Caption         =   "Away"
         End
         Begin VB.Menu mnuFileMyStatusOnThePhone 
            Caption         =   "On The Phone"
         End
         Begin VB.Menu mnuFileMyStatusOutToLunch 
            Caption         =   "Out to Lunch"
         End
         Begin VB.Menu mnuFileMyStatusAppearOffline 
            Caption         =   "Appear Offline"
         End
      End
   End
   Begin VB.Menu mnuContacts 
      Caption         =   "Contacts"
      Begin VB.Menu mnuContactsAddContact 
         Caption         =   "Add Contact..."
      End
      Begin VB.Menu mnuContactsSearchList 
         Caption         =   "Search Contact List"
      End
   End
   Begin VB.Menu mbuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpShowDebug 
         Caption         =   "Show Debug"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents oWinHTTP As WinHttp.WinHttpRequest
Attribute oWinHTTP.VB_VarHelpID = -1

Private bGotAuthRedir As Boolean
Private sChallenge As String
Private sAuthLocation As String
Private dTransID As Double
Private lLsts As Long
Public lLstCount As Long

Public ClientIDNo As String

Public group_names As New Scripting.Dictionary
Public group_no As New Scripting.Dictionary
Public contact_groups As New Scripting.Dictionary
Public contacts As New Scripting.Dictionary
Public online_contacts As New Scripting.Dictionary

Private Function TransactionID() As Double

    If dTransID = 2 ^ 32 - 1 Then
        dTransID = 1
    End If
    
    TransactionID = dTransID
    dTransID = dTransID + 1
    
End Function


Private Sub Form_Resize()
On Error Resume Next
trvContacts.Width = Me.Width - 240
imgMenuBG.Width = Me.Width + 10
imgBMBG.Width = Me.Width + 10
picBottomMenu.Width = Me.Width + 10
trvContacts.Height = Me.Height - imgMenuBG.Height - 280 - imgBMBG.Height - 550
picBottomMenu.Top = trvContacts.Top + trvContacts.Height
End Sub

Private Sub imgDude_Click()
    lblStatus_Click
End Sub

Private Sub lblName_Change()
    lblStatus.Left = lblName.Left + lblName.Width + 70
End Sub

Private Sub lblName_Click()
    sName = InputBox("Please Enter Your New FriendlyName", "Change Your FriendlyName", lblName.Caption)
    If sName = "" Then Exit Sub
    sName = URL_Encode(MSNEncode(sName))
    Call SendData("REA *%* " & txtSigninName.Text & " " & sName & vbCrLf)
End Sub

Private Sub lblStatus_Click()
    PopupMenu mnuFileMyStatus
End Sub

Private Sub mnuContactsAddContact_Click()
    'sAdd = InputBox("Enter the email address you want to add to your contacts.", "Contacts Email Address", "@hotmail.com")
    'If sAdd = "" Or sAdd = "@hotmail.com" Then Exit Sub
    'Call SendData("ADD *%* FL " & sAdd & " " & sAdd & " 1")
    'Call SendData("ADD *%* AL " & sAdd & " " & sAdd & " 1")
    'trvContacts.Nodes.Add "Offline", tvwChild, sAdd, sAdd, 5
    ClientIDNo = "1"
End Sub

Private Sub mnuContactsSearchList_Click()
    frmSearchList.Show
End Sub

Private Sub mnuFileMyStatusAppearOffline_Click()
    Call CreateMSNObject
    Call SendData("CHG *%* HDN " & ClientIDNo & " " & URL_Encode(MSNObject) & vbCrLf)
End Sub

Private Sub mnuFileMyStatusAway_Click()
    Call CreateMSNObject
    Call SendData("CHG *%* AWY " & ClientIDNo & " " & URL_Encode(MSNObject) & vbCrLf)
End Sub

Private Sub mnuFileMyStatusBRB_Click()
    Call CreateMSNObject
    Call SendData("CHG *%* BRB " & ClientIDNo & " " & URL_Encode(MSNObject) & vbCrLf)
End Sub

Private Sub mnuFileMyStatusBusy_Click()
    Call CreateMSNObject
    Call SendData("CHG *%* BSY " & ClientIDNo & " " & URL_Encode(MSNObject) & vbCrLf)
End Sub

Private Sub mnuFileMyStatusOnline_Click()
    Call CreateMSNObject
    Call SendData("CHG *%* NLN " & ClientIDNo & " " & URL_Encode(MSNObject) & vbCrLf)
End Sub

Private Sub mnuFileMyStatusOnThePhone_Click()
    Call CreateMSNObject
    Call SendData("CHG *%* PHN " & ClientIDNo & " " & URL_Encode(MSNObject) & vbCrLf)
End Sub

Private Sub mnuFileMyStatusOutToLunch_Click()
    Call CreateMSNObject
    Call SendData("CHG *%* LUN " & ClientIDNo & " " & URL_Encode(MSNObject) & vbCrLf)
End Sub

Private Sub mnuFileSignOut_Click()
Call SendData("OUT" & vbCrLf)
Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuHelpShowDebug_Click()
frmDebug.Show
End Sub

Private Sub trvContacts_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Image = 2
End Sub


Private Sub trvContacts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then MsgBox trvContacts.SelectedItem.Key
End Sub


Private Sub trvContacts_DblClick()
    
    If trvContacts.SelectedItem.Image = 1 Or trvContacts.SelectedItem.Image = 2 Then Exit Sub
    If trvContacts.SelectedItem.Image = 5 Then
        'Show Profile Here / Send Email
        Exit Sub
    End If
    
    For Each frm In Forms
        If InStr(1, frm.Caption, trvContacts.SelectedItem.Text & " - <" & trvContacts.SelectedItem.Key & ">") Then
        frm.WindowState = vbNormal
        'frm.SetFocus
        Exit Sub
        End If
    Next frm

    SendData "XFR *%* SB" & vbCrLf
    If Left(trvContacts.SelectedItem.Key, 3) = "xt_" Then
        BuddyConnect = Right(trvContacts.SelectedItem.Key, Len(trvContacts.SelectedItem.Key) - 3)
    Else
        BuddyConnect = trvContacts.SelectedItem.Key
    End If
    BuddyConName = trvContacts.SelectedItem.Text
End Sub

Private Sub trvContacts_Expand(ByVal Node As MSComctlLib.Node)
    Node.Sorted = True
    Node.Image = 1
End Sub

Public Function SignInMSN()

    Set oWinHTTP = New WinHttp.WinHttpRequest
    lLsts = 0
    lLstCount = 0
    
    trvContacts.Nodes.Clear
    contacts.RemoveAll
    contact_groups.RemoveAll
    group_no.RemoveAll
    group_names.RemoveAll
    online_contacts.RemoveAll
    
    lblName.Caption = vbNullString
    lblStatus.Caption = vbNullString
    
    sckMain.Close
    sckMain.Connect "messenger.hotmail.com", 1863
    
End Function

Private Sub Form_Load()
    lLsts = 0
    'ClientIDNo = "1" 'Mobile Device
    'ClientIDNo = "536870932"
    ClientIDNo = "805306412"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub sckMain_Connect()
    Call SendData("VER *%* MSNP9" & vbCrLf)
End Sub

Public Sub SendData(Data As String)
    Data = Replace(Data, "*%*", TransactionID)
    
    Call sckMain.SendData(Data)
    
    'Output the data, trimming the newline if it has one.
    If Mid(Data, Len(Data) - 1) = vbCrLf Then
        Debug.Print ("<<<: " & Mid(Data, 1, Len(Data) - 2))
    Else
        Debug.Print ("<<<: " & Data)
    End If
    
    frmDebug.Text1.Text = frmDebug.Text1.Text & vbCrLf & vbCrLf & "******************" & vbCrLf & vbCrLf & Data
    frmDebug.Text1.SelLength = 0
    frmDebug.Text1.SelStart = Len(frmDebug.Text1.Text)
End Sub

Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
'On Error Resume Next
Dim sCommand As String, sData As String, sbuffer As String, sParams() As String
Dim i As Long

    Do
        Call sckMain.PeekData(sbuffer, vbString, bytesTotal)
        
        'Check theres a full command
        If InStr(1, sbuffer, vbCrLf) = 0 Then
            Exit Sub
        End If
        
        'Put the command into its own var
        If InStr(1, sbuffer, " ") Then
            sCommand = Mid$(sbuffer, 1, InStr(1, sbuffer, " ") - 1)
        Else
            sCommand = Mid$(sbuffer, 1, InStr(1, sbuffer, vbCrLf) - 1)
        End If

        'If the command contains payload data
        If sCommand = "MSG" Or sCommand = "NOT" Then
            i = InStr(1, sbuffer, vbCrLf)
            sParams() = Split(Mid$(sbuffer, 1, i - 1), " ")
            
            If CLng(Len(Mid$(sbuffer, i + 2))) < CLng(sParams(3)) Then
                'The message will be sent across multiple packets
                Exit Do
            End If
            
            sData = Mid(GetData(vbNullString, False, i + sParams(3) + 1), i + 2)

        Else
            sData = GetData(sbuffer)
            sParams() = Split(sData, " ")
            sData = vbNullString
            
        End If

        Call ProcessData(sParams, sData)

    Loop While sckMain.BytesReceived <> 0

End Sub

Private Sub ProcessData(sParams() As String, sPayload As String)
Dim sSubParams() As String
Dim sChat As New frmIM

    Select Case sParams(0)
        Case "VER"
            Call SendData("CVR 2 0x0409 winnt 5.1 i386 MSNMSGR 6.0.0254 MSMSGS " & txtSigninName.Text & vbCrLf)
            
        Case "CVR"
            Call SendData("USR 3 TWN I " & txtSigninName.Text & vbCrLf)
         
        Case "XFR"
            
            ' Redirect from server
            If sParams(2) = "NS" Then
                sSubParams() = Split(sParams(3), ":") 'Sub paramaters if needed
                Call sckMain.Close
                Call sckMain.Connect(sSubParams(0), sSubParams(1)) 'Connect to our refered server
            Else
            
            sSubParams() = Split(sParams(3), ":")
            
                For Each frm In Forms
                    If InStr(1, frm.Caption, "<" & BuddyConnect & ">") Then
                        With frm
                            .sckMSG.Tag = "USR *%* " & txtSigninName.Text & " " & sParams(5)
                            .sckMSG.Close
                            .sckMSG.Connect sSubParams(0), sSubParams(1)
                            '.lblCon.Caption = BuddyConnect
                            .WindowState = vbNormal
                            .Caption = BuddyConName & " - <" & BuddyConnect & ">"
                            .Visible = True
                            .BudEmail = BuddyConnect
                            .BudName = BuddyConName
                            '.lstParticipants.AddItem BuddyConnect
                            Exit Sub
                        End With
                    End If
                Next frm
                            
                'If not make a new chat window
                Load sChat
                With sChat
                    .sckMSG.Tag = "USR *%* " & txtSigninName.Text & " " & sParams(5)
                    .sckMSG.Close
                    .sckMSG.Connect sSubParams(0), sSubParams(1)
                    '.lblCon.Caption = BuddyConnect
                    .WindowState = vbNormal
                    .Caption = BuddyConName & " - <" & BuddyConnect & ">"
                    .Visible = True
                    .BudEmail = BuddyConnect
                    .BudName = BuddyConName
                    .lstParticipants.AddItem BuddyConnect
                End With
                
            End If

        Case "USR"
            If sParams(2) = "OK" Then
                lblName.Caption = MSNDecode(sParams(4))
                Call SendData("SYN *%* 0" & vbCrLf)
                
            ElseIf sParams(2) = "TWN" Then
                sChallenge = sParams(4)
                oWinHTTP.Option(WinHttpRequestOption_EnableRedirects) = False
                
                If Not bGotAuthRedir Then
                    Call oWinHTTP.Open("GET", "https://nexus.passport.com/rdr/pprdr.asp", True)
                    Call oWinHTTP.Send
                    
                Else
                    Call oWinHTTP.Open("GET", sAuthLocation, True)
                    Call oWinHTTP.SetRequestHeader("Authorization", "Passport1.4 OrgVerb=GET,OrgURL=http%3A%2F%2Fmessenger%2Emsn%2Ecom,sign-in=" & Replace(txtSigninName.Text, "@", "%40") & ",pwd=" & txtPassword.Text & "," & sParams(4))
                    Call oWinHTTP.Send
                    
                End If
               
            End If
            
        Case "SYN"
            lLstCount = sParams(3)
            
        Case "CHL"
            Call SendData("QRY *%* PROD0038W!61ZTF9 32" & vbCrLf _
            & MD5(sParams(2) & "VT6PX?UQTM4WM%YR"))

        Case "LSG"
            trvContacts.Nodes.Add , tvwParent, "list_" & sParams(1), URL_Decode(sParams(2)), 1
            trvContacts.Nodes("list_" & sParams(1)).Expanded = True
            trvContacts.Nodes("list_" & sParams(1)).Bold = True
            group_no.Add "list_" & sParams(1), 0
            group_names.Add "list_" & sParams(1), URL_Decode(sParams(2))
            
        Case "LST"
            lLsts = lLsts + 1
            
            If lLsts = 1 Then
                MakeExtendedList
                
                trvContacts.Nodes.Add , tvwParent, "Offline", "Offline", 1
                trvContacts.Nodes("Offline").Expanded = True
                trvContacts.Nodes("Offline").Bold = True
                group_no.Add "Offline", 0
            End If
            
            thegroup = "lst_0"
            If UBound(sParams) = 4 Then
                thelistor = Split(sParams(4), ",")
                thegroup = "list_" & thelistor(0)
            End If
            contact_groups.Add sParams(1), thegroup
                   
            If sParams(3) And 1 Then 'User is in contact list
                group_no("Offline") = group_no("Offline") + 1
                trvContacts.Nodes.Add "Offline", tvwChild, sParams(1), MSNDecode(sParams(2)), 5
                contacts.Add sParams(1), MSNDecode(sParams(2))
            ElseIf sParams(3) And 2 Then 'Allow list
            
            ElseIf sParams(3) And 4 Then 'Block List
            
            ElseIf sParams(3) And 8 Then 'Reverse list
                elperson = MSNDecode(sParams(2))
                elemailo = sParams(1)
                MsgBox elperson & " [" & elemailo & "] has added you to his/her contact list."
            End If
            
            If lLsts = lLstCount Then
                mnuFileMyStatusOnline_Click
                lblStatus.Caption = "(Online)"
            End If
            
        Case "ILN"
            thekey = trvContacts.Nodes(sParams(3)).Key
            thename = MSNDecode(sParams(4))
            contacts(thekey) = thename
            thetext = thename & WhatStatusName(sParams(2))
            trvContacts.Nodes.Remove sParams(3)
            
            group_no(contact_groups(sParams(3))) = group_no(contact_groups(sParams(3))) + 1
            trvContacts.Nodes(contact_groups(sParams(3))).Text = group_names(contact_groups(sParams(3))) & " (" & group_no(contact_groups(sParams(3))) & ")"
        
            trvContacts.Nodes.Add contact_groups(sParams(3)), tvwChild, thekey, thetext, WhatStatusNo(sParams(2))
            online_contacts(thekey) = contacts(thekey)
            
            trvContacts.Nodes("Offline").Sorted = True
        
        Case "NLN"
                
            If trvContacts.Nodes(sParams(2)).Image = 5 Then MsgBox trvContacts.Nodes(sParams(2)).Text & " has just signed in."
            
            thekey = trvContacts.Nodes(sParams(2)).Key
            thename = MSNDecode(sParams(3))
            thetext = thename & WhatStatusName(sParams(1))
            contacts(thekey) = thename
            trvContacts.Nodes.Remove sParams(2)
            
            'If sParams(1) = "NLN" Then group_no(contact_groups(sParams(2))) = group_no(contact_groups(sParams(2))) + 1
            'trvContacts.Nodes(contact_groups(sParams(2))).Text = group_names(contact_groups(sParams(2))) & " (" & group_no(contact_groups(sParams(2))) & ")"
        
            trvContacts.Nodes.Add contact_groups(sParams(2)), tvwChild, thekey, thetext, WhatStatusNo(sParams(1))
            online_contacts(thekey) = contacts(thekey)
    
            trvContacts.Nodes("Offline").Sorted = True
            
        Case "FLN"
            thekey = trvContacts.Nodes(sParams(1)).Key
            thetext = FixedUp(trvContacts.Nodes(sParams(1)).Text)
            
            group_no(contact_groups(thekey)) = group_no(contact_groups(thekey)) - 1
            If Not group_no(contact_groups(thekey)) = 0 Then thextra = " (" & group_no(contact_groups(thekey)) & ")"
            trvContacts.Nodes(contact_groups(thekey)).Text = group_names(contact_groups(thekey)) & thextra
            
            trvContacts.Nodes.Remove sParams(1)
            
            group_no("Offline") = group_no("Offline") + 1
            If Not group_no("Offline") = 0 Then theext = " (" & group_no("Offline") & ")"
            trvContacts.Nodes("Offline").Text = "Offline" & theext
            
            MsgBox thetext & " is now offline."
        
            trvContacts.Nodes.Add "Offline", tvwChild, thekey, thetext, 5
            online_contacts.Remove thekey
            
        Case "REA"
            lblName.Caption = MSNDecode(sParams(4))
            
        Case "CHG"
            Select Case sParams(2)
                Case "HDN"
                    lblStatus.Caption = "(Appear Offline)"
                    imgDude.Picture = imgOffline.Picture
                Case "AWY"
                    lblStatus.Caption = "(Away)"
                    imgDude.Picture = imgAway.Picture
                Case "BRB"
                    lblStatus.Caption = "(Be Right Back)"
                    imgDude.Picture = imgAway.Picture
                Case "BSY"
                    lblStatus.Caption = "(Busy)"
                    imgDude.Picture = imgBusy.Picture
                Case "NLN"
                    lblStatus.Caption = "(Online)"
                    imgDude.Picture = imgOnline.Picture
                Case "PHN"
                    lblStatus.Caption = "(On The Phone)"
                    imgDude.Picture = imgBusy.Picture
                Case "LUN"
                    lblStatus.Caption = "(Out To Lunch)"
                    imgDude.Picture = imgAway.Picture
            End Select

        Case "RNG"
            
            sNewS = Split(sParams(2), ":")
                    
           'Look for window that is already open
            For Each frm In Forms
                If InStr(1, frm.Caption, "<" & BuddyConnect & ">") Then
                    With frm
                        .sckMSG.Tag = "ANS *%* " & txtSigninName.Text & " " & sParams(4) & " " & sParams(1)
                        .sckMSG.Close
                        .sckMSG.Connect sNewS(0), sNewS(1)
                        '.lblCon.Caption = sParams(5)
                        .WindowState = vbMinimized
                        .Caption = URL_Decode(MSNDecode(sParams(6))) & " - <" & sParams(5) & ">"
                        .BudEmail = BuddyConnect
                        .BudName = BuddyConName
                        '.lstParticipants.AddItem BuddyConnect
                        Exit Sub
                    End With
                End If
            Next frm
                    
            'Make a new chat window
            Load sChat
            With sChat
                .sckMSG.Tag = "ANS *%* " & txtSigninName.Text & " " & sParams(4) & " " & sParams(1)
                .sckMSG.Close
                .sckMSG.Connect sNewS(0), sNewS(1)
                .Caption = URL_Decode(MSNDecode(sParams(6))) & " - <" & sParams(5) & ">"
                '.lblCon.Caption = sParams(5)
                .BudName = URL_Decode(MSNDecode(sParams(6)))
                .BudEmail = sParams(5)
                .lstParticipants.AddItem sParams(5)
            End With
            
        Case "OUT"
            If sParams(1) = "OTH" Then MsgBox "You have been signed out because you signed in from another location."
        
    End Select

End Sub

Function MakeExtendedList()
Dim ExtendedTable As New ADODB.Recordset
ExtendedTable.Open "SELECT * FROM Extended WHERE eOwner='" & EncryptMe(txtSigninName.Text) & "'", db, adOpenDynamic, adLockOptimistic

trvContacts.Nodes.Add , tvwParent, "l_extended", "Extended Contacts", 1
trvContacts.Nodes("l_extended").Expanded = True
trvContacts.Nodes("l_extended").Bold = True
group_no.Add "l_extended", 0

If Not ExtendedTable.EOF Then

    Do Until ExtendedTable.EOF
        trvContacts.Nodes.Add "l_extended", tvwChild, "xt_" & DecryptMe(ExtendedTable("eEmail")), DecryptMe(ExtendedTable("eDisplay")), 4
        ExtendedTable.MoveNext
    Loop
    
End If

ExtendedTable.Close
End Function

Function FixedUp(ByVal thename) As String

If Right(thename, 9) = " (Online)" Then FixedUp = Left(thename, Len(thename) - 9): Exit Function
If Right(thename, 7) = " (Busy)" Then FixedUp = Left(thename, Len(thename) - 7): Exit Function
If Right(thename, 15) = " (On The Phone)" Then FixedUp = Left(thename, Len(thename) - 15): Exit Function
If Right(thename, 7) = " (Away)" Then FixedUp = Left(thename, Len(thename) - 7): Exit Function
If Right(thename, 7) = " (Idle)" Then FixedUp = Left(thename, Len(thename) - 7): Exit Function
If Right(thename, 16) = " (Be Right Back)" Then FixedUp = Left(thename, Len(thename) - 16): Exit Function
If Right(thename, 15) = " (Out To Lunch)" Then FixedUp = Left(thename, Len(thename) - 15): Exit Function

End Function


Function WhatStatusName(statustype) As String
    If statustype = "NLN" Then WhatStatusName = " (Online)"
    If statustype = "BSY" Then WhatStatusName = " (Busy)"
    If statustype = "PHN" Then WhatStatusName = " (On The Phone)"
    If statustype = "AWY" Then WhatStatusName = " (Away)"
    If statustype = "IDL" Then WhatStatusName = " (Idle)"
    If statustype = "BRB" Then WhatStatusName = " (Be Right Back)"
    If statustype = "LUN" Then WhatStatusName = " (Out To Lunch)"
End Function

Function WhatStatusNo(statustype) As Integer
    If statustype = "NLN" Then WhatStatusNo = 4
    If statustype = "BSY" Then WhatStatusNo = 6
    If statustype = "PHN" Then WhatStatusNo = 6
    If statustype = "AWY" Then WhatStatusNo = 7
    If statustype = "IDL" Then WhatStatusNo = 7
    If statustype = "BRB" Then WhatStatusNo = 7
    If statustype = "LUN" Then WhatStatusNo = 7
End Function

Private Function GetData(sbuffer As String, Optional bTrim As Boolean = True, Optional lLength As Long = 0) As String
Dim sData As String
Dim i As Long

    If lLength = 0 Then 'The length function paramater should only be specified if your getting a payload command
        i = InStr(1, sbuffer, vbCrLf, vbTextCompare)
        Call sckMain.GetData(sData, vbString, i + 1)
    Else
        Call sckMain.GetData(sData, vbString, lLength)
    End If
    
    If bTrim = True Then 'Cut off the vbCrLf off the end, not needed for payload commands
        GetData = Mid(sData, 1, Len(sData) - 2)
    Else
        GetData = sData
    End If
    
    frmDebug.Text1.Text = frmDebug.Text1.Text & vbCrLf & vbCrLf & "------------------" & vbCrLf & vbCrLf & sData
    frmDebug.Text1.SelLength = 0
    frmDebug.Text1.SelStart = Len(frmDebug.Text1.Text)
    
    Debug.Print ">>>: " & GetData
    
End Function

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description
End Sub

Private Sub oWinHTTP_OnResponseFinished()
Dim i As Integer
    
    If Not bGotAuthRedir Then
        'Get the login url from nexus
        i = InStr(1, oWinHTTP.GetResponseHeader("PassportURLs"), "DALogin")
        sAuthLocation = "https://" & Mid$(oWinHTTP.GetResponseHeader("PassportURLs"), i + 8, InStr(i + 1, oWinHTTP.GetResponseHeader("PassportURLs"), ",") - i - 8)
        bGotAuthRedir = True
        Call oWinHTTP.Open("GET", sAuthLocation, True)
        Call oWinHTTP.SetRequestHeader("Authorization", "Passport1.4 OrgVerb=GET,OrgURL=http%3A%2F%2Fmessenger%2Emsn%2Ecom,sign-in=" & Replace(txtSigninName.Text, "@", "%40") & ",pwd=" & txtPassword.Text & "," & sChallenge)
        Call oWinHTTP.Send
        
    Else
        If IsHeader("Authentication-Info") Then
            'Check for redirection
            If InStr(1, oWinHTTP.GetResponseHeader("Authentication-Info"), "Passport1.4 da-status=redir") Then
                Call oWinHTTP.Open("GET", oWinHTTP.GetResponseHeader("Location"), True)
                Call oWinHTTP.SetRequestHeader("Authorization", "Passport1.4 OrgVerb=GET,OrgURL=http%3A%2F%2Fmessenger%2Emsn%2Ecom,sign-in=" & Replace(txtSigninName.Text, "@", "%40") & ",pwd=" & txtPassword.Text & "," & sChallenge)
                Call oWinHTTP.Send

            ElseIf InStr(1, oWinHTTP.GetResponseHeader("Authentication-Info"), "Passport1.4 da-status=success") Then
                'Successfull login, send the ticket to NS
                i = InStr(1, oWinHTTP.GetResponseHeader("Authentication-Info"), "'")
                Call SendData("USR 4 TWN S " & Mid$(oWinHTTP.GetResponseHeader("Authentication-Info"), i + 1, InStrRev(oWinHTTP.GetResponseHeader("Authentication-Info"), "'") - i - 1) & vbCrLf)
            
            End If
        
        ElseIf IsHeader("WWW-Authenticate") Then
            If InStr(1, oWinHTTP.GetResponseHeader("WWW-Authenticate"), "Passport1.4 da-status=failed") Then
                MsgBox "invalid password"
                
            End If
            
        End If

    End If

End Sub

Private Function IsHeader(sHeaderName As String) As Boolean
On Error Resume Next
Dim sValue As String

    sValue = oWinHTTP.GetResponseHeader(sHeaderName)
    
    If err.Number = 0 Then
        IsHeader = True
    End If
    
End Function
