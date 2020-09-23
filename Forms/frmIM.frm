VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIM 
   BackColor       =   &H00E8992F&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6600
   ClientLeft      =   2115
   ClientTop       =   2385
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmIM.frx":058A
   ScaleHeight     =   6600
   ScaleWidth      =   9135
   WindowState     =   1  'Minimized
   Begin VB.ListBox lstParticipants 
      Height          =   3960
      Left            =   7320
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock sckMSG 
      Left            =   6960
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox pctCon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E8992F&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   7320
      ScaleHeight     =   1185
      ScaleWidth      =   1425
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox pctMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E8992F&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   7440
      ScaleHeight     =   825
      ScaleWidth      =   1065
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7440
      Top             =   2400
   End
   Begin MSComDlg.CommonDialog d1 
      Left            =   7920
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   10
   End
   Begin VB.TextBox txtMSG 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   905
      Left            =   150
      MaxLength       =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5280
      Width           =   5880
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00E8992F&
      Caption         =   "Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   885
      Left            =   6120
      MaskColor       =   &H00E8992F&
      TabIndex        =   1
      Top             =   5280
      Width           =   735
   End
   Begin RichTextLib.RichTextBox txtHist 
      Height          =   4050
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   7144
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmIM.frx":C4BAC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgColour 
      Height          =   450
      Left            =   960
      Picture         =   "frmIM.frx":C4C27
      Top             =   4785
      Width           =   615
   End
   Begin VB.Label lblCon 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   180
      Width           =   6540
   End
   Begin VB.Image imgEmoticon 
      Height          =   450
      Left            =   1800
      Picture         =   "frmIM.frx":C5AF1
      Top             =   4785
      Width           =   240
   End
   Begin VB.Image imgFont 
      Height          =   450
      Left            =   360
      Picture         =   "frmIM.frx":C60D3
      Top             =   4790
      Width           =   420
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E8992F&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   6250
      Width           =   6615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save as..."
      End
      Begin VB.Menu mnuSPc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuCon 
      Caption         =   "Contacts"
      Begin VB.Menu mnuCAddC 
         Caption         =   "Add contact..."
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Information"
      Begin VB.Menu mnuICon 
         Caption         =   ""
      End
      Begin VB.Menu mnuIYou 
         Caption         =   "You"
      End
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public SBIndex As Integer
Public BudName As String
Public BudEmail As String
Public sStat As String
Dim ClFontSnd As String, ClFontEffects As String, ClFontColour As String, OldStat As String
Dim bOK As Boolean, MsgSnd As String

Dim theSessID

Private Function ProcessMSG(sData As String)
'On Error Resume Next
Dim sLines() As String, sNewM As String, i As Integer

    sData2 = sData
    sData2 = Replace(sData2, vbCrLf, " ")
    sLines1 = Split(sData, " ")
    sLines2 = Split(sData2, " ")
    Buffer = sData
        
        Select Case sLines2(7)
            
            Case "text/x-msmsgscontrol"
                BudName = MSNDecode(sLines1(2))
            
                OldStat = lblStat.Caption
                lblStat.Caption = BudName & " is typing a message."
                
            Case "text/plain;"
                sNewM = Mid$(sData, InStr(1, sData, vbCrLf & vbCrLf) + 4)
                
                txtHist.SelText = vbTab
                
                BudName = MSNDecode(sLines1(2))
                
                If Left(sNewM, 1) = "«" And Right(sNewM, 1) = "»" Then
                    theaction = Right(sNewM, Len(sNewM) - 3)
                    theaction = Left(theaction, Len(theaction) - 1)
                    Call LogConvo(vbCrLf & BudName & " " & theaction, "156,0,156")
                    ScrollBottom
                Else
                    Call LogConvo(vbCrLf & BudName & " says:" & vbCrLf)
                    Call DecodeFont(sLines2(10), sLines2(11), sLines2(12))
                    
                    txtHist.SelText = sNewM
                                        
                    ScrollBottom
                    lblStat.Caption = "Last message received at " & Time & "."
                    If Left(lblStat.Caption, 4) = "Last" Then
                    OldStat = lblStat.Caption
                    Timer1.Enabled = True
                    Else
                    Exit Function
                    End If
                End If
                
            Case "text/x-msmsgsinvite;"
                
            
            Case "application/x-msnmsgrp2p"
                'For i = 15 To UBound(sLines2)
                '    sNewM = sNewM & sLines2(i) & " "
                'Next i
                Call ProcessP2P(sData)
        End Select
End Function

Private Function DecodeFont(ByVal Font As String, ByVal Effect As String, ByVal Colour As String)
    'Remove some crap from the strings
    Font = Replace(Font, "FN=", ""): Font = Replace(Font, "%20", " "): Font = Replace(Font, ";", "")
    Colour = Replace(Colour, "CO=", ""): Colour = Replace(Colour, ";", "")
    
    If InStr(Effect, "B") > 0 Then txtHist.SelBold = True Else: txtHist.SelBold = False
    If InStr(Effect, "I") > 0 Then txtHist.SelItalic = True Else: txtHist.SelItalic = False
    If InStr(Effect, "S") > 0 Then txtHist.SelStrikeThru = True Else: txtHist.SelStrikeThru = False
    If InStr(Effect, "U") > 0 Then txtHist.SelUnderline = True Else: txtHist.SelUnderline = False
    
    txtHist.SelColor = bgrhex2rgb(Replace(Colour, "#", ""))
    txtHist.SelFontName = URL_Decode(Font)
    
End Function

Public Function ProcessP2P(ByVal Data As String)
    If BasePlus = 0 Then BasePlus = -3
    
    Dim rEmail As String, rEUF As String, rHeaders As String, iSessID As Long, rBranch As String, rCallID As String
    Dim sL() As String, sP() As String, i As Long, j As Long: j = 1
    Dim Fields As New Collection, PartField As String

    'Strip out fields
    sL = Split(Data, vbCrLf)
    rEmail = Split(Data, " ")(1)
    For i = 0 To UBound(sL)
        If sL(i) <> "" Then
            sP = Split(sL(i), " ")
            Select Case sP(0)
                Case "EUF-GUID:":   rEUF = sP(1)
                Case "SessionID:":  iSessID = CLng(sP(1)): theSessID = CLng(sP(1))
                Case "Call-ID:":    rCallID = sP(1)
                Case "Via:":        rBranch = "{" & GetBetween(sL(i), "{", "}") & "}"
            End Select
        End If
    Next i
    
    MsgBox Data
    
    'Get the values of the fields (leave them in their original state -> strings)
    rHeaders = GetBetween(Data, vbCrLf & vbCrLf, , 48)
    For i = 1 To Len(rHeaders)
        PartField = PartField & Mid$(rHeaders, i, 1)
        If j <> 3 And j <> 4 And j <> 9 Then
            If Len(PartField) = 4 Then
                Fields.Add PartField, CStr(j): j = j + 1: PartField = ""
            End If
        Else
            If Len(PartField) = 8 Then
                Fields.Add PartField, CStr(j): j = j + 1: PartField = ""
            End If
        End If
    Next i
    
    'Process
    If InStr(1, Data, "INVITE MSNMSGR") Then
    
        'What kind of invitation is this? You can reconize them by their EUF-GUI.
        'DP = {A4268EEC-FEC5-49E5-95C3-F126696BDBF6}
        'FTP = {5D3E02AB-6190-11D3-BBBB-00C04F795683}
        'EMO = ?
        
        'Add to dictionary? (To store their SessID)
        If dctP2PTransfers.Exists(rEmail) = False Then
            dctP2PTransfers.Add rEmail, ""
        End If
        
        Select Case rEUF
            Case "{A4268EEC-FEC5-49E5-95C3-F126696BDBF6}"
                'First we send the BaseIdentifier, this BaseID is being send to people
                'to indicate what our BaseID is. After the BaseID we can send our 200OK
                'message indicating to the client it is cleared to proceed :)
                Call SendSB(GetBaseIdent(Fields, rEmail)): DoEvents
                Call SendSB(Get200OK(Fields, rEmail, rBranch, rCallID, iSessID)): DoEvents
                dctP2PTransfers.Item(rEmail) = iSessID
        End Select
        
    ElseIf InStr(1, Data, "BYE MSNMSGR:") Then
    
        'This is the Bye Acknowledgement message, this is the very last message
        'and is sent in response to the Bye message.
        Call Me.SendSB(GetByeAck(Fields, rEmail))
        
    Else
        'Something else, check!
        'If dctP2PTransfers.Exists(rEmail) = False Then Exit Function 'Doesn't belong here
        
        'Okay, we have used predefined identifiers so we can easily check what
        'kind of message is next
        Select Case GetDWord(Fields(CStr(8)))
            Case OK200
                'This is the message after our 200OK message, it is called the
                'Data Preperation message, and is send to let the RC know we are going
                'to send him data
                Call SendSB(GetDataPrep(rEmail, dctP2PTransfers(rEmail)))
                
            Case DataPrep
                'Okay, we have done the DataPrep, now we can finally send
                'the data! :)
                Call SendDataDP(dctP2PTransfers(rEmail), rEmail): DoEvents
        End Select
    End If
End Function

Public Function SendDataDP(ByVal SessID As String, ToEmail As String)
    Dim BinHeader As String
    Dim FileD As String: FileD = DisplayPicData
    Dim TFileD As String, L As Long: L = 1
    
    SessID = 5425345
    
    'This is the hardest part I think. We first check if we have to send multiple
    'messages (each message can contain a maximum of 1202 bytes of picture).
    If Len(FileD) > 1202 Then
        
        Do
            'First we grab the data we need to send in this message
            TFileD = Mid$(FileD, L, 1202): L = L + Len(TFileD)
            
            'Then we generate the Binary header. Watch as we now use the offset
            'field (field 3). The total size of the image is put into field 4 and
            'the size of the data transferred in this message is field 5.
            BinHeader = MakeDWord(SessID) & _
                        MakeDWord(BaseIdentID + BasePlus) & _
                        MakeDWord(L - Len(TFileD) - 1) & MakeDWord(0) & _
                        MakeDWord(Len(DisplayPicData)) & MakeDWord(0) & _
                        MakeDWord(Len(TFileD)) & _
                        MakeDWord(0) & _
                        MakeDWord(DataSend) & _
                        MakeDWord(0) & _
                        MakeDWord(0) & MakeDWord(0) & _
                        TFileD & _
                        MakeDWord(1, False)
        
            'Finish
            BinHeader = "MIME-Version: 1.0" & vbCrLf & _
                        "Content-Type: application/x-msnmsgrp2p" & vbCrLf & _
                        "P2P-Dest: " & ToEmail & vbCrLf & vbCrLf & BinHeader
        
            'Send! :D
            Call SendSB("MSG *%* D " & Len(BinHeader) & vbCrLf & BinHeader)
        Loop While L < Len(FileD)
        BasePlus = IIf(BasePlus = -1, 1, BasePlus + 1)
        
    Else
        'OK, we can send the picture in one message.
        'Generate the binary header, and put the file data in it.
        'At the end there is a DWord with the value 1.
        BinHeader = MakeDWord(SessID) & _
                    MakeDWord(BaseIdentID + BasePlus) & _
                    MakeDWord(0) & MakeDWord(0) & _
                    MakeDWord(Len(FileD)) & MakeDWord(0) & _
                    MakeDWord(Len(FileD)) & _
                    MakeDWord(0) & _
                    MakeDWord(DataSend) & _
                    MakeDWord(0) & _
                    MakeDWord(0) & MakeDWord(0) & _
                    FileD & _
                    MakeDWord(1, False): BasePlus = IIf(BasePlus = -1, 1, BasePlus + 1)
        
        'Finish
        BinHeader = "MIME-Version: 1.0" & vbCrLf & _
                    "Content-Type: application/x-msnmsgrp2p" & vbCrLf & _
                    "P2P-Dest: " & ToEmail & vbCrLf & vbCrLf & BinHeader
        
        'Send! :D
        Call SendSB("MSG *%* D " & Len(BinHeader) & vbCrLf & BinHeader)
    End If
End Function

Public Function SendSB(ByVal command As String)
Dim SBTrial As Long
    'Okay, this is just a function to send data to the SB (conversation)
    'First we need to UTF-8 encode the data
    
    'Now we replace % with the correct Trial ID.
    If Left$(command, 3) = "MSG" Then 'Message
        command = Replace$(command, "*%*", SBTrial)
    Else 'Not a message
        command = Replace$(command, "*%*", SBTrial) & vbCrLf
    End If

    'Then send it
    If sckMSG.State = 7 Then
        sckMSG.SendData command
    End If

    frmDebug.Text1.Text = frmDebug.Text1.Text & vbCrLf & vbCrLf & "*****Switch*******" & vbCrLf & vbCrLf & command
    frmDebug.Text1.SelLength = 0
    frmDebug.Text1.SelStart = Len(frmDebug.Text1.Text)
    
    'Count SBTrial up
    SBTrial = SBTrial + 1
End Function

Public Function CreateMSG(MsgText As String, Optional color As String = "0", Optional Font As String = "Arial", Optional Effects As String = "", Optional ByVal P4Name As String = "") As String
    'This function creates an MSG command
    Dim iLength As Long
    Dim CMD As String
    
    'Don't forget to UTF-8 encode the message...
    If Not Left(MsgText, 4) = "/me " Then
        MsgText = EncodeUTF8(MsgText)
    Else
        MsgText = "Â«I " & EncodeUTF8(Right(MsgText, Len(MsgText) - 4)) & "Â»"
    End If
    '...and hex-encoded the font name
    Font = URL_Encode(Font)
    
    If Not P4Name = "" Then thisline = "P4-Context: " & MSNEncode(P4Name) & vbCrLf

    CMD = "MIME-Version: 1.0" & vbCrLf & _
    "Content-Type: text/plain; charset=UTF-8" & vbCrLf & _
    thisline & _
    "X-MMS-IM-Format: FN=" & Font & "; EF=" & Effects & "; CO=" & color & "; CS=0; PF=22" & vbCrLf & vbCrLf & _
    MsgText
    
    iLength = Len(CMD)
    MSGdata = "MSG *%* N " & iLength & vbCrLf & CMD
    CreateMSG = MSGdata
End Function

Public Function LogConvo(Text As String, Optional TxtColour As String = "150,150,150")
    TheColours = Split(TxtColour, ",")
    txtHist.SelStart = Len(txtHist.Text)
    txtHist.SelIndent = 0
    txtHist.SelFontName = "MS Sans Serif"
    txtHist.SelStrikeThru = False
    txtHist.SelUnderline = False
    txtHist.SelBold = False
    txtHist.SelItalic = False
    txtHist.SelColor = RGB(TheColours(0), TheColours(1), TheColours(2))
    txtHist.SelText = Text
    txtHist.SelStart = Len(txtHist.Text)
    txtHist.SelColor = RGB(0, 0, 0)
End Function

Public Function ScrollBottom()
    txtHist.SelStart = Len(txtHist.Text)
    txtHist.SelText = ""
End Function


Private Sub Form_Activate()
    mnuICon.Caption = BudName
    'lblCon.Caption = BudEmail
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    sckMSG.Close
    Unload sckMSG
    Unload Me
End Sub

Private Sub imgEmoticon_Click()
    thename = InputBox("enter name", "thename")
    MsgSnd = txtMSG.Text
    txtMSG.Text = ""
    Call SendSB(CreateMSG(MsgSnd, ClFontColour, ClFontSnd, ClFontEffects, thename)) '
End Sub

Private Sub imgFont_Click()
'On Error Resume Next
    d1.Flags = &H3 ' + &H100 '+ &H80000
    d1.ShowFont
    If d1.FontName = "" Then Exit Sub
    SaveSetting "SojMessenger", "Font", "Name", d1.FontName
    txtMSG.FontName = d1.FontName
    'ClFontSnd = Replace(d1.FontName, " ", "%20")
    ClFontSnd = d1.FontName
        
    If d1.FontBold = True Then isbold = "B"
    If d1.FontItalic = True Then isitalic = "I"
    If d1.FontStrikethru = True Then isstrike = "S"
    If d1.FontUnderline = True Then isunderline = "U"
        
    ClFontEffects = isbold & isitalic & isstrike & isunderline
    SaveSetting "SojMessenger", "Font", "FX", ClFontEffects
    DecodeEffects ClFontEffects, ClFontColour
End Sub

Private Sub imgColour_Click()
'On Error Resume Next
    d1.Flags = &H1 + &H2
    d1.ShowColor
    
    txtMSG.ForeColor = d1.color
    
    splittherup = Split(colortorgb(d1.color), ",")
    
    ClFontColour = Replace(rgbtohex(splittherup(0), splittherup(1), splittherup(2)), "#", "")
    
    'ClFontColour = Right(ClFontColour, Len(ClFontColour) - 1)
        
    SaveSetting "SojMessenger", "Font", "Colour", ClFontColour
    DecodeEffects ClFontEffects, ClFontColour
End Sub

Private Sub mnuCAddC_Click()
'theemail = InputBox("Enter the email of the contact you would like to invite.", "Add Contact")
'If Not theemail = "" Then SendSB "CAL *%* " & theemail
frmInvite.Show
End Sub

Private Sub mnuICon_Click()
    MsgBox "Email Address: " & BudEmail & vbCrLf & _
           "Display Name: " & BudName & vbCrLf _
           , vbInformation, "Information about " & BudEmail
End Sub

Private Sub mnuIYou_Click()
Dim i%, iCou%

    iCou = frmMain.lLstCount
    
    If d1.FileName = "" Then
    sLoc = App.Path & "\Display Data\helloworld.png"
    Else
    sLoc = d1.FileName
    End If
    
    MsgBox "Email Address: " & UserName & vbCrLf & _
           "Display Name: " & frmMain.lblName.Caption & vbCrLf & _
           "Status: " & frmMain.lblStatus.Caption & vbCrLf & _
           "IP Address: " & sckMSG.LocalIP & vbCrLf & _
           "Font Name: " & d1.FontName & vbCrLf & _
           "Number Of Contacts: " & iCou & " out of 150" & vbCrLf & _
           "Location Of Display Picture: " & sLoc, vbInformation, "Information About You"
End Sub


Private Sub sckMSG_Close()
    bOK = False
End Sub

Private Sub sckMSG_DataArrival(ByVal bytesTotal As Long)
'On Error Resume Next
Dim sData As String

    'This gets the data then makes it a array cos its easier to handle :)
    sckMSG.GetData sData
    sData = DecodeUTF8(sData)
    Debug.Print sData
    
    sParams = Split(sData, " ")
    
    frmDebug.Text1.Text = frmDebug.Text1.Text & vbCrLf & vbCrLf & "-----Switch-------" & vbCrLf & vbCrLf & sData
    frmDebug.Text1.SelLength = 0
    frmDebug.Text1.SelStart = Len(frmDebug.Text1.Text)
    
    Select Case Left(sData, 3)

        Case "MSG"
            ProcessMSG sData

        Case "USR"
            SendSB "CAL *%* " & BudEmail

        Case "JOI"
            bOK = True
            lblCon.Caption = lblCon.Caption & ", " & sParams(1)
            If Left(lblCon.Caption, 2) = ", " Then lblCon.Caption = Right(lblCon.Caption, Len(lblCon.Caption) - 2)
            lstParticipants.AddItem sParams(1)
            Call LogConvo(vbCrLf & "***" & MSNDecode(Left(sParams(2), Len(sParams(2)) - 2)) & " has joined the conversation.***" & vbCrLf, "0,128,0")
            ScrollBottom
            
        Case "BYE"
            bOK = False
            frmMain.SendData "XFR *%* SB"
            BuddyConnect = BudEmail
            Call LogConvo(vbCrLf & "***" & Left(sParams(1), Len(sParams(1)) - 2) & " has closed the chat window.***" & vbCrLf, "128,0,0")
            ScrollBottom

        Case "ANS"
            bOK = True
            
        Case "217"
            'lblCon.Caption = BudEmail & " - <Appears to be offline>"
            txtMSG.Enabled = False
            cmdSend.Enabled = False

    End Select
End Sub

Public Sub cmdSend_Click()

If Left(txtMSG.Text, 1) = "." Then ProcessMSG (txtMSG.Text): Exit Sub

If sckMSG.State = 7 Then
    If Not Left(txtMSG.Text, 4) = "/me " Then
        Call LogConvo(vbCrLf & frmMain.lblName.Caption & " says:" & vbCrLf)
        DecodeFont ClFontSnd, ClFontEffects, ClFontColour
        txtHist.SelText = txtMSG.Text
        ScrollBottom
    Else
        theaction = Right(txtMSG.Text, Len(txtMSG.Text) - 4)
        Call LogConvo(vbCrLf & frmMain.lblName.Caption & " " & theaction, "156,0,156")
        ScrollBottom
    End If
    MsgSnd = txtMSG.Text
    txtMSG.Text = ""
    'Do: DoEvents: Loop While bOK <> True
    Call SendSB(CreateMSG(MsgSnd, ClFontColour, ClFontSnd, ClFontEffects))
    Exit Sub
End If

If sckMSG.State <> 7 Then
    If Not Left(txtMSG.Text, 4) = "/me " Then
        Call LogConvo(vbCrLf & frmMain.lblName.Caption & " says:" & vbCrLf)
        DecodeFont ClFontSnd, ClFontEffects, ClFontColour
        txtHist.SelText = txtMSG.Text & vbCrLf
        ScrollBottom
    Else
        theaction = Right(txtMSG.Text, Len(txtMSG.Text) - 4)
        Call LogConvo(vbCrLf & frmMain.lblName.Caption & " " & theaction, "156,0,156")
        ScrollBottom
    End If
    MsgSnd = txtMSG.Text
    txtMSG.Text = ""
    BuddyConnect = BudEmail
    Call frmMain.SendData("XFR *%* SB")
    'Do: DoEvents: Loop Until sckMSG.State = 7 And bOK = True
    Call SendSB(CreateMSG(MsgSnd, ClFontColour, ClFontSnd, ClFontEffects))
    Exit Sub
End If
End Sub

Private Sub Form_Load()
On Error GoTo err

    bOK = False
    d1.FontName = GetSetting("SojMessenger", "Font", "Name")
    ClFontEffects = GetSetting("SojMessenger", "Font", "FX")
    txtMSG.FontName = d1.FontName
    'ClFontSnd = Replace(d1.FontName, " ", "%20")
    
    ClFontColour = "000000"
    
    DecodeEffects ClFontEffects, ClFontColour
    lstParticipants.AddItem frmMain.txtSigninName.Text
    
    Exit Sub
err:
d1.FontName = "MS Sans Serif"
ClFontEffects = ""
txtMSG.Font = d1.FontName
End Sub


Private Function DecodeEffects(ByVal Effect As String, ByVal TheColour As String)

    If InStr(Effect, "B") > 0 Then txtMSG.FontBold = True Else: txtMSG.FontBold = False
    If InStr(Effect, "I") > 0 Then txtMSG.FontItalic = True Else: txtMSG.FontItalic = False
    If InStr(Effect, "S") > 0 Then txtMSG.FontStrikethru = True Else: txtMSG.FontStrikethru = False
    If InStr(Effect, "U") > 0 Then txtMSG.FontUnderline = True Else: txtMSG.FontUnderline = False
    
    txtMSG.ForeColor = rgbhex2rgb(Replace(TheColour, "#", "")) 'hextorgb("#" & TheColour)
    
End Function


Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuSaveAs_Click()
On Error GoTo err
    d1.Filter = "Rich Text Format|*.rtf"
    d1.DialogTitle = "Select a folder to save the chat log in..."
    d1.ShowSave
    txtHist.SaveFile d1.FileName

err:
End Sub

Private Sub sckMSG_Connect()
    Call SendSB(sckMSG.Tag)
    sckMSG.Tag = ""
End Sub

Private Sub Timer1_Timer()
    lblStat.Caption = OldStat
    Timer1.Enabled = False
End Sub

Private Sub txtHist_Change()
    
    If Me.Visible = False Then Me.Visible = True
    
End Sub

Private Sub txtMSG_Change()
    If Len(txtMSG) = 0 Then cmdSend.Enabled = False Else cmdSend.Enabled = True
End Sub
