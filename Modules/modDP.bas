Attribute VB_Name = "modDP"

Public MSNObject As String
Public DisplayPicData As String
Public dplocation As String


Public Const BaseIdentID = 1000 '<-- Our predefined BaseID, we constantly change this using BasePlus
Public Const OK200 = 665544 '<-- 200OK message has been sent
Public Const DataPrep = 332211 '<-- Data Preperation has been sent
Public Const DataSend = 996633 '<-- Data has been sent
Public BasePlus As Long '<-- We use this variable to create the BaseIdentifier thruout the session :)
Public dctP2PTransfers As New Scripting.Dictionary

Public Function CreateMSNObject()
        
    dplocation = ""
    'First read in the file
    If dplocation <> "" Then
        Open dplocation For Binary As #1
            DisplayPicData = Input(LOF(1), #1)
            sLoc = dplocation
        Close #1
    Else
        Open App.Path & "\Display Data\image.bmp" For Binary As #1
            DisplayPicData = Input(LOF(1), #1)
            sLoc = App.Path & "\Display Data\image.bmp"
        Close #1
    End If
        
    'Create the MSNObject
    Dim Obj As String
    Obj = "<msnobj Creator=""" & frmMain.txtSigninName.Text & """ " & _
          "Size=""" & Len(DisplayPicData) & """ " & _
          "Type=""3"" Location=""dpdata.tmp"" Friendly=""AAA="" " & _
          "SHA1D=""" & Base64Encode(HexToBin(SHAHash(DisplayPicData))) & """ "
    
    'Create the SHA1C hash
    Dim SHA1C As String, SHAArray() As String, TSha As String, i As Long
    SHAArray = Split(Trim$(Obj), " "): SHAArray(0) = ""
    For i = 1 To UBound(SHAArray) ' - 1
        SHAArray(i) = GetBetween(SHAArray(i), , "=") & GetBetween(SHAArray(i), "=""", """")
    Next i
    
    'Finish object
    SHA1C = Join$(SHAArray, "")
    SHA1C = Base64Encode(HexToBin(SHAHash(SHA1C)))
    MSNObject = Obj & "SHA1C=""" & SHA1C & """/>"
    
End Function


Public Function HexToBin(ByVal Data As String)
    Dim DataOut As String, x As Long, sHex As String
    For x = 1 To Len(Data) Step 2
        sHex = Mid$(Data, x, 2)
        DataOut = DataOut & Chr(Val("&H" & sHex))
    Next
    HexToBin = DataOut
End Function



Function Base10ToBinary(ByVal Base10 As Long) As String
    Dim PrevResult As Integer, CurResult As Integer
    If Base10 = 0 Then
        Base10ToBinary = "0"
        Exit Function
    End If
    Do
        CurResult = Int(Log(Base10) / Log(2))
        If PrevResult = 0 Then PrevResult = CurResult + 1
        Base10ToBinary = Base10ToBinary & String$(PrevResult - CurResult - 1, "0") & "1"
        Base10 = Base10 - 2 ^ CurResult
        PrevResult = CurResult
    Loop Until Base10 = 0
    Base10ToBinary = Base10ToBinary & String$(CurResult, "0")
End Function

Function BinaryToBase10(ByVal Binary As String) As Long
    Dim i As Integer
    For i = Len(Binary) To 1 Step -1
        BinaryToBase10 = BinaryToBase10 + Val(Mid$(Binary, i, 1)) * 2 ^ (Len(Binary) - i)
    Next
End Function

Sub Bin3x8To4x6(ByVal Bin1Len8 As String, ByVal Bin2Len8 As String, ByVal Bin3Len8 As String, ByRef Bin1Len6 As String, ByRef Bin2Len6 As String, ByRef Bin3Len6 As String, ByRef Bin4Len6 As String)
    Bin1Len8 = Right$("0000000" & Bin1Len8, 8)
    Bin2Len8 = Right$("0000000" & Bin2Len8, 8)
    Bin3Len8 = Right$("0000000" & Bin3Len8, 8)
    Bin1Len6 = Left$(Bin1Len8, 6)
    Bin2Len6 = Right$(Bin1Len8, 2) & Left$(Bin2Len8, 4)
    Bin3Len6 = Right$(Bin2Len8, 4) & Left$(Bin3Len8, 2)
    Bin4Len6 = Right$(Bin3Len8, 6)
End Sub

Sub Bin4x6To3x8(ByVal Bin1Len6 As String, ByVal Bin2Len6 As String, ByVal Bin3Len6 As String, ByVal Bin4Len6 As String, ByRef Bin1Len8 As String, ByRef Bin2Len8 As String, ByRef Bin3Len8 As String)
    Bin1Len6 = Right$("00000" & Bin1Len6, 6)
    Bin2Len6 = Right$("00000" & Bin2Len6, 6)
    Bin3Len6 = Right$("00000" & Bin3Len6, 6)
    Bin4Len6 = Right$("00000" & Bin4Len6, 6)
    Bin1Len8 = Bin1Len6 & Left$(Bin2Len6, 2)
    Bin2Len8 = Right$(Bin2Len6, 4) & Left$(Bin3Len6, 4)
    Bin3Len8 = Right$(Bin3Len6, 2) & Bin4Len6
End Sub

Function Base64Encode(ByVal NormalString As String, Optional ByVal Break As Integer = 0) As String
    Dim i As Integer, Bin1Len8 As String, Bin2Len8 As String, Bin3Len8 As String
    Dim Bin1Len6 As String, Bin2Len6 As String, Bin3Len6 As String, Bin4Len6 As String
    If NormalString = vbNullString Then Exit Function
    
    For i = 1 To Len(NormalString) - 3 Step 3
        Bin1Len8 = Base10ToBinary(Asc(Mid$(NormalString, i, 1)))
        Bin2Len8 = Base10ToBinary(Asc(Mid$(NormalString, i + 1, 1)))
        Bin3Len8 = Base10ToBinary(Asc(Mid$(NormalString, i + 2, 1)))
        Call Bin3x8To4x6(Bin1Len8, Bin2Len8, Bin3Len8, Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6)
        Base64Encode = Base64Encode & Mid$(Base64Chars, BinaryToBase10(Bin1Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid$(Base64Chars, BinaryToBase10(Bin2Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid$(Base64Chars, BinaryToBase10(Bin3Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid$(Base64Chars, BinaryToBase10(Bin4Len6) + 1, 1)
    Next
    NormalString = Right$(NormalString, Len(NormalString) - IIf(Len(NormalString) / 3 = Int(Len(NormalString) / 3), Len(NormalString) - 3, Int(Len(NormalString) / 3) * 3))
    Bin1Len8 = Base10ToBinary(Asc(Left$(NormalString, 1)))

    If Len(NormalString) >= 2 Then Bin2Len8 = Base10ToBinary(Asc(Mid$(NormalString, 2, 1))) Else Bin2Len8 = "0"
    If Len(NormalString) = 3 Then Bin3Len8 = Base10ToBinary(Asc(Right$(NormalString, 1))) Else Bin3Len8 = "0"
    Call Bin3x8To4x6(Bin1Len8, Bin2Len8, Bin3Len8, Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6)
    Base64Encode = Base64Encode & Mid$(Base64Chars, BinaryToBase10(Bin1Len6) + 1, 1)
    Base64Encode = Base64Encode & Mid$(Base64Chars, BinaryToBase10(Bin2Len6) + 1, 1)
    Base64Encode = Base64Encode & IIf(Len(NormalString) >= 2, Mid(Base64Chars, BinaryToBase10(Bin3Len6) + 1, 1), "=")
    Base64Encode = Base64Encode & IIf(Len(NormalString) = 3, Mid(Base64Chars, BinaryToBase10(Bin4Len6) + 1, 1), "=")
    If Break > 0 Then
        i = Break + 1
        While i < Len(Base64Encode)
            Base64Encode = Left$(Base64Encode, i - 1) & vbCrLf & Mid$(Base64Encode, i)
            i = i + Break + 2
        Wend
    End If
End Function

Function Base64Decode(ByVal Base64String As String) As String
    Dim i As Long, Bin1Len8 As String, Bin2Len8 As String, Bin3Len8 As String
    Dim Bin1Len6 As String, Bin2Len6 As String, Bin3Len6 As String, Bin4Len6 As String

    Base64String = Replace$(Base64String, vbCr, "")
    Base64String = Replace$(Base64String, vbLf, "")

    If Base64String = vbNullString Then Exit Function
    For i = 0 To 255
        If InStr(Base64String, Chr$(i)) > 0 And Not _
            ((InStr(Base64Chars, Chr$(i)) > 0) Or (i = Asc("="))) Then Exit Function
    Next
    If Not Len(Base64String) / 4 = Len(Base64String) \ 4 Then Exit Function

    For i = 1 To Len(Base64String) Step 4
        Bin1Len6 = Base10ToBinary(InStr(Base64Chars, Mid$(Base64String, i, 1)) - 1)
        Bin2Len6 = Base10ToBinary(InStr(Base64Chars, Mid$(Base64String, i + 1, 1)) - 1)
        If Mid$(Base64String, i + 2, 1) = "=" Then Bin3Len6 = "0" Else Bin3Len6 = Base10ToBinary(InStr(Base64Chars, Mid$(Base64String, i + 2, 1)) - 1)
        If Mid$(Base64String, i + 3, 1) = "=" Then Bin4Len6 = "0" Else Bin4Len6 = Base10ToBinary(InStr(Base64Chars, Mid$(Base64String, i + 3, 1)) - 1)
        Call Bin4x6To3x8(Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6, Bin1Len8, Bin2Len8, Bin3Len8)

        Base64Decode = Base64Decode & Chr(BinaryToBase10(Bin1Len8))
        If Not Mid$(Base64String, i + 2, 1) = "=" Then Base64Decode = Base64Decode & Chr$(BinaryToBase10(Bin2Len8))
        If Not Mid$(Base64String, i + 3, 1) = "=" Then Base64Decode = Base64Decode & Chr$(BinaryToBase10(Bin3Len8))
    Next i
End Function



Public Function GetBaseIdent(DataFields As Collection, ToEmail As String)
    Dim ReturnMSG As String, BinHeader As String
    
    'First compile the binary header
    'A binary header exists of 6 DWords and 3 QWords. A QWord can simply
    'be made by putting two DWords togehter (VB6 doesn't support QWords!)
    
    'At the end of each message there is another DWord with a Littl Endian order.
    'This DWord is set to 0 while not transferring image data.
    
    'Note that you can use this BaseIdent function with FTP's aswell
    BinHeader = MakeDWord(0) & _
                MakeDWord(BaseIdentID) & _
                MakeDWord(0) & MakeDWord(0) & _
                DataFields(CStr(4)) & _
                MakeDWord(0) & _
                MakeDWord(2) & _
                DataFields(CStr(2)) & _
                DataFields(CStr(7)) & _
                DataFields(CStr(4)) & _
                MakeDWord(0, False)

    'Then the other header stuff...
    BinHeader = "MIME-Version: 1.0" & vbCrLf & _
                "Content-Type: application/x-msnmsgrp2p" & vbCrLf & _
                "P2P-Dest: " & ToEmail & vbCrLf & vbCrLf & BinHeader

    'Then return the data! :)
    GetBaseIdent = "MSG % D " & Len(BinHeader) & vbCrLf & BinHeader
End Function

Public Function Get200OK(DataFields As Collection, ToEmail As String, Via As String, CallID As String, iSessID As Long)
    Dim ReturnMSG As String, BinHeader As String, Data As String, MSNSLP As String
    
    'The data in the message of the 200OK message is the SessionID that has been
    'send to us by the client :). The data is always followed with an 0x00 character!
    Data = vbCrLf & vbCrLf & "SessionID: " & iSessID & Chr$(0)
    
    'First we generate our MSNSLP message, this message is vital to the 200OK.
    'Note that the two \r\n's are also being counted at Content-Length
    MSNSLP = "MSNSLP/1.0 200 OK" & vbCrLf & _
             "To: <msnmsgr:" & ToEmail & ">" & vbCrLf & _
             "From: <msnmsgr:" & basNs.UserName & ">" & vbCrLf & _
             "Via: MSNSLP/1.0/TLP ;branch=" & Via & vbCrLf & _
             "CSeq: 1 " & vbCrLf & _
             "Call-ID: " & CallID & vbCrLf & _
             "Max-Forwards: 0" & vbCrLf & _
             "Content-Type: application/x-msnmsgr-sessionreqbody" & vbCrLf & _
             "Content-Length: " & Len(Data) & Data
    
    'Then create the binary header again (see BaseIdent function for more info)
    BinHeader = MakeDWord(0) & _
                MakeDWord(BaseIdentID + BasePlus) & _
                MakeDWord(0) & MakeDWord(0) & _
                MakeDWord(Len(MSNSLP)) & MakeDWord(0) & _
                MakeDWord(Len(MSNSLP)) & _
                MakeDWord(0) & _
                MakeDWord(OK200) & _
                MakeDWord(0) & _
                MakeDWord(0) & MakeDWord(0): BasePlus = IIf(BasePlus = -1, 1, BasePlus + 1)

    'Create the final message
    MSNSLP = "MIME-Version: 1.0" & vbCrLf & _
             "Content-Type: application/x-msnmsgrp2p" & vbCrLf & _
             "P2P-Dest: " & ToEmail & vbCrLf & vbCrLf & BinHeader & _
             MSNSLP & MakeDWord(0, False)
    
    'Return!
    Get200OK = "MSG % D " & Len(MSNSLP) & vbCrLf & MSNSLP
End Function




Public Function GetByeAck(DataFields As Collection, ToEmail As String)
    Dim BinHeader As String
    
    'Create the binary header
    BinHeader = MakeDWord(0) & _
                MakeDWord(BaseIdentID + BasePlus) & _
                MakeDWord(0) & MakeDWord(0) & _
                DataFields(CStr(4)) & _
                MakeDWord(0) & _
                MakeDWord(2) & _
                DataFields(CStr(2)) & _
                DataFields(CStr(7)) & _
                DataFields(CStr(4)) & _
                MakeDWord(0, False): BasePlus = IIf(BasePlus = -1, 1, BasePlus + 1)

    'Complete message
    BinHeader = "MIME-Version: 1.0" & vbCrLf & _
                "Content-Type: application/x-msnmsgrp2p" & vbCrLf & _
                "P2P-Dest: " & ToEmail & vbCrLf & vbCrLf & BinHeader
                
    'Return!
    GetByeAck = "MSG % D " & Len(BinHeader) & vbCrLf & BinHeader
End Function


Public Function GetDataPrep(ToEmail As String, SessID As String)
    Dim BinHeader As String
    
    'Again, create a binary header. In the DataPrep message 4 0x00 characters are being
    'sent so the client will know when the data is comming.
    BinHeader = MakeDWord(SessID) & _
                MakeDWord(BaseIdentID + BasePlus) & _
                MakeDWord(0) & MakeDWord(0) & _
                MakeDWord(4) & MakeDWord(0) & _
                MakeDWord(4) & _
                MakeDWord(0) & _
                MakeDWord(DataPrep) & _
                MakeDWord(0) & _
                MakeDWord(0) & MakeDWord(0) & _
                Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & _
                MakeDWord(1, False): BasePlus = IIf(BasePlus = -1, 1, BasePlus + 1)

    'Complete message
    BinHeader = "MIME-Version: 1.0" & vbCrLf & _
                "Content-Type: application/x-msnmsgrp2p" & vbCrLf & _
                "P2P-Dest: " & ToEmail & vbCrLf & vbCrLf & BinHeader
                
    'Return!
    GetDataPrep = "MSG % D " & Len(BinHeader) & vbCrLf & BinHeader
End Function


