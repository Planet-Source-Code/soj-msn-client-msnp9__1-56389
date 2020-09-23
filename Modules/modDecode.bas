Attribute VB_Name = "modDecode"
Public BuddyConnect As String, BuddyConName As String

'Decodeing Section
'----------
Public Function MSNDecode(ByVal Utf8Str As String) As String
    Utf8Str = Replace(Utf8Str, "%20", " ")
    Utf8Str = Replace(Utf8Str, "ãƒ„", "?")
    Utf8Str = Replace(Utf8Str, "â„¢", "™")
    Utf8Str = Replace(Utf8Str, "â&#8218;¬", "&#8364;")
    Utf8Str = Replace(Utf8Str, "Â", "")
    Utf8Str = Replace(Utf8Str, "â&#8364;&#353;", "&#8218;")
    Utf8Str = Replace(Utf8Str, "Æ&#8217;", "&#402;")
    Utf8Str = Replace(Utf8Str, "â&#8364;&#382;", "&#8222;")
    Utf8Str = Replace(Utf8Str, "â&#8364;¦", "&#8230;")
    Utf8Str = Replace(Utf8Str, "â&#8364; ", "&#8224;")
    Utf8Str = Replace(Utf8Str, "â&#8364;¡", "&#8225;")
    Utf8Str = Replace(Utf8Str, "Ë&#8224;", "&#710;")
    Utf8Str = Replace(Utf8Str, "â&#8364;°", "&#8240;")
    Utf8Str = Replace(Utf8Str, "Å ", "&#352;")
    Utf8Str = Replace(Utf8Str, "â&#8364;¹", "&#8249;")
    Utf8Str = Replace(Utf8Str, "Å&#8217;", "&#338;")
    Utf8Str = Replace(Utf8Str, "Â", "")
    Utf8Str = Replace(Utf8Str, "Å½", "&#381;")
    Utf8Str = Replace(Utf8Str, "Â", "")
    Utf8Str = Replace(Utf8Str, "Â", "")
    Utf8Str = Replace(Utf8Str, "â&#8364;&#732;", "&#8216;")
    Utf8Str = Replace(Utf8Str, "â&#8364;&#8482;", "&#8217;")
    Utf8Str = Replace(Utf8Str, "â&#8364;&#339;", "&#8220;")
    Utf8Str = Replace(Utf8Str, "â&#8364;", "&#8221;")
    Utf8Str = Replace(Utf8Str, "â&#8364;¢", "&#8226;")
    Utf8Str = Replace(Utf8Str, "â&#8364;&#8220;", "&#8211;")
    Utf8Str = Replace(Utf8Str, "â&#8364;&#8221;", "&#8212;")
    Utf8Str = Replace(Utf8Str, "Ë&#339;", "&#732;")
    Utf8Str = Replace(Utf8Str, "â&#8222;¢", "&#8482;")
    Utf8Str = Replace(Utf8Str, "Å¡", "&#353;")
    Utf8Str = Replace(Utf8Str, "â&#8364;º", "&#8250;")
    Utf8Str = Replace(Utf8Str, "Å&#8220;", "&#339;")
    Utf8Str = Replace(Utf8Str, "Â", "")
    Utf8Str = Replace(Utf8Str, "'Å¾", "&#382;")
    Utf8Str = Replace(Utf8Str, "Å¸", "&#376;")
    Utf8Str = Replace(Utf8Str, "Â ", " ")
    Utf8Str = Replace(Utf8Str, "Â¡", "¡")
    Utf8Str = Replace(Utf8Str, "Â¢", "¢")
    Utf8Str = Replace(Utf8Str, "Â£", "£")
    Utf8Str = Replace(Utf8Str, "Â¤", "¤")
    Utf8Str = Replace(Utf8Str, "Â¥", "¥")
    Utf8Str = Replace(Utf8Str, "Â¦", "¦")
    Utf8Str = Replace(Utf8Str, "Â§", "§")
    Utf8Str = Replace(Utf8Str, "Â¨", "¨")
    Utf8Str = Replace(Utf8Str, "Â©", "©")
    Utf8Str = Replace(Utf8Str, "Âª", "ª")
    Utf8Str = Replace(Utf8Str, "Â«", "«")
    Utf8Str = Replace(Utf8Str, "Â¬", "¬")
    Utf8Str = Replace(Utf8Str, "Â­", "­")
    Utf8Str = Replace(Utf8Str, "Â®", "®")
    Utf8Str = Replace(Utf8Str, "Â¯", "¯")
    Utf8Str = Replace(Utf8Str, "Â°", "°")
    Utf8Str = Replace(Utf8Str, "Â±", "±")
    Utf8Str = Replace(Utf8Str, "Â²", "²")
    Utf8Str = Replace(Utf8Str, "Â³", "³")
    Utf8Str = Replace(Utf8Str, "Â´", "´")
    Utf8Str = Replace(Utf8Str, "Âµ", "µ")
    Utf8Str = Replace(Utf8Str, "Â¶", "¶")
    Utf8Str = Replace(Utf8Str, "Â·", "·")
    Utf8Str = Replace(Utf8Str, "Â¸", "¸")
    Utf8Str = Replace(Utf8Str, "Â¹", "¹")
    Utf8Str = Replace(Utf8Str, "Âº", "º")
    Utf8Str = Replace(Utf8Str, "Â»", "»")
    Utf8Str = Replace(Utf8Str, "Â¼", "¼")
    Utf8Str = Replace(Utf8Str, "Â½", "½")
    Utf8Str = Replace(Utf8Str, "Â¾", "¾")
    Utf8Str = Replace(Utf8Str, "Â¿", "¿")
    Utf8Str = Replace(Utf8Str, "Ã ", "à")
    Utf8Str = Replace(Utf8Str, "Ã¡", "á")
    Utf8Str = Replace(Utf8Str, "Ã¢", "â")
    Utf8Str = Replace(Utf8Str, "Ã£", "ã")
    Utf8Str = Replace(Utf8Str, "Ã¤", "ä")
    Utf8Str = Replace(Utf8Str, "Ã¥", "å")
    Utf8Str = Replace(Utf8Str, "Ã¦", "æ")
    Utf8Str = Replace(Utf8Str, "Ã§", "ç")
    Utf8Str = Replace(Utf8Str, "Ã¨", "è")
    Utf8Str = Replace(Utf8Str, "Ã©", "é")
    Utf8Str = Replace(Utf8Str, "Ãª", "ê")
    Utf8Str = Replace(Utf8Str, "Ã«", "ë")
    Utf8Str = Replace(Utf8Str, "Ã¬", "ì")
    Utf8Str = Replace(Utf8Str, "Ã­", "í")
    Utf8Str = Replace(Utf8Str, "Ã®", "î")
    Utf8Str = Replace(Utf8Str, "Ã¯", "ï")
    Utf8Str = Replace(Utf8Str, "Ã°", "ð")
    Utf8Str = Replace(Utf8Str, "Ã±", "ñ")
    Utf8Str = Replace(Utf8Str, "Ã²", "ò")
    Utf8Str = Replace(Utf8Str, "Ã³", "ó")
    Utf8Str = Replace(Utf8Str, "Ã´", "ô")
    Utf8Str = Replace(Utf8Str, "Ãµ", "õ")
    Utf8Str = Replace(Utf8Str, "Ã¶", "ö")
    Utf8Str = Replace(Utf8Str, "Ã·", "÷")
    Utf8Str = Replace(Utf8Str, "Ã¸", "ø")
    Utf8Str = Replace(Utf8Str, "Ã¹", "ù")
    Utf8Str = Replace(Utf8Str, "Ãº", "ú")
    Utf8Str = Replace(Utf8Str, "Ã»", "û")
    Utf8Str = Replace(Utf8Str, "Ã¼", "ü")
    Utf8Str = Replace(Utf8Str, "Ã½", "ý")
    Utf8Str = Replace(Utf8Str, "Ã¾", "þ")
    Utf8Str = Replace(Utf8Str, "Ã¿", "ÿ")
    Utf8Str = Replace(Utf8Str, "Ã&#8364;", "À")
    Utf8Str = Replace(Utf8Str, "Ã", "Á")
    Utf8Str = Replace(Utf8Str, "Ã&#8218;", "Â")
    Utf8Str = Replace(Utf8Str, "Ã&#402;", "Ã")
    Utf8Str = Replace(Utf8Str, "Ã&#8222;", "Ä")
    Utf8Str = Replace(Utf8Str, "Ã&#8230;", "Å")
    Utf8Str = Replace(Utf8Str, "Ã&#8224;", "Æ")
    Utf8Str = Replace(Utf8Str, "Ã&#8225;", "Ç")
    Utf8Str = Replace(Utf8Str, "Ã&#710;", "È")
    Utf8Str = Replace(Utf8Str, "Ã&#8240;", "É")
    Utf8Str = Replace(Utf8Str, "Ã&#352;", "Ê")
    Utf8Str = Replace(Utf8Str, "Ã&#8249;", "Ë")
    Utf8Str = Replace(Utf8Str, "Ã&#338;", "Ì")
    Utf8Str = Replace(Utf8Str, "Ã", "Í")
    Utf8Str = Replace(Utf8Str, "Ã&#381;", "Î")
    Utf8Str = Replace(Utf8Str, "Ã", "Ï")
    Utf8Str = Replace(Utf8Str, "Ã", "Ð")
    Utf8Str = Replace(Utf8Str, "Ã&#8216;", "Ñ")
    Utf8Str = Replace(Utf8Str, "Ã&#8217;", "Ò")
    Utf8Str = Replace(Utf8Str, "Ã&#8220;", "Ó")
    Utf8Str = Replace(Utf8Str, "Ã&#8221;", "Ô")
    Utf8Str = Replace(Utf8Str, "Ã&#8226;", "Õ")
    Utf8Str = Replace(Utf8Str, "Ã&#8211;", "Ö")
    Utf8Str = Replace(Utf8Str, "Ã&#8212;", "×")
    Utf8Str = Replace(Utf8Str, "Ã&#732;", "Ø")
    Utf8Str = Replace(Utf8Str, "Ã&#8482;", "Ù")
    Utf8Str = Replace(Utf8Str, "Ã&#353;", "Ú")
    Utf8Str = Replace(Utf8Str, "Ã&#8250;", "Û")
    Utf8Str = Replace(Utf8Str, "Ã&#339;", "Ü")
    Utf8Str = Replace(Utf8Str, "Ã", "Ý")
    Utf8Str = Replace(Utf8Str, "Ã&#382;", "Þ")
    Utf8Str = Replace(Utf8Str, "Ã&#376;", "ß")
    Utf8Str = Replace(Utf8Str, "%40", "@")
    Utf8Str = Replace(Utf8Str, "%2E", ".")
    Utf8Str = Replace(Utf8Str, "%20", " ")
    MSNDecode = Utf8Str
End Function


Public Function MSNEncode(ByVal sText) As String

    'sText = Replace(sText, "€", Chr(Hex(1)))
    
    MSNEncode = sText

End Function


Public Function URL_Encode(ByVal s As String) As String
    Dim sChar As String, sAsc As String, sHex As String, sName As String
    Dim i As Integer

    For i = 1 To Len(s)
        sChar = Mid$(s, i, 1)
        sAsc = Asc(sChar)
        If (sAsc > 47 And sAsc < 58) Or (sAsc > 64 And sAsc < 91) Or (sAsc > 96 And sAsc < 123) Then
            sHex = sChar
        Else
            sHex = "%" & Hex(sAsc)
        End If

        sName = sName & sHex
    Next i

    URL_Encode = sName
End Function

Public Function URL_Decode(s As String) As String
    Dim sChar As String, sHex As String, sName As String
    Dim i As Long
    
    'Get the Unicode name
    If InStr(1, s, "%") Then
        For i = 1 To Len(s)
            sChar = Mid$(s, i, 1)
            sHex = Mid$(s, i + 1, 2)
    
            If sChar = "%" Then
                sName = sName & Chr$(Val("&H" & sHex)): i = i + 2
            Else
                sName = sName & sChar
            End If
        Next i
    Else
        sName = s
    End If
    
    URL_Decode = sName
End Function


Public Function GetBetween(Str As String, Optional dStart As String, Optional dEnd As String, Optional Length As Long) As String
    Dim x1 As Long, x2 As Long
    
    'Start?
    x1 = IIf(dStart = "", 1, InStr(1, LCase$(Str), LCase$(dStart)) + Len(dStart))
    
    'Rip the string :0
    If x1 > 0 Then
        If dEnd = "" Then
            GetBetween = Mid$(Str, x1)
        Else
            x2 = InStr(x1, LCase$(Str), LCase$(dEnd)) - x1
            If x2 > 0 Then
                GetBetween = Mid$(Str, x1, x2)
            Else
                GetBetween = "n/f"
            End If
        End If
    Else
        GetBetween = "n/f"
    End If
    
    'Length?
    If Length > 0 And GetBetween <> "n/f" Then GetBetween = Left$(GetBetween, Length)
End Function
