Attribute VB_Name = "modMSG"
Option Compare Binary
'Option Explicit

Public Const CP_ACP = 0
Public Const CP_UTF8 = 65001

Public Type FONTSIGNATURE
        fsUsb(4) As Long
        fsCsb(2) As Long
End Type

Public Type CHARSETINFO
        ciCharset As Long
        ciACP As Long
        fs As FONTSIGNATURE
End Type

Public Const LOCALE_IDEFAULTCODEPAGE = &HB
Public Const LOCALE_IDEFAULTANSICODEPAGE = &H1004
Public Const TCI_SRCCODEPAGE = 2

Public Declare Function GetACP Lib "kernel32" () As Long
Public Declare Function GetLocaleInfoA Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function TranslateCharsetInfo Lib "gdi32" (lpSrc As Long, lpcs As CHARSETINFO, ByVal dwFlags As Long) As Long

Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long


Public Function bgrhex2rgb(code) As String
On Error Resume Next
  Dim newcode As String
  newcode = VBA.String(6 - Len(code), "0") & code
  If Len(newcode) = 7 Then newcode = Right(newcode, Len(newcode) - 1)
  bgrhex2rgb = RGB(Val("&H" & VBA.Right(newcode, 2)), Val("&H" & VBA.Mid(newcode, 3, 2)), Val("&H" & VBA.Left(newcode, 2)))
  'bghex2rgb = hextorgb(code)
End Function

Public Function rgbhex2rgb(code) As String
On Error Resume Next
  Dim newcode As String
  newcode = VBA.String(6 - Len(code), "0") & code
  'rgbhex2rgb = RGB(Val("&H" & VBA.Right(newcode, 2)), Val("&H" & VBA.Mid(newcode, 3, 2)), Val("&H" & VBA.Left(newcode, 2)))
  'rgbhex2rgb = RGB(Val("&H" & VBA.Left(newcode, 2)), Val("&H" & VBA.Mid(newcode, 3, 2)), Val("&H" & VBA.Right(newcode, 2)))
  rgbhex2rgb = bgrhex2rgb(code)
End Function


Public Function hextorgb(ByVal hexcolor As String)
'input format = #FFCC00
Dim r, g, b As Byte
r = "&H" & Mid(hexcolor, 2, 2)
g = "&H" & Mid(hexcolor, 4, 2)
b = "&H" & Mid(hexcolor, 6, 2)
hextorgb = r & "," & g & "," & b
End Function

Public Function AToW(ByVal st As String, Optional ByVal cpg As Long = -1, Optional lFlags As Long = 0) As String
    Dim stBuffer As String
    Dim cwch As Long
    Dim pwz As Long
    Dim pwzBuffer As Long
        
    If cpg = -1 Then cpg = GetACP()
    pwz = StrPtr(st)
    cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, 0&, 0&)
    stBuffer = String$(cwch + 1, vbNullChar)
    pwzBuffer = StrPtr(stBuffer)
    cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer))
    AToW = Left$(stBuffer, cwch - 1)
End Function

Public Function WToA(ByVal st As String, Optional ByVal cpg As Long = -1, Optional lFlags As Long = 0) As String
    Dim stBuffer As String
    Dim cwch As Long
    Dim pwz As Long
    Dim pwzBuffer As Long
    
    If cpg = -1 Then cpg = GetACP()
    pwz = StrPtr(st)
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, 0&, 0&, ByVal 0&, ByVal 0&)
    stBuffer = String$(cwch + 1, vbNullChar)
    pwzBuffer = StrPtr(stBuffer)
    cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, pwzBuffer, Len(stBuffer), ByVal 0&, ByVal 0&)
    WToA = Left$(stBuffer, cwch - 1)
End Function

Public Function EncodeUTF8(ByVal cnvUni As String)
    Dim sCodePage, cv As String
    sCodePage = CP_UTF8
    
    If cnvUni = vbNullString Then Exit Function
    cv = StrConv(WToA(cnvUni, sCodePage, 0), vbUnicode)
    EncodeUTF8 = Trim$(Replace$(cv, Chr(0), ""))
End Function

Public Function DecodeUTF8(ByVal cnvUni As String)
    Dim sCodePage, cnvUni2
    sCodePage = CP_UTF8
    
    If cnvUni = vbNullString Then Exit Function
    cnvUni2 = WToA(cnvUni, CP_ACP)
    DecodeUTF8 = AToW(cnvUni2, sCodePage)
End Function

Public Function Decode(ByVal s As String) As String
    Dim sChar As String, sHex As String, sName As String
    Dim i As Integer

    sName = ""
    For i = 1 To Len(s)
        sChar = Mid$(s, i, 1)
        sHex = Mid$(s, i + 1, 2)
        'Check if it is a hexcode...
        If sChar = "%" Then
            sName = sName & Chr(Val("&H" & sHex))
            i = i + 2
        Else
            sName = sName & sChar
        End If
    Next i

    Decode = sName
End Function

Public Function Encode(ByVal s As String) As String
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

    Encode = sName
End Function


Public Function rgbtohex(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)

'input format = 255,255,255

'get the r value
If r < 16 Then
hex1 = 0 & Hex(r)
Else
hex1 = Hex(r)
End If


'get the g value
If r < 16 Then
hex2 = 0 & Hex(g)
Else
hex2 = Hex(g)
End If


'get the b value
If b < 16 Then
hex3 = 0 & Hex(b)
Else
hex3 = Hex(b)
End If

'rgbtohex = "#" & hex1 & hex2 & hex3
rgbtohex = "#" & hex1 & hex2 & hex3
End Function


Public Function rgbtocolor(r As Byte, g As Byte, b As Byte)

rgbtocolor = r + (g * 256) + (b * 65536)
End Function


Public Function colortorgb(color As Long)

Dim r, g, b As Byte
r = color And 255
g = (color \ 256) And 255
b = (color \ 65536) And 255
colortorgb = r & "," & g & "," & b
End Function
