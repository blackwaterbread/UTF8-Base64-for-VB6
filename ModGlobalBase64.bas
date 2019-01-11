Attribute VB_Name = "ModGlobalBase64"


Private Const Base64Char As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

Public Function Base64Encode(ByVal strToEncode As String) As String

    On Error Resume Next
    
    Dim bArray() As Byte, I As Long, n1 As Long, n2 As Long, n3 As Long, c1 As Long, c2 As Long, c3 As Long, c4 As Long
    Dim ReByte() As Byte
    
    If strToEncode = "" Then Exit Function
    
    bArray() = ConvertStringToUtf8Bytes(strToEncode)
    
    For I = 0 To UBound(bArray) Step 3
    
        n1 = CLng(bArray(I))
        If I + 1 <= UBound(bArray) Then n2 = CLng(bArray(I + 1)) Else n2 = -1
        If I + 2 <= UBound(bArray) Then n3 = CLng(bArray(I + 2)) Else n3 = -1
        c1 = -1: c2 = -1: c3 = -1: c4 = -1
        c1 = n1 \ 4
        c2 = (n1 And 3) * 16
        If n2 >= 0 Then c2 = c2 + (n2 \ 16): c3 = (n2 And 15) * 4
        If n3 >= 0 Then c3 = c3 + (n3 \ 64): c4 = n3 And 63
        Base64Encode = Base64Encode & Mid$(Base64Char, c1 + 1, 1)
        Base64Encode = Base64Encode & Mid$(Base64Char, c2 + 1, 1)
        If c3 >= 0 Then Base64Encode = Base64Encode & Mid$(Base64Char, c3 + 1, 1)
        If c4 >= 0 Then Base64Encode = Base64Encode & Mid$(Base64Char, c4 + 1, 1)
    
    Next
    
    Base64Encode = Base64Encode & String$(((UBound(bArray) + 1) * 8) Mod 3, "=")
    
    
    
End Function

Public Function Base64Decode(ByVal strToDecode As String) As String

    On Error Resume Next
    Dim DecodedBytes() As Byte, Length As Long, w1 As Long, w2 As Long, w3 As Long, w4 As Long, c1 As Long, c2 As Long, c3 As Long, I As Long, j As Long
    
    If strToDecode = "" Then Exit Function
    
    strToDecode = RemoveElseCharacters(strToDecode)
    Length = Int(Len(Replace$(strToDecode, "=", "")) * 0.75)
    
    ReDim DecodedBytes(Length - 1) As Byte
    
    j = 0
    
    For I = 1 To Len(strToDecode) Step 4
        w1 = InStr(Base64Char, Mid$(strToDecode, I, 1)) - 1
        w2 = InStr(Base64Char, Mid$(strToDecode, I + 1, 1)) - 1
        If Mid$(strToDecode, I + 2, 1) <> "=" Then w3 = InStr(Base64Char, Mid$(strToDecode, I + 2, 1)) - 1 Else w3 = -1
        If Mid$(strToDecode, I + 3, 1) <> "=" Then w4 = InStr(Base64Char, Mid$(strToDecode, I + 3, 1)) - 1 Else w4 = -1
        c1 = -1: c2 = -1: c3 = -1
        c1 = w1 * 4 + (w2 \ 16)
        c2 = (w2 And 15) * 16
        If w3 >= 0 Then
            c2 = c2 + (w3 \ 4)
            c3 = (w3 And 3) * 64
        End If
        If w4 >= 0 Then
            c3 = c3 + w4
        End If
        DecodedBytes(j) = CByte(c1 And &HFF)
        If UBound(DecodedBytes) >= j + 1 Then DecodedBytes(j + 1) = CByte(c2 And &HFF)
        If c3 >= 0 Then DecodedBytes(j + 2) = CByte(c3 And &HFF)
        j = j + 3
    Next
    
    Base64Decode = ConvertUtf8BytesToString(DecodedBytes())

    
End Function

Private Function RemoveElseCharacters(ByVal strToProcess As String) As String
    On Error Resume Next
    Static oRegExp As Object
    Dim sProcess As String, I As Long
    If ObjPtr(oRegExp) = 0 Then Set oRegExp = CreateObject("VBScript.RegExp")
    If ObjPtr(oRegExp) Then
        oRegExp.Global = True
        oRegExp.Pattern = "[^A-Za-z0-9\+\/\=]"
        RemoveElseCharacters = oRegExp.Replace(strToProcess, "")
    Else
        For I = 1 To Len(strToProcess)
            If InStr(Base64Char, Mid$(strToProcess, I, 1)) Then
                sProcess = sProcess & Mid$(strToProcess, I, 1)
            End If
        Next
        RemoveElseCharacters = sProcess
    End If
End Function


'Declare Need:
'Microsoft ActiveX data Objects 2.5 Library
Public Function ConvertStringToUtf8Bytes(ByRef strText As String) As Byte()

    Dim objStream As ADODB.Stream
    Dim data() As Byte
   
    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.Mode = adModeReadWrite
    objStream.Type = adTypeText
    objStream.Open
   
    ' write bytes into stream
    objStream.WriteText strText
    objStream.Flush
   
    ' rewind stream and read text
    objStream.Position = 0
    objStream.Type = adTypeBinary
    objStream.Read 3
    data = objStream.Read()
   
    ' close up and return
    objStream.Close
    ConvertStringToUtf8Bytes = data


End Function


Public Function ConvertUtf8BytesToString(ByRef data() As Byte) As String


    Dim objStream As ADODB.Stream
    Dim strtmp As String
   
    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.Mode = adModeReadWrite
    objStream.Type = adTypeBinary
    objStream.Open
   
    ' write bytes into stream
    objStream.Write data
    objStream.Flush
   
    ' rewind stream and read text
    objStream.Position = 0
    objStream.Type = adTypeText
    strtmp = objStream.ReadText
   
    ' close up and return
    objStream.Close
    ConvertUtf8BytesToString = strtmp


End Function
