Public Function Base64Sha1(inputText As String, Optional secretKey = "") As String

    Dim asc As Object
    Dim enc As Object
    Dim textToHash() As Byte
    Dim SharedSecretKey() As Byte
    Dim bytes() As Byte
    
    If secretKey = "" Then secretKey = inputText
    
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")

    textToHash = asc.GetBytes_4(inputText)
    SharedSecretKey = asc.GetBytes_4(secretKey)
    enc.Key = SharedSecretKey

    bytes = enc.ComputeHash_2((textToHash))
    Base64Sha1 = EncodeBase64(bytes)

End Function

Private Function EncodeBase64(arrData() As Byte) As String

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text

End Function

Sub TestMe()
    
    Debug.Print Base64Sha1("asdf", "ThisIsTheSecretKey") = "DSmGEC8dUW9xRs+YfAPji59dxCM="
    Debug.Print Base64Sha1("asdf") = "qIQmNGgreJRqJroWUUu0MxLq2oo="
    Debug.Print Base64Sha1("asdf", "asdf") = "qIQmNGgreJRqJroWUUu0MxLq2oo="
    
End Sub
