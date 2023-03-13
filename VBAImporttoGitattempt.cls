'Declare our variables related to our URL.
Dim username As String
Dim repo_name As String
Dim file_name As String
Dim access_token As String
Dim payload As String


'Declare variables related to the HTTP request
Dim xml_obj As MSXML2.XMLHTTP60


'Declare variables related to the visual basic editor
Dim VBAEditor As VBIDE.VBE
Dim VBProj As VBIDE.VBProject
Dim VBCodeMod As VBIDE.CodeModule
Dim VBRawcode As String


'Create a reference to the VB editor
Set VBAEditor = Application.VBE

'Grab the Visual Project
Set VBProj = VBAEditor.VBProjects(2)

' Reference a single component and then grab the code module.
Set VBCodeMod = VBProj.VBComponents.Item("Module1").CodeModule

'Grab the raw code in the code module
VBRawcode = VBCodeMod.Lines(StartLine:=1, Count:=VBCodeMod.CountOfLines)

'Define our XML HTTP object

Set xml_obj = New MSXML2.XMLHTTP60

    'Define URL
    base_url = "http://api.github.com/repos/"
    repo_name = "VBA-challenge/"
    username = "wteklay/"
    file_name = "Challenge.vb"
    access_token = "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"
    
    'Build the full URL
    full_url = base_url + username + repo_name + "contents/" + file_name + "?ref=master"
    
    'Open a new request
    xml_obj.Open bstrMethod:="PUT", bstrURL:=full_url, varAsync:=True
    
    'Set the headers
    xml_obj.setRequestHeader bstrHeader:="Accept", bstrValue:="application/vnd.github.v3+json"
    xml_obj.setRequestHeader bstrHeader:="Authorization", bstrValue:="token " + access_token
    
    'Define the payload
    payload = "{""message"": ""This is my message"", ""content"":"""
    payload = payload + Application.Clean(RawCodeEncoded)
    payload = payload + """}"
    
    'Send the request
    xml_obj.send varBody:=payload
    
    
    While xml_obj.readyState <> 4
        DoEvents
    Wend
    
    
RawCodeEncoded = EncodeBase64(text:=VBRawcode)


End Sub

Function EncodeBase64(text As String) As String
'Define variables

Dim arrData() As Byte
Dim object As MSXML2.DOMDocument60
Dim objNode As MSXML2.IXMLDOMElement

'convert out string to a Unicode String
arrData = StrConv(text, vbFromUnicode)

'Define our DOM objects
Set objxml = New MSXML2.DOMDocument60
Set objNode = objxml.createElement("b64")

'Define the data type
objNode.DataType = "bin.base64"

'Assign the node value
objNode.nodeTypedValue = arrData

'Return the encoded text
EncodeBase64 = Replace(objNode.text, vbLf, "")

'Memory cleanup
Set objNode = Nothing
Set objxml = Nothing

End Function
