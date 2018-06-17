
'Class for sending SOAP 1.1 request 
 Class ServiceRequest 

    Public sWebServiceURL
    Public sSOAPRequest

    Private sResponse
    Private oWinHttp 
    Private sContentType 
    Private sUid, sPw
    Private sAuthValue
    ' Public servicename 
  
Private Sub Class_Initialize 
    Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1") 
    ' Web Service Content Type 
    sContentType ="text/xml;charset=UTF-8" 
End Sub 

Public Function SetUidPw(user, pw)
    sUid = user
    sPw = pw
    SetUidPw = True
End Function

'Public Function setHTTPheader(key, value)
'    sAuthKey = key
'    sAuthValue = value
'End Function
    'sAuth = Base64Encode(sUser & ":" & sPw)
    'objSOAP.setHTTPheader "Authorization", sAuth
  
Public Function SetSoapAction(servicename) 
    sWebServiceURL =  servicename
    SetSoapAction = True 
End Function 

Public Function OpenConnection
    'Open HTTP connection  
    oWinHttp.Open "POST", sWebServiceURL, False
    'Setting request headers  
    oWinHttp.setRequestHeader "Content-Type", sContentType
    ' Prepare for Basic Authentication
    sAuthValue = Base64Encode(sUid & ":" & sPw)
    oWinHttp.SetRequestHeader "Authorization", "Basic " & sAuthValue    
    OpenConnection = true 
End Function
  
Public Function SendRequest 
    'Send SOAP request 
    oWinHttp.Send  sSOAPRequest 
    'Get XML Response 
    sResponse = oWinHttp.ResponseText
    SendRequest = True
End Function 

Public Function getResponse
     getResponse = sResponse
End Function

Public Function getElementText(element)
    Dim oXML, oChdNd, nodes, node
    Set oXML = CreateObject("Microsoft.XMLDOM")
    'Load the XML Received
    oXML.LoadXML(getResponse)
    Set nodes = oXML.getElementsByTagName(element)
    getElementText = nodes(0).text
End Function
  
Public Sub Close 
    Set oWinHttp = Nothing 
End Sub

Private Function Base64Encode(sText)
    Dim oXML, oNode

    'Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oXML = CreateObject("Microsoft.XMLDOM")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
Private Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

End Class ' CServiceRequest