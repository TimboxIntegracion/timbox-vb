VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private strWSDLUrl As String
Private strUserName As String
Private strPassword As String

Public Function Timbrar(ByVal sXml As String) As String
    Dim strEnvelope As String
           
    ' Cuerpo de la peticion de timbrar
    strEnvelope = ""
    strEnvelope = strEnvelope & "<soapenv:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:urn=""urn:WashOut"">"
    strEnvelope = strEnvelope & "<soapenv:Header/>"
    strEnvelope = strEnvelope & "<soapenv:Body>"
    strEnvelope = strEnvelope & "<urn:timbrar_cfdi soapenv:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">"
    strEnvelope = strEnvelope & "<username xsi:type=""xsd:string"">" & strUserName & "</username>"
    strEnvelope = strEnvelope & "<password xsi:type=""xsd:string"">" & strPassword & "</password>"
    strEnvelope = strEnvelope & "<sxml xsi:type=""xsd:string"">" & sXml & "</sxml>"
    strEnvelope = strEnvelope & "</urn:timbrar_cfdi>"
    strEnvelope = strEnvelope & "</soapenv:Body>"
    strEnvelope = strEnvelope & "</soapenv:Envelope>"
    
    ' Llamar al servicio de timbox en el action timbrar_cfdi, con el envelope formado
    Timbrar = PostWebservice("timbrar_cfdi", strEnvelope)

End Function

Public Function Cancelar(ByVal UUID As String, ByVal RFC As String, ByVal pfxBase64 As String, ByVal pfxPassword As String) As String
    Dim strEnvelope As String
    
    ' Cuerpo de la peticion de cancelacion
    strEnvelope = ""
    strEnvelope = strEnvelope & "<soapenv:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:urn=""urn:WashOut"">"
    strEnvelope = strEnvelope & "<soapenv:Header/>"
    strEnvelope = strEnvelope & "<soapenv:Body>"
    strEnvelope = strEnvelope & "<urn:cancelar_cfdi soapenv:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">"
    strEnvelope = strEnvelope & "<username xsi:type=""xsd:string"">" & strUserName & "</username>"
    strEnvelope = strEnvelope & "<password xsi:type=""xsd:string"">" & strPassword & "</password>"
    strEnvelope = strEnvelope & "<rfcemisor xsi:type=""xsd:string"">" & RFC & "</rfcemisor>"
    strEnvelope = strEnvelope & "<uuids xsi:type=""urn:uuids"">"
    strEnvelope = strEnvelope & "<uuid xsi:type=""xsd:string"">" & UUID & "</uuid>"
    strEnvelope = strEnvelope & "</uuids>"
    strEnvelope = strEnvelope & "<pfxbase64 xsi:type=""xsd:string"">" & pfxBase64 & "</pfxbase64>"
    strEnvelope = strEnvelope & "<pfxpassword xsi:type=""xsd:string"">" & pfxPassword & "</pfxpassword>"
    strEnvelope = strEnvelope & "</urn:cancelar_cfdi>"
    strEnvelope = strEnvelope & "</soapenv:Body>"
    strEnvelope = strEnvelope & "</soapenv:Envelope>"
    
    ' Llamar al servicio de timbox en el action cancelar_cfdi, con el envelope formado
    Cancelar = PostWebservice("cancelar_cfdi", strEnvelope)
    
End Function

Public Function PostWebservice(ByVal SoapAction As String, ByVal XmlBody As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Load XML
    objDom.async = False
    objDom.LoadXml XmlBody

    ' Open the webservice
    objXmlHttp.open "POST", strWSDLUrl, False
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapAction
    
    ' Send XML command
    objXmlHttp.send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responseText
    
    ' Close object
    Set objXmlHttp = Nothing
    
    ' Return result
    PostWebservice = strRet
    
Exit Function
Err_PW:
    PostWebservice = "Error: " & Err.Number & " - " & Err.Description

End Function

Private Sub Class_Initialize()
    strWSDLUrl = "https://staging.ws.timbox.com.mx/timbrado_cfdi33/action"
    strUserName = "IAD121214B34"
    strPassword = "usar el password que se asigno"
End Sub

