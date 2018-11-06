VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnvioSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private usuario As String
Private clave As String
Private SoapClient As MSSOAPLib30.SoapClient30
Private WSDLPath As String
Private Proveedor As String

Public Sub Inicializar(strProveedor As String)
    Proveedor = strProveedor
    If Proveedor = "publimes" Then
        usuario = "109"
        clave = "IRB5639"
        Set SoapClient = New MSSOAPLib30.SoapClient30
        WSDLPath = "http://online.publimes.com:5000/Service.svc?wsdl"
        Call SoapClient.MSSoapInit(WSDLPath)
    ElseIf Proveedor = "smsmasivo" Then
        usuario = "rb"
        clave = "importadores"
        WSDLPath = "http://http.easyphone-ip.com:80/SmsMasivoWS?wsdl"
    ElseIf Proveedor = "aldeamo" Then
        usuario = "RBIntegracion"
        clave = "RBIntegracion123*"
        WSDLPath = "https://apismsi.aldeamo.com/sms/sms.wsdl"
        
        
    End If
End Sub

Public Sub Enviar(strMensaje As String, strCelular As String, Optional strOperadora As String = "C")
    Dim idcliente As String
    Dim idclave As String
    Dim idoperadora As String
    Dim idcelular As String
    Dim idmensaje As String
    Dim respuesta As Integer
    
    Dim sEnv As String
    Dim xmlhtp As New MSXML2.XMLHTTP40
    Dim xmlDoc As New DOMDocument
    
    If Len(strCelular) = 10 And strMensaje <> "" Then
        idcliente = usuario
        idclave = clave
        idoperadora = strOperadora
        idcelular = strCelular
        idmensaje = strMensaje
        If Proveedor = "publimes" Then
            respuesta = SoapClient.EnviarMensaje(idcliente, idclave, idoperadora, idcelular, idmensaje)
        ElseIf Proveedor = "smsmasivo" Then
            sEnv = "<?xml version=""1.0"" encoding=""utf-8""?>"
            sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ws=""http://ws.smsmasivo.synergy.com/"">"
            sEnv = sEnv & " <soapenv:Header/>"
            sEnv = sEnv & " <soapenv:Body>"
            sEnv = sEnv & " <ws:sendSms>"
            sEnv = sEnv & " <user>" & idcliente & "</user>"
            sEnv = sEnv & " <password>" & idclave & "</password>"
            sEnv = sEnv & " <message>" & idmensaje & "</message>"
            sEnv = sEnv & " <sender>JSN</sender>"
            sEnv = sEnv & " <number>+593" & IIf(Left(idcelular, 1) = "0", Right(idcelular, Len(idcelular) - 1), idcelular) & "</number>"
            sEnv = sEnv & " </ws:sendSms>"
            sEnv = sEnv & " </soapenv:Body>"
            sEnv = sEnv & "</soapenv:Envelope>"

            With xmlhtp
                .Open "post", WSDLPath, False
                .setRequestHeader "Host", "http.easyphone-ip.com"
                .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
                .setRequestHeader "soapAction", "http://http.easyphone-ip.com/SmsMasivoWS?wsdl"
                .send sEnv
                xmlDoc.loadXML .responseText
                If xmlDoc.childNodes(1).childNodes(0).childNodes(0).childNodes(0).childNodes(1).childNodes(0).Text = "Failed" Then
                    MsgBox "No se envio el SMS." & vbNewLine & xmlDoc.childNodes(1).childNodes(0).childNodes(0).childNodes(0).childNodes(0).childNodes(0).Text
                End If
            End With
        ElseIf Proveedor = "aldeamo" Then
            sEnv = "<?xml version=""1.0"" encoding=""utf-8""?>"
            sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:hab=""http://habitatclient.service.soapws.core.com/"">"
            sEnv = sEnv & " <soapenv:Header/>"
            sEnv = sEnv & " <soapenv:Body>"
            sEnv = sEnv & " <ws:smsSendSoap>"
            sEnv = sEnv & " <username>" & idcliente & "</username>"
            sEnv = sEnv & " <password>" & idclave & "</password>"
            sEnv = sEnv & " <country>593</country>"
            sEnv = sEnv & " <mobile>" & IIf(Left(idcelular, 1) = "0", Right(idcelular, Len(idcelular) - 1), idcelular) & "</mobile>"
            sEnv = sEnv & " <message>" & idmensaje & "</message>"
            sEnv = sEnv & " </hab:smsSendSoap>"
            sEnv = sEnv & " </soapenv:Body>"
            sEnv = sEnv & "</soapenv:Envelope>"

            With xmlhtp
                .Open "post", WSDLPath, False
                .setRequestHeader "Host", "apismsi.aldeamo.com"
                .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
                .setRequestHeader "soapAction", "http://apismsi.aldeamo.com/sms/sms.wsdl"
                .send sEnv
                xmlDoc.loadXML .responseText
                'MsgBox .responseText
            End With
            
        End If
                        
    End If
End Sub
