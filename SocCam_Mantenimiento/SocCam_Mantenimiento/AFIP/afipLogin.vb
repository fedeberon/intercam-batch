Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Xml
Imports System.Net
Imports System.Security
Imports System.Security.Cryptography
' Importar system.security como referencia
Imports System.Security.Cryptography.Pkcs
Imports System.Security.Cryptography.X509Certificates
Imports System.IO
Imports System.Runtime.InteropServices




Public Class AfipLogin

    Private Shared _globalId As UInt32 = 0

    Public Property Serv As String
    Public Property Url As String
    Private Property Cert_path As String
    Private Property Clave As SecureString

    Private Property XmlLoginTicketRequest As XmlDocument
    Private Property XmlLoginTicketResponse As XmlDocument
    Private Property uniqueId As UInt32
    Public Property GenerationTime As DateTime
    Public Property ExpirationTime As DateTime

    Public Property certificado As X509Certificate2

    Public Property XDocRequest As XDocument
    Public Property XDocResponse As XDocument

    Public Property RawResponse As String

    Public ReadOnly Property Token As String
    Public ReadOnly Property Sign As String

    Public ReadOnly Property IsLogedIn As Boolean
        Get
            Return Not Token = ""
        End Get
    End Property

    ' Nueva instancia del objeto
    Public Sub New(serv As String, url As String)
        Me.Serv = serv
        Me.Url = url
    End Sub


    Public Function Login(ByRef cert As X509Certificate2, Optional ByRef loginError As String = "") As Boolean
        Dim log As New Log
        log.LogFilePath = Module1.LOG_DIR

        If Module1.LOG Then
            log.LogLevel = 2
        End If

        Dim cmsFirmadoBase64 As String
        Dim loginTicketResponse As String

        Dim uniqueIdNode As XmlNode
        Dim generationTimeNode As XmlNode
        Dim expirationTimeNode As XmlNode
        Dim serviceNode As XmlNode

        Try
            Me._globalId = CInt(Math.Ceiling(Rnd() * 2147483647)) + 1


            XmlLoginTicketRequest = New XmlDocument
            TemplateLoader.LoadTemplate(XmlLoginTicketRequest, "LoginTemplate")

            uniqueIdNode = XmlLoginTicketRequest.SelectSingleNode("//uniqueId")
            generationTimeNode = XmlLoginTicketRequest.SelectSingleNode("//generationTime")
            expirationTimeNode = XmlLoginTicketRequest.SelectSingleNode("//expirationTime")
            serviceNode = XmlLoginTicketRequest.SelectSingleNode("//service")
            generationTimeNode.InnerText = DateTime.Now.ToString("s")
            expirationTimeNode.InnerText = DateTime.Now.AddMinutes(+40).ToString("s")
            uniqueIdNode.InnerText = CStr(_globalId)
            serviceNode.InnerText = Serv

            Dim msgBytes As Byte() = Encoding.UTF8.GetBytes(XmlLoginTicketRequest.OuterXml)

            Dim infoContenido As New ContentInfo(msgBytes)
            Dim cmsFirmado As New SignedCms(infoContenido)

            Dim cmsFirmante As New CmsSigner(cert)
            cmsFirmante.IncludeOption = X509IncludeOption.EndCertOnly

            cmsFirmado.ComputeSignature(cmsFirmante)

            cmsFirmadoBase64 = Convert.ToBase64String(cmsFirmado.Encode())

            Dim servicio As New WSAA.LoginCMSService
            servicio.Url = Url

            loginTicketResponse = servicio.loginCms(cmsFirmadoBase64)

            XmlLoginTicketResponse = New XmlDocument
            XmlLoginTicketResponse.LoadXml(loginTicketResponse)

            Me.RawResponse = XmlLoginTicketResponse.InnerText

            _Token = XmlLoginTicketResponse.SelectSingleNode("//token").InnerText
            _Sign = XmlLoginTicketResponse.SelectSingleNode("//sign").InnerText

            GenerationTime = DateTime.Parse(generationTimeNode.InnerText)
            ExpirationTime = DateTime.Parse(expirationTimeNode.InnerText)

            XDocRequest = XDocument.Parse(XmlLoginTicketRequest.OuterXml)
            XDocResponse = XDocument.Parse(XmlLoginTicketResponse.OuterXml)

            Return True

        Catch ex As Exception
            Debug.Print(ex.Message)
            loginError = ex.Message
            log.SetError($"Facturacion [ FALLO ]: {ex.Message}", "afipLogin", "Login")
            Return False
        End Try
    End Function

End Class
