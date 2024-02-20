Public Class Afip

    ' URLs
    Public Const HOMOLOGAGION_AUTH_URL As String = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
    Public Const HOMOLOGAGION_INFO_URL As String = "https://awshomo.afip.gov.ar/sr-padron/webservices/personaServiceA5?WSDL"
    Public Const HOMOLOGAGION_FE_SOLICITAR_URL As String = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?op=FECAESolicitar"
    Public Const HOMOLOGAGION_FE_ULTIMO_NUMERO_URL As String = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?op=FECompUltimoAutorizado"

    Public Const PRODUCCION_AUTH_URL As String = "https://wsaa.afip.gov.ar/ws/services/LoginCms"
    Public Const PRODUCCION_INFO_URL As String = "https://aws.afip.gov.ar/sr-padron/webservices/personaServiceA5?WSDL"
    Public Const PRODUCCION_FE_SOLICITAR_URL As String = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL"
    Public Const PRODUCCION_FE_ULTIMO_NUMERO_URL As String = ""

    Public Const PRODUCCION_CONSTANCIA_INSCRIPCION_URL As String = "https://soa.afip.gob.ar/sr-padron/v1/constancia/CUIT"

    ' SERVICIOS
    Public Const SERVICIO_FE As String = "wsfe"
    Public Const SERVICIO_PADRON As String = "ws_sr_padron_a5"

    Private Property Produccion As Boolean = False

    Public Sub New()
    End Sub

    Public Sub New(ByVal gc As GlobalConfig)
        Produccion = gc.Produccion
    End Sub

    Public Sub New(ByVal homologacion As Boolean)
        Produccion = Not homologacion
    End Sub

    Public ReadOnly Property FE_SOLICITAR_URL() As String
        Get
            If Produccion Then
                Return PRODUCCION_FE_SOLICITAR_URL
            Else
                Return HOMOLOGAGION_FE_SOLICITAR_URL
            End If
        End Get
    End Property

    Public ReadOnly Property AUTH_URL() As String
        Get
            If Produccion Then
                Return PRODUCCION_AUTH_URL
            Else
                Return HOMOLOGAGION_AUTH_URL
            End If
        End Get
    End Property

    Public ReadOnly Property INFO_URL As String
        Get
            If Produccion Then
                Return PRODUCCION_INFO_URL
            Else
                Return HOMOLOGAGION_INFO_URL
            End If
        End Get
    End Property

    Public ReadOnly Property ULTIMO_NUMERO_URL As String
        Get
            If Produccion Then
                Return PRODUCCION_FE_ULTIMO_NUMERO_URL
            Else
                Return HOMOLOGAGION_FE_ULTIMO_NUMERO_URL
            End If
        End Get
    End Property

    Public ReadOnly Property Homologacion As Boolean
        Get
            Return Not Produccion
        End Get
    End Property

    Public Function VerificarEstadoServicioPadron(ByRef errorMsg As String) As Boolean
        Dim personaService As New WSPSA5.PersonaServiceA5
        If Me.Produccion Then
            personaService.Url = PRODUCCION_INFO_URL
        Else
            personaService.Url = HOMOLOGAGION_INFO_URL
        End If

        Dim cuitData As New WSPSA5.personaReturn
        Try
            Dim serviceStatus As New WSPSA5.dummyReturn
            serviceStatus = personaService.dummy()
            If serviceStatus.appserver <> "OK" Then
                errorMsg = "Servidor de Padron A5 AFIP caido"
                Return False
            End If
            If serviceStatus.authserver <> "OK" Then
                errorMsg = "Servidor de autorización AFIP caido"
                Return False
            End If
            If serviceStatus.dbserver <> "OK" Then
                errorMsg = "Servidor de base de datos del Padron A5 AFIP caido"
                Return False
            End If

            Return True
        Catch ex As Exception
            Debug.Print(ex.Message)
            errorMsg = ex.Message
            Return False
        End Try
    End Function


End Class
