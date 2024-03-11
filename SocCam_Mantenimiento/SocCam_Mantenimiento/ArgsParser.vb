Imports System.Collections.ObjectModel

Public Class ArgsParser

    Public Property SW_SILENT As Boolean = False
    Public Property SW_HELP As Boolean = False
    Public Property SW_DEBUG As Boolean = False
    Public Property SW_VERSION As Boolean = False
    Public Property SW_CREAR_CUOTAS_SOCIALES As Boolean = False
    Public Property SW_ACTUALIZAR_VENCIMIENTOS_CUOTAS As Boolean = False
    Public Property SW_ACTUALIZAR_PADRON_AFIP As Boolean = False
    Public Property SW_SYNC_CUOTA_FACTURA As Boolean = False
    Public Property SW_NO_FACTURAR As Boolean = False
    Public Property SW_CUOTA_MES_VENCIDO As Boolean = False
    Public Property SW_CUOTA_AUTOCOBRAR As Boolean = False
    Public Property SW_CUOTA_OMITIR_SOCIOS_COFRE As Boolean = False

    Public Property SW_CUOTA_EXTRA As Boolean = False
    Public Property SW_CUOTA_EXTRA_IMPORTE As Decimal = 0
    Public Property SW_CUOTA_EXTRA_SECTORES As String = ""

    Public Property SW_ERROR As Boolean = True
    Public Property SW_SENDMAIL As Boolean = False
    Public Property ERROR_LIST As String = "Faltan argumentos"

    Public Property SW_LOG As Boolean = False

    Public Property PARAM_PADRON_URL As String = ""
    Public Property PARAM_HOMOLOGACION As Boolean = True
    Public Property PARAM_PERIODO_CUOTA_SOCIAL As Integer = 0
    Public Property PARAM_ANIO_CUOTA_SOCIAL As Integer = 0
    Public Property PARAM_LOGFILE As String = ""

    Public Overrides Function ToString() As String
        Dim res As String = $"SW_SILENT = {SW_SILENT}
        SW_HELP = {SW_HELP}
        SW_DEBUG = {SW_DEBUG}
        SW_VERSION = {SW_VERSION}
        SW_CREAR_CUOTAS_SOCIALES = {SW_CREAR_CUOTAS_SOCIALES}
        SW_ACTUALIZAR_VENCIMIENTOS_CUOTAS = {SW_ACTUALIZAR_VENCIMIENTOS_CUOTAS}
        SW_ACTUALIZAR_PADRON_AFIP = {SW_ACTUALIZAR_PADRON_AFIP}
        SW_SYNC_CUOTA_FACTURA = {SW_SYNC_CUOTA_FACTURA}
        SW_NO_FACTURAR = {SW_NO_FACTURAR}
        SW_CUOTA_MES_VENCIDO = {SW_CUOTA_MES_VENCIDO}
        SW_CUOTA_AUTOCOBRAR = {SW_CUOTA_AUTOCOBRAR}
        SW_CUOTA_OMITIR_SOCIOS_COFRE = {SW_CUOTA_OMITIR_SOCIOS_COFRE}
        SW_CUOTA_EXTRA = {SW_CUOTA_EXTRA}
        SW_CUOTA_EXTRA_IMPORTE = {SW_CUOTA_EXTRA_IMPORTE}
        SW_CUOTA_EXTRA_SECTORES = {SW_CUOTA_EXTRA_SECTORES}
        SW_ERROR = {SW_ERROR}
        SW_SENDMAIL = {SW_SENDMAIL}
        SW_LOG = {SW_LOG}
        ERROR_LIST = {ERROR_LIST}
        PARAM_PADRON_URL = {PARAM_PADRON_URL}
        PARAM_HOMOLOGACION = {PARAM_HOMOLOGACION}
        PARAM_PERIODO_CUOTA_SOCIAL = {PARAM_PERIODO_CUOTA_SOCIAL}
        PARAM_ANIO_CUOTA_SOCIAL = {PARAM_ANIO_CUOTA_SOCIAL}
        PARAM_ANIO_CUOTA_SOCIAL = {PARAM_LOGFILE}"

        Return res
    End Function

    Public Sub ParseArguments(ByVal args As ReadOnlyCollection(Of String), Optional ByVal argsD As List(Of String) = Nothing)
        Dim waitForParam As Integer = 0

        ' DEBUG MODE
        If Not IsNothing(argsD) Then
            For Each arg As String In argsD
                If waitForParam <> 0 Then
                    Select Case waitForParam
                        Case 1
                            If arg.StartsWith("http") Then
                                PARAM_PADRON_URL = arg
                                SW_ERROR = False
                                waitForParam = 0
                            Else
                                SW_ERROR = True
                                Exit Sub
                            End If

                        Case 2
                            'Se comentaron las líneas de SW_CUOTA_EXTRA para evitar cobro extra.
                            'Anteriormente la camara enviaba cobradores por zona a los socios y eso era un cargo extra.
                            Select Case arg.ToLower
                                Case "ccb"
                                    'SW_CUOTA_EXTRA_IMPORTE = 60
                                    'SW_CUOTA_EXTRA_SECTORES = "1,2,3,4,5,6,7"
                                    'SW_CUOTA_EXTRA = True
                                    SW_CUOTA_OMITIR_SOCIOS_COFRE = True
                                    'SW_SENDMAIL = True
                                Case "autocobrar", "ac"
                                    SW_CUOTA_AUTOCOBRAR = True
                                    SW_ERROR = False
                                Case "no-facturar", "nf"
                                    SW_NO_FACTURAR = True
                                    SW_ERROR = False
                                Case "cuota-omitir-cofres"
                                    SW_CUOTA_OMITIR_SOCIOS_COFRE = True
                                    SW_ERROR = False
                                Case "homologacion", "h"
                                    If SW_NO_FACTURAR Then
                                        SW_ERROR = True
                                        Exit Sub
                                    Else
                                        PARAM_HOMOLOGACION = True
                                        SW_ERROR = False
                                    End If
                                Case "produccion", "p"
                                    If SW_NO_FACTURAR Then
                                        SW_ERROR = True
                                        Exit Sub
                                    Else
                                        PARAM_HOMOLOGACION = False
                                        SW_ERROR = False
                                    End If
                                Case "mes-vencido", "mv"
                                    SW_CUOTA_MES_VENCIDO = True
                                    SW_ERROR = False
                                Case "p1", "p2", "p3", "p4", "p5", "p6", "p7", "p8", "p9", "p10", "p11", "p12"
                                    PARAM_PERIODO_CUOTA_SOCIAL = CInt(arg.ToLower.Replace("p", ""))
                                    SW_ERROR = False
                                Case Else
                                    If arg.StartsWith("a") Then
                                        If IsNumeric(arg.ToLower.Replace("a", "")) Then
                                            PARAM_ANIO_CUOTA_SOCIAL = CInt(arg.ToLower.Replace("a", ""))
                                            SW_ERROR = False
                                            waitForParam = 0
                                        Else
                                            SW_ERROR = True
                                            Exit Sub
                                        End If
                                    Else
                                        SW_ERROR = True
                                        Exit Sub
                                    End If
                            End Select
                        Case 3
                            If My.Computer.FileSystem.DirectoryExists(arg) Then
                                PARAM_LOGFILE = arg
                                SW_ERROR = False
                                waitForParam = 0
                            Else
                                SW_ERROR = True
                                Exit Sub
                            End If
                    End Select
                Else
                    Select Case arg.ToLower
                        Case "--help", "-h"
                            SW_HELP = True
                            SW_ERROR = False
                        Case "--version", "-v"
                            SW_VERSION = True
                            SW_ERROR = False
                        Case "--silent", "-s"
                            SW_SILENT = True
                            SW_ERROR = False
                        Case "--actualizar-padron", "-p"
                            SW_ACTUALIZAR_PADRON_AFIP = True
                            ERROR_LIST = "Falta URL del padrón de AFIP"
                            waitForParam = 1
                        Case "--crear-cuotas-sociales", "-ccs"
                            SW_CREAR_CUOTAS_SOCIALES = True
                            ERROR_LIST = "Falta parametro cuotas sociales: HOMOLOGACION"
                            waitForParam = 0
                        Case "--sync-cuota-factura"
                            SW_SYNC_CUOTA_FACTURA = True
                        Case "--log", "-l"
                            SW_LOG = True
                            ERROR_LIST = "Falta ubicación del archivo de log"
                            waitForParam = 4
                    End Select
                End If
            Next

            If SW_DEBUG Then
                Exit Sub
            End If
        End If

        ' PRODUCTION MODE
        If args.Count = 0 Then
            SW_HELP = True
            SW_ERROR = False
            Exit Sub
        End If

        For Each arg As String In args
            If waitForParam <> 0 Then
                Select Case waitForParam
                    Case 1
                        If arg.StartsWith("http") Then
                            PARAM_PADRON_URL = arg
                            SW_ERROR = False
                            waitForParam = 0
                        Else
                            SW_ERROR = True
                            Exit Sub
                        End If

                    Case 2
                        'Se comentaron las líneas de SW_CUOTA_EXTRA para evitar cobro extra.
                        'Anteriormente la camara enviaba cobradores por zona a los socios y eso era un cargo extra.
                        Select Case arg.ToLower
                            Case "ccb"
                                'SW_CUOTA_EXTRA_IMPORTE = 60
                                'SW_CUOTA_EXTRA_SECTORES = "1,2,3,4,5,6,7"
                                'SW_CUOTA_EXTRA = True
                                SW_CUOTA_OMITIR_SOCIOS_COFRE = True
                                'SW_SENDMAIL = True
                            Case "autocobrar", "ac"
                                SW_CUOTA_AUTOCOBRAR = True
                                SW_ERROR = False
                            Case "no-facturar", "nf"
                                SW_NO_FACTURAR = True
                                SW_ERROR = False
                            Case "cuota-omitir-cofres"
                                SW_CUOTA_OMITIR_SOCIOS_COFRE = True
                                SW_ERROR = False
                            Case "homologacion", "h"
                                PARAM_HOMOLOGACION = True
                                SW_ERROR = False
                            Case "produccion", "p"
                                PARAM_HOMOLOGACION = False
                                SW_ERROR = False
                            Case "p1", "p2", "p3", "p4", "p5", "p6", "p7", "p8", "p9", "p10", "p11", "p12"
                                PARAM_PERIODO_CUOTA_SOCIAL = CInt(arg.ToLower.Replace("p", ""))
                                SW_ERROR = False
                            Case Else
                                If arg.StartsWith("a") Then
                                    If IsNumeric(arg.ToLower.Replace("a", "")) Then
                                        PARAM_ANIO_CUOTA_SOCIAL = CInt(arg.ToLower.Replace("a", ""))
                                        SW_ERROR = False
                                        waitForParam = 0
                                    Else
                                        SW_ERROR = True
                                        Exit Sub
                                    End If
                                Else
                                    SW_ERROR = True
                                    Exit Sub
                                End If
                        End Select
                    Case 3
                        If My.Computer.FileSystem.DirectoryExists(arg) Then
                            PARAM_LOGFILE = arg
                            SW_ERROR = False
                            waitForParam = 0
                        Else
                            SW_ERROR = True
                            Exit Sub
                        End If
                End Select
            Else
                Select Case arg.ToLower
                    Case "--help", "-h"
                        SW_HELP = True
                        SW_ERROR = False
                    Case "--version", "-v"
                        SW_VERSION = True
                        SW_ERROR = False
                    Case "--silent", "-s"
                        SW_SILENT = True
                        SW_ERROR = False
                    Case "--actualizar-padron", "-p"
                        SW_ACTUALIZAR_PADRON_AFIP = True
                        ERROR_LIST = "Falta URL del padrón de AFIP"
                        waitForParam = 1
                    Case "--crear-cuotas-sociales", "-ccs"
                        SW_CREAR_CUOTAS_SOCIALES = True
                        ERROR_LIST = "Falta parametro cuotas sociales: HOMOLOGACION"
                        waitForParam = 2
                    Case "--log", "-l"
                        SW_LOG = True
                        ERROR_LIST = "Falta ubicación del archivo de log"
                        waitForParam = 3
                End Select
            End If
        Next

    End Sub

End Class
