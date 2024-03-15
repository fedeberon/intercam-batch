Imports helix

Module Module1

    Public Const ERR_SUCCESS As Integer = 0
    Public Const ERR_MISSING_ARGS As Integer = 1

    Private silentFlag As Boolean = False

    Public sqle As New SQLEngine

    Dim ConsoleOut As New ConsoleOut

    Public Property LOG As Boolean = False
    Public Property LOG_DIR As String = ""


    Sub Main(ByVal sArgs() As String)
        Dim exitCode As Integer = ERR_SUCCESS

        Dim argParse As New ArgsParser
        'If My.Computer.Name = "DESKTOP" Then
        'Dim dbg As New List(Of String)
        'dbg.Add("--crear-cuotas-sociales")
        'dbg.Add("homologacion")
        'dbg.Add($"ccb")
        'dbg.Add($"p{2}")
        'dbg.Add($"a{2019}")
        'dbg.Add($"--log")
        'dbg.Add($"C:\Users\logico\Downloads\TEST")
        '
        'argParse.SW_DEBUG = True
        'argParse.ParseArguments(My.Application.CommandLineArgs, dbg)
        'Else
        'argParse.ParseArguments(My.Application.CommandLineArgs)
        'End If
        argParse.ParseArguments(My.Application.CommandLineArgs)

        ' ==============================================================
        ' Zona de 1 solo argumento

        If argParse.SW_ERROR Then
            Console.WriteLine(argParse.ERROR_LIST)
            exitCode = ERR_MISSING_ARGS
            End
        End If

        If argParse.SW_HELP Then
            PrintUsage()
            Environment.ExitCode = exitCode
            End
        End If

        If argParse.SW_VERSION Then
            PrintVersion()
            Environment.ExitCode = exitCode
            End
        End If
        ' ==============================================================




        Environment.ExitCode = Execute(argParse)
    End Sub

    ''' <summary>
    ''' Ejecuta la tarea
    ''' </summary>
    ''' <param name="argParse">Listado de parametros a ejecutar</param>
    ''' <returns>Codigo de salida del programa</returns>
    Public Function Execute(ByVal argParse As ArgsParser) As Integer
        Dim ejecutor As New Executor
        ejecutor.Silent = silentFlag

        LOG = argParse.SW_LOG
        LOG_DIR = argParse.PARAM_LOGFILE

        If argParse.SW_ACTUALIZAR_PADRON_AFIP Then
            Return ejecutor.ActualizarPadronAfip(argParse.PARAM_PADRON_URL)
        End If

        If argParse.SW_CREAR_CUOTAS_SOCIALES Then
            If argParse.SW_NO_FACTURAR Then
                'No genera el movimiento ni la factura. NO SE USA MAS.
                Return ejecutor.GenerarCuotasSocios(argParse.PARAM_PERIODO_CUOTA_SOCIAL - 1, argParse.PARAM_ANIO_CUOTA_SOCIAL, argParse.SW_CUOTA_MES_VENCIDO, argParse.SW_CUOTA_AUTOCOBRAR, argParse.SW_CUOTA_OMITIR_SOCIOS_COFRE)
            Else
                'Se genera cuota, recibo y movimiento. No se factura.

                'Camara: envia mail. NO ES MAS NECESARIO. Solo se enviaban cuando se generaban las facturas antes. Ahora solo se genera recibo.
                'Return ejecutor.GenerarCuotasSocios(argParse.PARAM_PERIODO_CUOTA_SOCIAL - 1, argParse.PARAM_ANIO_CUOTA_SOCIAL, argParse.SW_CUOTA_OMITIR_SOCIOS_COFRE, argParse.SW_SENDMAIL)

                'DB Local: no enviar mail.
                Return ejecutor.GenerarCuotasSocios(argParse.PARAM_PERIODO_CUOTA_SOCIAL, argParse.PARAM_ANIO_CUOTA_SOCIAL, argParse.SW_CUOTA_OMITIR_SOCIOS_COFRE)
            End If
        End If
    End Function


    Public Sub PrintUsage()
        Console.WriteLine("SocCam Manteninimento Ver.: " & My.Application.Info.Version.ToString)
        Console.WriteLine("Programa para el mantenimiento interno del sistema SocCam")
        Console.WriteLine("Uso: soccam_mantenimiento.exe --actualizar-padron")
        Console.WriteLine("")
        Console.WriteLine("Listado de opciones.")
        Console.WriteLine($"--silent, -s  {vbTab} Modo silencioso, no muestra ningun mensaje en consola")
        Console.WriteLine($"--version, -v {vbTab} Versión del programa")
        Console.WriteLine($"--actualizar-padron, -p [URL]... {vbTab} Actualiza la base de datos del padrón de AFIP")
        Console.WriteLine($"--crear-cuotas-sociales, -ccs [OPCIONES]... {vbTab} p[MES] a[ANIO]{vbTab}Genera cuotas sociales")
    End Sub

    Public Sub PrintVersion()
        Console.WriteLine("SocCam Mantenimiento Ver.: " & My.Application.Info.Version.ToString)
        Console.WriteLine("Programa para el mantenimiento interno del sistema SocCam")
        Console.WriteLine("Intentar con 'soccam_mantenimiento --help' para mas info")
    End Sub

    Public Sub PrintError(ByVal errorCode As Integer)
        Select Case errorCode
            Case 1
                Console.WriteLine("DBackup: Faltan argumentos")
            Case Else
                Console.WriteLine("DBackup: Error desconocido " & errorCode)
        End Select
    End Sub

End Module
