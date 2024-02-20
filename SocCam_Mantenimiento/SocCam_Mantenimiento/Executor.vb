Imports System.IO
Imports System.IO.Compression
Imports System.Windows.Shapes
Imports helix
Imports SelectPdf

Public Class Executor

    Public Property Silent As Boolean = False

    Public Property SqleGlobal As New SQLEngine

    Public Sub New()
        SqleGlobal.dbType = helix.SQLEngine.dataBaseType.SQL_SERVER
        SqleGlobal.DatabaseName = "soccam_test" 'Cambiar con el nombre de la DB correspondiente.
        SqleGlobal.RequireCredentials = False
        'SqleGlobal.Username = ""
        'SqleGlobal.Password = 

        'Si la DB esta en tu equipo, no es necesario cambiar esta linea.
        SqleGlobal.Path = My.Computer.Name & "\" & "SQLEXPRESS"

        SqleGlobal.Start()
    End Sub

    Public Function ActualizarPadronAfip(ByVal url As String) As Boolean
        Dim filePath = My.Computer.FileSystem.SpecialDirectories.Temp
        If Not Utils.DescargarArchivo(url, filePath, Silent) Then
            Return False
        End If
        Dim objProcess As System.Diagnostics.Process

        Dim ConsoleOut As New ConsoleOut

        Try

            If Not Silent Then
                ConsoleOut.Print($"- Descomprimiendo padron...")
            End If
            objProcess = New System.Diagnostics.Process()
            objProcess.StartInfo.UseShellExecute = True
            objProcess.StartInfo.FileName = $"{My.Application.Info.DirectoryPath}\7za.exe"
            objProcess.StartInfo.Arguments = $"x {filePath}{url.Substring(url.LastIndexOf("/"), (url.Length - url.LastIndexOf("/")))} -y"
            objProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
            objProcess.Start()

            objProcess.WaitForExit()
            objProcess.Close()

        Catch ex As Exception
            Debug.Print(ex.Message)
            If Not Silent Then
                ConsoleOut.Print($"- FALLO: {ex.Message} [FAIL]")
            End If
            Return False
        End Try


        If Not Silent Then
            ConsoleOut.Print($"- Descomprimir [OK]")
        End If

        filePath = $"{My.Application.Info.DirectoryPath}\utlfile\padr\"

        For Each foundFile As String In My.Computer.FileSystem.GetFiles(filePath,
    Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.tmp")
            Dim sPadron As New SQLEngine
            sPadron.dbType = helix.SQLEngine.dataBaseType.SQL_SERVER
            sPadron.DatabaseName = "soccam_test" 'Cambiar con el nombre de la DB correspondiente.
            sPadron.RequireCredentials = False
            'sPadron.Username = ""
            'sPadron.Password = 

            'Si la DB esta en tu equipo, no es necesario cambiar esta linea.
            sPadron.Path = My.Computer.Name & "\" & "SQLEXPRESS"

            If Not sPadron.Start Then
                If Not Silent Then
                    ConsoleOut.Print($"- Actualizar padron: No se pudo conectar a la base de datos [FAIL]")
                End If
                Return False
            End If
            Dim tst As New AfipCondicionFiscal(sPadron)

            If tst.DeleteAll() Then
                If Not Silent Then
                    ConsoleOut.Print($"- Procesando padron...")
                End If
                tst.ImportarPadronDB(foundFile)
            Else
                If Not Silent Then
                    ConsoleOut.Print($"- Actualizar padron: No se pudo conectar a la base de datos [FAIL]")
                End If
                Return False
            End If
        Next

        Return True
    End Function

    Public Function GenerarCuotasSocios(ByVal mes As Integer, ByVal anio As Integer, ByVal mesVencido As Boolean, ByVal autocobrar As Boolean, ByVal omitirUsuariosCofres As Boolean) As Boolean
        Dim socios As New Socio
        Dim dtSocios As New DataTable
        Dim dtrSocios As DataTableReader
        Dim ConsoleOut As New ConsoleOut
        Dim sGC As New SQLEngine

        Dim sSocios As New SQLEngine
        Dim sTipo As New SQLEngine
        Dim sCuota As New SQLEngine
        Dim sSocio As New SQLEngine

        sSocios.dbType = helix.SQLEngine.dataBaseType.SQL_SERVER
        sSocios.DatabaseName = "soccam_test"
        sSocios.RequireCredentials = False
        'sSocios.Username = ""
        'sSocios.Password = 

        'Si la DB esta en tu equipo, no es necesario cambiar esta linea.
        sSocios.Path = My.Computer.Name & "\" & "SQLEXPRESS"

        ConsoleOut.Print($"sSocios.Path: {sSocios.Path}")
        If Not sSocios.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            Return False
        End If

        sTipo = sSocios
        If Not sTipo.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            Return False
        End If

        sCuota = sSocios
        If Not sCuota.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            Return False
        End If

        sSocio = sSocios
        If Not sSocio.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            Return False
        End If

        ' Buscar todos los socios activos
        If mesVencido Then
            If mes = 0 Then
                mes = 11
                anio -= 1

            Else
                mes -= 1
            End If
        End If

        If socios.LoadAll(sSocios, dtSocios, True) Then
            dtrSocios = dtSocios.CreateDataReader

            If Not Silent Then
                ConsoleOut.Print($"- Generando cuotas sociales")
            End If

            Dim totalSocios As Integer = dtSocios.Rows.Count - 1
            Dim currentProcess As Integer = 1

            While dtrSocios.Read

                Dim currSocio As New Socio
                currSocio.LoadMe(sSocio, dtrSocios(TABLA_SOCIO.ID))

                If Not Silent Then
                    ConsoleOut.Print($"{ConsoleOut.ProgressBarStep} {currentProcess}/{totalSocios} - {currSocio.Nombre.Trim.PadRight(80, " ")}")
                End If

                Dim plan As New SocioTipo
                plan.sqle = sTipo
                plan.LoadMe(dtrSocios(TABLA_SOCIO.TIPO_SOCIO))

                Dim periodo As Integer = GetPeriodoFromFecha(mes, plan.getMesesPorPeriodo)

                Dim c As New CuotaSocio
                c.sqle = sTipo

                If Not c.CuotaExist(sCuota, periodo, plan.periodicidad, anio, dtrSocios(TABLA_SOCIO.ID)) Then

                    If ((currSocio.FechaAceptacion.Month - 1) = periodo) And (currSocio.FechaAceptacion.Year = anio) Then
                        If currSocio.FechaAceptacion.Day <= 10 Then
                            Continue While
                        End If
                    End If

                    c.anio = anio
                    c.Periodo = periodo
                    c.Periodicidad = plan.periodicidad
                    c.PlanID = plan.id
                    c.monto = plan.importe

                    Dim cobrador As New Cobrador
                    cobrador.LoadMe(sTipo, dtrSocios(TABLA_SOCIO.SECTOR), True)
                    c.cobradorID = cobrador.ID
                    c.estado = CuotaSocio.ESTADO_SOCIO.PENDIENTE
                    c.socioID = dtrSocios(TABLA_SOCIO.ID)
                    If autocobrar And c.monto = 0 Then
                        c.estado = CuotaSocio.ESTADO_SOCIO.AL_DIA
                        c.fechaPago = Now
                    End If
                    c.Save(sCuota, 0)
                End If

                currentProcess += 1
            End While

            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ OK ]")
            End If

            Return True
        End If
        Return False
    End Function

    Public Function GenerarCuotasSocios(ByVal mes As Integer,
                                        ByVal anio As Integer,
                                        ByVal homologacion As Boolean,
                                        ByVal extraCuota As Decimal,
                                        ByVal extraSectores As String,
                                        ByVal omitirUsuariosCofres As Boolean,
                                        ByVal enviarFacturaPorMail As Boolean) As Boolean
        Dim socios As New Socio
        Dim dtSocios As New DataTable
        Dim dtrSocios As DataTableReader
        Dim ConsoleOut As New ConsoleOut
        Dim sGC As New SQLEngine

        Dim sSocios As New SQLEngine
        Dim sTipo As New SQLEngine
        Dim sCuota As New SQLEngine
        Dim sSocio As New SQLEngine
        Dim sCofre As New SQLEngine

        Dim log As New Log
        log.LogFilePath = Module1.LOG_DIR

        If Module1.LOG Then
            log.LogLevel = 2
        End If

        sSocios.dbType = helix.SQLEngine.dataBaseType.SQL_SERVER
        sSocios.DatabaseName = "soccam_test" 'Cambiar con el nombre de la DB correspondiente.
        sSocios.RequireCredentials = False
        'sSocios.Username = ""
        'sSocios.Password =

        'Si la DB esta en tu maquina, no es necesario cambiar esta linea.
        sSocios.Path = My.Computer.Name & "\" & "SQLEXPRESS"

        'If My.Computer.Name = "ALPHA" Then
        '    sSocios.Path = My.Computer.Name
        '    sSocios.DatabaseName = "soccam_ccb"
        '    Silent = True

        'End If

        If Not sSocios.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            log.SetError("Generacion de cuotas sociales [ FALLO ]: No se pudo iniciar el motor SQL socios", "Executor", "sSocios.Start")
            Return False
        End If

        sTipo = sSocios
        If Not sTipo.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            log.SetError("Generacion de cuotas sociales [ FALLO ]: No se pudo iniciar el motor SQL tipo socios", "Executor", "sTipo.Start")
            Return False
        End If

        sCuota = sSocios
        If Not sCuota.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            log.SetError("Generacion de cuotas sociales [ FALLO ]: No se pudo iniciar el motor SQL cuotas socios", "Executor", "sCuota.Start")
            Return False
        End If

        sSocio = sSocios
        If Not sSocio.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            log.SetError("Generacion de cuotas sociales [ FALLO ]: No se pudo iniciar el motor SQL socio", "Executor", "sSocio.Start")
            Return False
        End If

        sCofre = sSocios
        If Not sCofre.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            log.SetError("Generacion de cuotas sociales [ FALLO ]: No se pudo iniciar el motor SQL cofres", "Executor", "sCofre.Start")
            Return False
        End If

        ' Buscar todos los socios activos
        If socios.LoadAll(sSocios, dtSocios, True) Then
            dtrSocios = dtSocios.CreateDataReader

            Dim totalSocios As Integer = dtSocios.Rows.Count - 1

            If Not Silent Then
                ConsoleOut.Print($"- Generando cuotas sociales")
            End If

            Dim currentProcess As Integer = 1
            While dtrSocios.Read
                Dim currSocio As New Socio
                currSocio.LoadMe(sSocio, dtrSocios(TABLA_SOCIO.ID))


                If Not Silent Then
                    ConsoleOut.UpdateLastLine($"{ConsoleOut.ProgressBarStep} {currentProcess}/{totalSocios} - {currSocio.Nombre.Trim.PadRight(80, " ")}")
                Else
                    Debug.Print($"{ConsoleOut.ProgressBarStep} {currentProcess}/{totalSocios} - {currSocio.Nombre.Trim.PadRight(80, " ")}")
                End If

                If omitirUsuariosCofres Then
                    Dim tieneCofre As Boolean = False
                    Dim contratoCofre As New ContratoCofre(sCofre)
                    contratoCofre.QuickSearch(ContratoCofre.TABLA.ES_SOCIO_ID, SQLEngineQuery.OperatorCriteria.Igual, currSocio.InternalID)

                    If contratoCofre.SearchResult.Count > 0 Then
                        For Each contrato As ContratoCofre In contratoCofre.SearchResult
                            If Not contrato.Deleted Then
                                tieneCofre = True
                                Exit For
                            End If
                        Next

                    End If


                    If tieneCofre Then
                        currentProcess += 1
                        Continue While
                    End If
                End If

                Dim plan As New SocioTipo
                plan.sqle = sTipo
                plan.LoadMe(dtrSocios(TABLA_SOCIO.TIPO_SOCIO))

                Dim c As New CuotaSocio
                c.sqle = sTipo

                Dim periodo As Integer = GetPeriodoFromFecha(mes, plan.getMesesPorPeriodo)

                If Not c.CuotaExist(sCuota, periodo, plan.periodicidad, anio, dtrSocios(TABLA_SOCIO.ID)) Then

                    If ((currSocio.FechaAceptacion.Month - 1) = periodo) And (currSocio.FechaAceptacion.Year = anio) Then
                        If currSocio.FechaAceptacion.Day <= 10 Then
                            Continue While
                        End If
                    End If

                    c.anio = anio
                    c.Periodo = periodo
                    c.Periodicidad = plan.periodicidad

                    c.PlanID = plan.id



                    Dim cobrador As New Cobrador
                    cobrador.LoadMe(sTipo, dtrSocios(TABLA_SOCIO.SECTOR), True)

                    If currSocio.Sector >= 1 And currSocio.Sector <= 7 Then
                        c.monto = plan.importe + cobrador.ComisionFija
                    Else
                        c.monto = plan.importe
                    End If

                    'If My.Computer.Name = "ALPHA" And System.Diagnostics.Debugger.IsAttached Then Continue While

                    c.cobradorID = cobrador.ID
                    c.estado = CuotaSocio.ESTADO_SOCIO.PENDIENTE
                    c.socioID = dtrSocios(TABLA_SOCIO.ID)

                    ' GENERAR FACTURA AQUI
                    If c.monto = 0 Then
                        Debug.Print($"{currSocio.Apellido} {currSocio.Nombre}")
                        c.Delete(sCuota, c.id)
                        Continue While
                    End If
                    c.Save(sCuota, 0)

                    Dim facturaID As Integer
                    Dim feNum As Integer = GenerarFE(mes, anio, currSocio, homologacion, c.monto, c.id, enviarFacturaPorMail, facturaID)
                    If feNum <> 0 Then
                        Dim mov As New MovimientoCuentaCorrienteSocio(Me.SqleGlobal)
                        mov.ClienteId = c.socioID
                        mov.ComprobanteRelacionado = facturaID
                        mov.ComprobanteTipo = MovimientoCuentaCorrienteSocio.TIPO.FACTURA_C
                        mov.FechaIngreso = Utils.DateTo8601(Now.Date)
                        mov.Importe = c.monto * -1
                        mov.ImporteCobrar = mov.Importe
                        mov.Save(MovimientoCuentaCorrienteSocio.Guardar.NUEVO)

                        c.observaciones = $"FC-{feNum}"
                        c.MovimientoCC = mov.Id
                        c.Save(sCuota, 1)
                        mov.CuotasSociales.Add(c)
                        mov.Save(MovimientoCuentaCorrienteSocio.Guardar.EDITAR)
                    Else
                        Debug.Print($"{currSocio.Apellido} {currSocio.Nombre}")
                        log.SetError($"Generacion de factura [ FALLO ]: No se pudo facturar {currSocio.Apellido} {currSocio.Nombre} ", "Executor", "feNum")
                        c.Delete(sCuota, c.id)
                    End If
                End If
                currentProcess += 1
            End While

            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ OK ]")
            End If

            Return True
        End If

        Return False
    End Function

    Private Sub GetDatosContribuyente(ByVal cuit As Long, ByVal globalConfig As GlobalConfig, ByRef razonSocial As String, ByRef domicilio As String)
        Dim afip As New Afip(globalConfig)

        Dim estadoErr As String = ""

        If Not afip.VerificarEstadoServicioPadron(estadoErr) Then
            MsgBox(estadoErr, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Estado servicios AFIP")
            Exit Sub
        End If

        Dim ConsoleOut As New ConsoleOut
        Dim sAuth As New SQLEngine
        sAuth.dbType = helix.SQLEngine.dataBaseType.SQL_SERVER
        sAuth.DatabaseName = "soccam_test" 'Cambiar con el nombre de la DB correspondiente.
        sAuth.RequireCredentials = False
        'sAuth.Username = "prueba"
        'sAuth.Password = 123456

        'Si la DB esta en tu maquina, no es necesario cambiar esta linea.
        sAuth.Path = My.Computer.Name & "\" & "SQLEXPRESS"

        If Not sAuth.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            Exit Sub
        End If

        Dim auth As New AfipAuth(sAuth)
        Dim login As New AfipLogin(Afip.SERVICIO_PADRON, afip.AUTH_URL)

        Dim cert As New AfipCert(sAuth)
        cert.Homologacion = afip.Homologacion
        auth.Homologacion = afip.Homologacion

        If cert.LoadActive Then
            If Not auth.LoadActive(login.Serv, Now.Ticks) Then
                Dim loginError As String = ""
                If Not login.Login(cert.Certificado, loginError) Then
                    Exit Sub
                Else
                    auth.Sign = login.Sign
                    auth.Token = login.Token
                    auth.Req = login.XDocRequest.ToString
                    auth.Res = login.XDocResponse.ToString
                    auth.GenerationTime = login.GenerationTime.Ticks
                    auth.ExpirationTime = login.ExpirationTime.Ticks
                    auth.Servicio = login.Serv

                    If Not auth.Save(AfipAuth.Guardar.NUEVO) Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        Dim personaService As New WSPSA5.PersonaServiceA5
        personaService.Url = afip.INFO_URL

        Dim cuitData As New WSPSA5.personaReturn

        Try
            cuitData = personaService.getPersona(auth.Token, auth.Sign, globalConfig.Cuit, CLng(cuit))
            If IsNothing(cuitData.datosRegimenGeneral) And IsNothing(cuitData.datosMonotributo) Then
                razonSocial = ""
                domicilio = ""
            Else
                If IsNothing(cuitData.datosGenerales.razonSocial) Then
                    razonSocial = $"{cuitData.datosGenerales.apellido} {cuitData.datosGenerales.nombre}"
                Else
                    razonSocial = cuitData.datosGenerales.razonSocial
                End If
                domicilio = $"{cuitData.datosGenerales.domicilioFiscal.direccion} - {cuitData.datosGenerales.domicilioFiscal.localidad}, {cuitData.datosGenerales.domicilioFiscal.descripcionProvincia}"
            End If
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub


    Public Function GenerarFE(ByVal periodo As Integer, ByVal anio As Integer,
                              ByVal socio As Socio,
                              ByVal homologacion As Boolean,
                              ByVal importe As Decimal,
                              ByVal cuotaId As Integer,
                              ByVal enviarPorMail As Boolean,
                              Optional ByRef facturaID As Integer = 0) As Integer
        Dim ConsoleOut As New ConsoleOut

        Dim sFe As New SQLEngine
        Dim sFX As New SQLEngine
        Dim sCF As New SQLEngine
        Dim sGC As New SQLEngine

        Dim log As New Log
        log.LogFilePath = Module1.LOG_DIR

        If Module1.LOG Then
            log.LogLevel = 2
        End If

        sFe.dbType = helix.SQLEngine.dataBaseType.SQL_SERVER
        sFe.DatabaseName = "soccam_test" 'Cambiar con el nombre de la DB correspondiente.
        sFe.RequireCredentials = False
        'sFe.Username = "prueba"
        'sFe.Password = 123456

        'Si la DB esta en tu maquina, no es necesario cambiar esta linea.
        sFe.Path = My.Computer.Name & "\" & "SQLEXPRESS"

        'sFe.Path = If(My.Computer.Name = "ALPHA", My.Computer.Name, My.Computer.Name & "\" & "SQLEXPRESS")
        'sFe.DatabaseName = If(My.Computer.Name = "ALPHA", "soccam_ccb", "soccam")
        If Not sFe.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            Return False
        End If

        sFX = sFe
        If Not sFX.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            Return False
        End If

        sCF = sFe
        If Not sCF.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            Return False
        End If

        sGC = sFe
        If Not sGC.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            Return False
        End If

        Dim FE As New AfipFactura(sFe)
        Dim FEX As New AfipFacturaEX(sFX)
        Dim FERender As New AfipFERenderer
        Dim condicionFiscal As New AfipCondicionFiscal(sCF)


        Dim globalConfig As New GlobalConfig(sGC)
        If Not globalConfig.LoadMe(1) Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            Return False
        End If


        FE.Homologacion = Not globalConfig.Produccion
        FE.CuitEmisor = globalConfig.Cuit
        FE.PuntoVenta = globalConfig.PuntoVenta
        'FE.Numero = FE.GetUltimoNumeroLocal + 1
        FE.Numero = FE.GetUltimoNumero + 1
        If FE.Numero = 0 Then
            Return 0
        End If

        FE.ComprobanteTipo = AfipFactura.Tipo.FACTURA_C
        FE.FechaEmision = Utils.DateTo8601(Now.Date)

        FE.Concept = AfipFactura.Concepto.SERVICIOS
        FE.FechaServicioDesde = Utils.GetPrimerDiaMesISO(periodo, anio)
        FE.FechaServicioHasta = Utils.GetUltimoDiaMesISO(periodo, anio)
        FE.FechaVencimientoPago = Utils.GetUltimoDiaMesISO(Now.Month - 1, Now.Year)

        FE.Moneda = AfipFactura.MONEDA_PESO
        FE.MonedaCotizacion = 1

        condicionFiscal.LoadMe(FE.CuitEmisor.ToString)
        FEX.DomicilioEmisor = globalConfig.DomicilioComercial

        ' Cambiar la razon social para otros clientes
        FEX.RazonSocialEmisor = "CAMARA COMERCIAL E INDUSTRIAL DE BOLIVAR"

        Dim razonSocial As String = ""
        Dim domicilio As String = ""

        If socio.CUIT.Trim.Length = 11 Then
            FE.DocumentoTipo = AfipFactura.Documento.CUIT
            FE.DocumentoNumero = CLng(socio.CUIT)
            GetDatosContribuyente(socio.CUIT, globalConfig, razonSocial, domicilio)
            If razonSocial = "" Then
                Try
                    If socio.DNI.Trim.Length >= 6 Then
                        FE.DocumentoTipo = AfipFactura.Documento.DNI
                        FE.DocumentoNumero = CLng(socio.DNI)
                    Else
                        Return 0
                    End If
                Catch ex As Exception
                    Return 0
                End Try
                FEX.RazonSocialReceptor = $"{socio.Apellido} {socio.Nombre}"
                FEX.CondicionFiscalStringReceptor = "Consumidor Final"
                FEX.DomicilioReceptor = socio.Domicilio
            Else
                FEX.RazonSocialReceptor = razonSocial
                condicionFiscal.LoadMe(socio.CUIT)
                FEX.CondicionFiscalStringReceptor = GetCondicionFiscalString(condicionFiscal.Condicion)
                FEX.DomicilioReceptor = domicilio
            End If
        Else
            FE.DocumentoTipo = AfipFactura.Documento.DNI
            FE.DocumentoNumero = CLng(socio.DNI)
            FEX.RazonSocialReceptor = $"{socio.Apellido} {socio.Nombre}"
            FEX.CondicionFiscalStringReceptor = "Consumidor Final"
            FEX.DomicilioReceptor = socio.Domicilio
        End If


        FEX.CondicionContado = True


        Dim totalFactura As Decimal = 0

        Dim det As New AfipFacturaDetalle
        det.Codigo = "0"
        det.ProductoServicio = $"Cuota social {periodo + 1}/{anio}"
        det.Cantidad = 1
        det.UnidadMedida = AfipFacturaDetalle.Unidad.OTRAS_UNIDADES
        det.PrecioUnitario = importe
        det.BonificacionPercent = 0
        det.CuotaId = cuotaId

        totalFactura += (det.PrecioUnitario * det.Cantidad) - ((det.PrecioUnitario * det.Cantidad) * det.BonificacionPercent) / 100
        FE.Detalles.Add(det)

        FE.ImporteTotal = totalFactura
        FE.ImporteNeto = FE.ImporteTotal

        Dim numeroFE As Integer = 0

        If FE.Autorizar Then
            FE.Save(AfipFactura.Guardar.NUEVO)
            FEX.FacturaId = FE.Id
            facturaID = FE.Id
            FEX.Save(AfipFacturaEX.Guardar.NUEVO)
            FEX.FacturaRendered = FERender.templateFE(FE, FEX, globalConfig)
            numeroFE = FE.Numero

            If enviarPorMail Then
                If socio.EnviarMail And My.Computer.Name <> "ALPHA" Then
                    EnviarFacturaMailAuto(1, socio.InternalID, numeroFE, globalConfig, homologacion, FEX.FacturaRendered)
                End If
            End If
        End If

        Return numeroFE
    End Function


    Public Function GetPeriodoFromFecha(ByVal mes As Integer, ByVal mesesPorPeriodo As Byte) As Integer
        Return ((mes) \ mesesPorPeriodo)
    End Function


    Public Function GetCondicionFiscalString(ByVal condicionFiscal As Integer) As String
        Select Case condicionFiscal
            Case 1
                Return "IVA Responsable Inscripto"
            Case 4
                Return "IVA Sujeto Exento"
            Case 5
                Return "Consumidor Final"
            Case 6
                Return "Responsable Monotributo"
            Case 8
                Return "Proveedor del Exterior"
            Case 9
                Return "Cliente del Exterior"
            Case 10
                Return "IVA Liberado - Ley Nº 19.640"
            Case 11
                Return "IVA Responsable Inscripto - Agente de Percepción"
            Case 13
                Return "Monotributista Social"
            Case Else
                Return "IVA No Alcanzado"

        End Select
    End Function

    ''' <summary>
    ''' Envia en automatico las facturas
    ''' </summary>
    ''' <param name="origen">Origen del evento
    ''' 0: Todos
    ''' 1: Cuota social
    ''' 2: Pago Contratos
    ''' 3: FE manual
    ''' </param>
    ''' <param name="origenID">Origen </param>
    ''' <returns>
    ''' 0: si termino correctamente
    ''' 1: No tiene destinatario
    ''' </returns>
    Public Function EnviarFacturaMailAuto(ByVal origen As Integer, ByVal origenID As Integer, ByVal facturaNumero As Integer, ByVal globalConfig As GlobalConfig, ByVal homologacion As Boolean, Optional ByVal facturaRenderizada As String = "") As Integer
        Dim fe As New AfipFactura(SqleGlobal)
        Dim fx As New AfipFacturaEX(SqleGlobal)
        Dim lstCorreos As New List(Of String)

        fe.Homologacion = Not globalConfig.Produccion

        Select Case origen
            Case 0
                fe.LoadMe(origenID, facturaNumero, fx)
                lstCorreos = GetListaEmails(fe.DocumentoNumero, 0)
            Case 1
                Dim cSocio As New CuotaSocio()
                cSocio.LoadMe(SqleGlobal, origenID)
                fe.LoadMe(globalConfig.PuntoVenta, facturaNumero, fx)
                lstCorreos = GetListaEmails(fe.DocumentoNumero, 1)
        End Select

        If lstCorreos.Count = 0 Then
            Return 1
        End If

        Dim htmlFE As String
        Dim htmlMail As String
        Dim fileName As String
        Dim feNombre As String

        Dim render As New AfipFERenderer()

        htmlFE = facturaRenderizada

        htmlMail = facturaRenderizada
        feNombre = $"{fe.ComprobanteTipo} {Utils.ComponerNumeroComprobante(globalConfig.PuntoVenta, fe.Numero)}".Replace("_", " ")
        fileName = $"{fe.DocumentoNumero}_{Utils.ComponerNumeroComprobante(globalConfig.PuntoVenta, fe.Numero)}.pdf"

        Dim fullPath As String = $"{IO.Path.GetTempPath}{fileName}"

        Try
            Dim pdfRender As New HtmlToPdf
            Dim doc As PdfDocument = pdfRender.ConvertHtmlString(htmlMail)
            doc.Save(fullPath)
            doc.Close()
            doc = Nothing
            GC.Collect()

            Dim m As New Mail()
            m.Smtp_username = "pamela.gelvez@camarabolivar.com.ar"
            m.Smtp_password = "PamGel-7021!"
            m.Smtp_host = "mail.camarabolivar.com.ar"
            m.Smtp_port = 25
            m.Smtp_SSL = False

            m.FromAddress = "pamela.gelvez@camarabolivar.com.ar"
            m.FromName = "Cámara Bolívar"
            m.Subject = $"La Cámara te acerca tu {feNombre}"
            m.ToAddress = lstCorreos(0)
            m.IsHTML = True
            m.HTMLCode = htmlMail.Replace("<br/>{#}<br/><br/>", "")

            m.Adjunto = fullPath

            m.SendMail()
            m = Nothing
            GC.Collect()

            IO.File.Delete(fullPath)

        Catch ex As Exception
            If ex.HResult = -2147024864 Then
                Return 0
            Else
                Return 2
            End If
        End Try

        Return 0

    End Function


    ''' <summary>
    ''' Buscar los correos correspondientes a un DNI o CUIT
    ''' </summary>
    ''' <param name="documento">Numero de DNI o CUIT</param>
    ''' <param name="buscarBancoDatos">El banco de datos que usar
    ''' 0: Todos
    ''' 1: Cuota social
    ''' 2: Contratos
    ''' 3: Agenda
    ''' </param>
    ''' <returns></returns>
    Public Function GetListaEmails(ByVal documento As String, Optional buscarBancoDatos As Integer = 0) As List(Of String)
        Dim lstMails As New List(Of String)
        With SqleGlobal.Query
            If buscarBancoDatos = 1 Or buscarBancoDatos = 0 Then
                .Reset()
                .TableName = TABLA_SOCIO.TABLA_NOMBRE
                .AddSelectColumn(TABLA_SOCIO.MAIL)
                .WHEREstring = $"({TABLA_SOCIO.CUIT} = ? OR {TABLA_SOCIO.DNI} = ?) AND {TABLA_SOCIO.DELETED} = ?"
                .AddWHEREparam(documento)
                .AddWHEREparam(documento)
                .AddWHEREparam(False)
                If .Query() Then
                    While .QueryRead
                        lstMails.Add(.GetQueryData(0))
                    End While
                End If
            End If

            If buscarBancoDatos = 2 Or buscarBancoDatos = 0 Then
                .Reset()
                .TableName = TABLA_CONTRATOS_COFRES_USUARIOS.TABLA_NOMBRE
                .AddSelectColumn(TABLA_CONTRATOS_COFRES_USUARIOS.MAIL)
                .WHEREstring = $"({TABLA_CONTRATOS_COFRES_USUARIOS.DNI} = ? OR {TABLA_CONTRATOS_COFRES_USUARIOS.CUIT} = ?) AND {TABLA_CONTRATOS_COFRES_USUARIOS.DELETED} = ?"
                .AddWHEREparam(documento)
                .AddWHEREparam(documento)
                .AddWHEREparam(False)
                If .Query() Then
                    While .QueryRead
                        lstMails.Add(.GetQueryData(0))
                    End While
                End If
            End If

            If buscarBancoDatos = 3 Or buscarBancoDatos = 0 Then
                .Reset()
                .TableName = Contacto.TABLA.TABLA_NOMBRE
                .AddSelectColumn(Contacto.TABLA.MAIL)
                .WHEREstring = $"({Contacto.TABLA.DNI} = ? OR {Contacto.TABLA.CUIT} = ?) AND {Contacto.TABLA.DELETED} = ?"
                .AddWHEREparam(documento)
                .AddWHEREparam(documento)
                .AddWHEREparam(False)
                If .Query() Then
                    While .QueryRead
                        lstMails.Add(.GetQueryData(0))
                    End While
                End If
            End If
        End With

        Return lstMails
    End Function

    Public Function ActualizarEstadoCuotasCofres() As Integer

    End Function

End Class
