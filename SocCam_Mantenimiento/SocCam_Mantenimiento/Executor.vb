Imports System.IO
Imports System.IO.Compression
Imports System.Windows.Documents
Imports System.Windows.Shapes
Imports helix
Imports SelectPdf
Imports SocCam_Mantenimiento.WSPSA5

Public Class Executor
    Public Property Silent As Boolean = False
    Public Property globalConfig As New GlobalConfig()
    Public Property config As New ConfigDatabase
    Public Property SqleGlobal As New SQLEngine

    Dim ConsoleOut As New ConsoleOut

    Public Sub New()
        Me.SqleGlobal = config.sqle

        If SqleGlobal.IsStarted Then
            Me.SqleGlobal.ColdBoot()
        Else
            Me.SqleGlobal.Start()
        End If
    End Sub

    Public Function ActualizarPadronAfip(ByVal url As String) As Boolean
        Dim filePath = My.Computer.FileSystem.SpecialDirectories.Temp
        If Not Utils.DescargarArchivo(url, filePath, Silent) Then
            Return False
        End If
        Dim objProcess As System.Diagnostics.Process

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

            If Not SqleGlobal.Start Then
                If Not Silent Then
                    ConsoleOut.Print($"- Actualizar padron: No se pudo conectar a la base de datos [FAIL]")
                End If
                Return False
            End If
            Dim tst As New AfipCondicionFiscal(SqleGlobal)

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

        If Not SqleGlobal.Start Then
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

        If socios.LoadAll(SqleGlobal, dtSocios, True) Then
            dtrSocios = dtSocios.CreateDataReader

            If Not Silent Then
                ConsoleOut.Print($"- Generando cuotas sociales")
            End If

            Dim totalSocios As Integer = dtSocios.Rows.Count - 1
            Dim currentProcess As Integer = 1

            While dtrSocios.Read

                Dim currSocio As New Socio
                currSocio.LoadMe(SqleGlobal, dtrSocios(TABLA_SOCIO.ID))

                If Not Silent Then
                    ConsoleOut.Print($"{ConsoleOut.ProgressBarStep} {currentProcess}/{totalSocios} - {currSocio.Nombre.Trim.PadRight(80, " ")}")
                End If

                Dim plan As New SocioTipo
                plan.sqle = SqleGlobal
                plan.LoadMe(dtrSocios(TABLA_SOCIO.TIPO_SOCIO))

                Dim periodo As Integer = GetPeriodoFromFecha(mes, plan.getMesesPorPeriodo)

                Dim c As New CuotaSocio
                c.sqle = SqleGlobal

                If Not c.CuotaExist(SqleGlobal, periodo, plan.periodicidad, anio, dtrSocios(TABLA_SOCIO.ID)) Then

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
                    cobrador.LoadMe(SqleGlobal, dtrSocios(TABLA_SOCIO.SECTOR), True)
                    c.cobradorID = cobrador.ID
                    c.estado = CuotaSocio.ESTADO_SOCIO.PENDIENTE
                    c.socioID = dtrSocios(TABLA_SOCIO.ID)
                    If autocobrar And c.monto = 0 Then
                        c.estado = CuotaSocio.ESTADO_SOCIO.AL_DIA
                        c.fechaPago = Now
                    End If
                    c.Save(SqleGlobal, 0)
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
                                        ByVal omitirUsuariosCofres As Boolean) As Boolean
        Dim socios As New Socio
        Dim dtSocios As New DataTable
        Dim dtrSocios As DataTableReader
        Dim ConsoleOut As New ConsoleOut
        Dim sSocios As New SQLEngine
        Dim log As New Log
        log.LogFilePath = Module1.LOG_DIR

        If Module1.LOG Then
            log.LogLevel = 2
        End If

        If Not SqleGlobal.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            log.SetError("Generacion de cuotas sociales [ FALLO ]: No se pudo iniciar el motor SQL socios", "Executor", "sSocios.Start")
            Return False
        End If


        globalConfig.Sqle = SqleGlobal
        If Not globalConfig.LoadMe(1) Then
            ConsoleOut.Print("No se pudo cargar la configuraicon global")
        End If


        ' Buscar todos los socios activos
        If socios.LoadAll(SqleGlobal, dtSocios, True) Then
            dtrSocios = dtSocios.CreateDataReader

            Dim totalSocios As Integer = dtSocios.Rows.Count - 1

            If Not Silent Then
                ConsoleOut.Print($"- Generando cuotas sociales")
            End If

            Dim currentProcess As Integer = 1

            Dim desde As DateTime = Now.Date
            Dim hoy As DateTime = DateTime.Today
            Dim primerDiaDelProximoMes As DateTime = New DateTime(hoy.Year, hoy.Month, 1).AddMonths(1)
            Dim ultimoDiaDelMesActual As DateTime = primerDiaDelProximoMes.AddDays(-1)
            Dim hasta As DateTime = ultimoDiaDelMesActual
            ConsoleOut.Print($"hasta: {hasta}")

            ConsoleOut.Print($"Iterando Socios.")
            While dtrSocios.Read
                Dim currSocio As New Socio
                currSocio.LoadMe(SqleGlobal, dtrSocios(TABLA_SOCIO.ID))
                ConsoleOut.Print($"Socio encontrado: {dtrSocios(TABLA_SOCIO.ID)}")

                'If Not Silent Then
                '    ConsoleOut.UpdateLastLine($"{ConsoleOut.ProgressBarStep} {currentProcess}/{totalSocios} - {currSocio.Nombre.Trim.PadRight(80, " ")}")
                'Else
                '    Debug.Print($"{ConsoleOut.ProgressBarStep} {currentProcess}/{totalSocios} - {currSocio.Nombre.Trim.PadRight(80, " ")}")
                'End If

                Dim fechaAprobacion As DateTime = dtrSocios(TABLA_SOCIO.FECHA_APROBACION)


                If omitirUsuariosCofres Then
                    Dim tieneCofre As Boolean = False

                    Dim contratoCofre As New ContratoCofre(SqleGlobal)
                    contratoCofre.QuickSearch(ContratoCofre.TABLA.ES_SOCIO_ID, SQLEngineQuery.OperatorCriteria.Igual, dtrSocios(TABLA_SOCIO.ID))

                    If contratoCofre.SearchResult.Count > 0 Then
                        For Each contrato As ContratoCofre In contratoCofre.SearchResult

                            'El If se ejecutará si y solo si contrato.Deleted = False.
                            If Not contrato.Deleted Then
                                tieneCofre = True
                                Exit For
                            End If
                        Next
                    End If


                    If tieneCofre Then
                        ConsoleOut.Print($"Tiene cofre. Continua al siguiente socio.")
                        ConsoleOut.Print($" ")
                        currentProcess += 1
                        Continue While
                    End If
                End If

                Dim plan As New SocioTipo
                plan.sqle = SqleGlobal
                plan.LoadMe(dtrSocios(TABLA_SOCIO.TIPO_SOCIO))

                Dim c As New CuotaSocio
                c.sqle = SqleGlobal

                Dim periodo As Integer = GetPeriodoFromFecha(mes, plan.getMesesPorPeriodo)

                If Not c.CuotaExist(SqleGlobal, periodo, plan.periodicidad, anio, dtrSocios(TABLA_SOCIO.ID)) Then


                    If (fechaAprobacion.Year > anio) Then
                        ConsoleOut.Print($"Socio dado de alta posterior al año {anio}.")
                        Continue While
                    ElseIf (fechaAprobacion.Year = anio) And ((fechaAprobacion.Month - 1) >= periodo) Then
                        ConsoleOut.Print($"Socio dado de alta posterior al mes {periodo + 1} del año {anio}.")
                        Continue While
                    End If

                    c.anio = anio
                    c.Periodo = periodo
                    c.Periodicidad = plan.periodicidad
                    c.PlanID = plan.id
                    c.monto = plan.importe
                    c.cobradorID = 1
                    c.estado = CuotaSocio.ESTADO_SOCIO.PENDIENTE
                    c.socioID = dtrSocios(TABLA_SOCIO.ID)

                    If c.monto = 0 Then
                        Debug.Print($"{currSocio.Apellido} {currSocio.Nombre}")
                        c.Delete(SqleGlobal, c.id)
                        Continue While
                    End If

                    'Recibo -----------------------------------------------------------------------------------
                    Dim lstDetalles As New List(Of AfipFacturaDetalle)
                    Dim detalle As New AfipFacturaDetalle()

                    Dim descripcion = $"Cuota social {Utils.GetNombreMes(periodo)} {Now.Year}"
                    detalle.ProductoServicio = $"{descripcion}: <b>{Utils.ToMoneyFormat(c.monto)}</b>"
                    detalle.Cantidad = 1
                    detalle.UnidadMedida = AfipFacturaDetalle.Unidad.OTRAS_UNIDADES
                    detalle.PrecioUnitario = c.monto
                    detalle.Codigo = 0

                    lstDetalles.Add(detalle)

                    Dim Localidades As New Localidad
                    Localidades.sqle = SqleGlobal
                    Localidades.LoadAll()

                    Dim numeroComprobante As Integer = 0
                    Dim idComprobante As Integer = GenerarComprobante(AfipFactura.Tipo.RECIBO, c.socioID, lstDetalles, globalConfig, desde, hasta, mes, anio, Localidades, numeroComprobante)

                    If idComprobante <= 0 Then
                        Utils.Scream("No se pudo guardar el comprobante. Vuelva a intentar.")
                    End If
                    'Recibo end -------------------------------------------------------------------------------

                    c.observaciones = $"RX-{numeroComprobante}"
                    c.Save(SqleGlobal, 0)

                    'Movimiento -------------------------------------------------------------------------------
                    Dim mov As New MovimientoCuentaCorrienteSocio(SqleGlobal)
                    Dim fechaIngreso As Date = New Date(anio, mes, 1)
                    mov.ClienteId = c.socioID
                    mov.ComprobanteRelacionado = idComprobante
                    mov.ComprobanteTipo = MovimientoCuentaCorrienteSocio.TIPO.RECIBO_X
                    mov.FechaIngreso = Utils.DateTo8601(fechaIngreso)
                    mov.Importe = c.monto * -1
                    mov.ImporteCobrar = mov.Importe
                    mov.CuotasSociales.Add(c)

                    mov.Save(MovimientoCuentaCorrienteSocio.Guardar.NUEVO)
                    'Movimiento end ---------------------------------------------------------------------------


                    'Else
                    '    Debug.Print($"{currSocio.Apellido} {currSocio.Nombre}")
                    '    log.SetError($"Generacion de factura [ FALLO ]: No se pudo facturar {currSocio.Apellido} {currSocio.Nombre} ", "Executor", "feNum")
                    '    c.Delete(sCuota, c.id)
                    'End If

                End If
                currentProcess += 1
                ConsoleOut.Print(" ")
            End While

            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ OK ]")
            End If

            Return True
        End If

        Return False
    End Function

    Public Function GetDatosContribuyente(ByVal cuit As Long, ByVal globalConfig As GlobalConfig)
        Dim log As New Log

        Dim afip As New Afip(globalConfig)

        Dim estadoErr As String = ""
        If Not afip.VerificarEstadoServicioPadron(estadoErr) Then
            'log.SetError($"Verificar estado servicio padron [ FALLO ]: {estadoErr}", "Verificar estado servicio padron", "Afip")
            ConsoleOut.Print("Error al verificar el estado del servicio del padron.")
            Exit Function
        End If

        If Not SqleGlobal.Start Then
            If Not Silent Then
                ConsoleOut.Print($"")
                ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
            End If
            Exit Function
        End If

        Dim auth As New AfipAuth(SqleGlobal)
        Dim login As New AfipLogin(Afip.SERVICIO_PADRON, afip.AUTH_URL)

        Dim cert As New AfipCert(SqleGlobal)
        cert.Homologacion = afip.Homologacion
        auth.Homologacion = afip.Homologacion

        If cert.LoadActive Then
            If Not auth.LoadActive(login.Serv, Now.Ticks) Then
                Dim loginError As String = ""
                If Not login.Login(cert.Certificado, loginError) Then
                    Exit Function
                Else
                    auth.Sign = login.Sign
                    auth.Token = login.Token
                    auth.Req = login.XDocRequest.ToString
                    auth.Res = login.XDocResponse.ToString
                    auth.GenerationTime = login.GenerationTime.Ticks
                    auth.ExpirationTime = login.ExpirationTime.Ticks
                    auth.Servicio = login.Serv

                    If Not auth.Save(AfipAuth.Guardar.NUEVO) Then
                        Exit Function
                    End If
                End If
            End If
        End If

        Dim personaService As New WSPSA5.PersonaServiceA5
        personaService.Url = afip.INFO_URL

        Dim cuitData As New WSPSA5.personaReturn
        Dim condicionFiscal As String = String.Empty
        Dim cuitRepresentada As String = "30528257921"


        Try
            cuitData = personaService.getPersona(auth.Token, auth.Sign, globalConfig.Cuit, cuit)

            If Not IsNothing(cuitData) Then

                If Not IsNothing(cuitData.datosMonotributo) Then
                    If cuitData.datosMonotributo.categoriaMonotributo.descripcionCategoria = "A MONOTRIBUTO SOCIAL LOCACION" Or cuitData.datosMonotributo.categoriaMonotributo.descripcionCategoria = "A MONOTRIBUTO SOCIAL VENTAS" Then
                        condicionFiscal = "Monotributista Social"

                    Else

                        For Each impuesto As impuesto In cuitData.datosMonotributo.impuesto
                            Dim descripcion As String = impuesto.descripcionImpuesto

                            If descripcion = "MONOTRIBUTO" Then
                                condicionFiscal = "Responsable Monotributo"
                                Exit For
                            End If
                        Next

                    End If

                ElseIf Not IsNothing(cuitData.datosRegimenGeneral) Then

                    For Each impuesto As impuesto In cuitData.datosRegimenGeneral.impuesto
                        Dim descripcion As String = impuesto.descripcionImpuesto

                        If descripcion = "IVA EXENTO" Then
                            condicionFiscal = "IVA Sujeto Exento"
                            Exit For
                        ElseIf descripcion = "IVA" Then
                            condicionFiscal = "IVA Responsable Inscripto"
                            Exit For
                        End If
                    Next

                Else
                    condicionFiscal = "Consumidor Final"
                End If

            End If

            Return condicionFiscal

        Catch ex As Exception
            'log.SetError($"Comprobar condicion fiscal contribuyente [ FALLO ]: {ex}", "Comprobar condicion fiscal contribuyente", "Afip")
            ConsoleOut.Print("Error al verificar la condicion fiscal del socio.")
        End Try
    End Function

    'Public Function GenerarFE(ByVal periodo As Integer, ByVal anio As Integer,
    '                          ByVal socio As Socio,
    '                          ByVal homologacion As Boolean,
    '                          ByVal importe As Decimal,
    '                          ByVal cuotaId As Integer,
    '                          ByVal enviarPorMail As Boolean,
    '                          Optional ByRef facturaID As Integer = 0) As Integer

    '    Dim ConsoleOut As New ConsoleOut

    '    Dim log As New Log
    '    log.LogFilePath = Module1.LOG_DIR

    '    If Module1.LOG Then
    '        log.LogLevel = 2
    '    End If

    '    Dim FE As New AfipFactura(SqleGlobal)
    '    Dim FEX As New AfipFacturaEX(SqleGlobal)
    '    Dim FERender As New AfipFERenderer
    '    Dim condicionFiscal As New AfipCondicionFiscal(SqleGlobal)


    '    Dim globalConfig As New GlobalConfig(SqleGlobal)
    '    If Not globalConfig.LoadMe(1) Then
    '        If Not Silent Then
    '            ConsoleOut.Print($"")
    '            ConsoleOut.Print($"- Generacion de cuotas sociales [ FALLO ]")
    '        End If
    '        Return False
    '    End If


    '    FE.Homologacion = Not globalConfig.Produccion
    '    FE.CuitEmisor = globalConfig.Cuit
    '    FE.PuntoVenta = globalConfig.PuntoVenta
    '    FE.Numero = FE.GetUltimoNumero + 1
    '    If FE.Numero = 0 Then
    '        Return 0
    '    End If

    '    FE.ComprobanteTipo = AfipFactura.Tipo.FACTURA_C
    '    FE.FechaEmision = Utils.DateTo8601(Now.Date)
    '    FE.Concept = AfipFactura.Concepto.SERVICIOS
    '    FE.FechaServicioDesde = Utils.GetPrimerDiaMesISO(periodo, anio)
    '    FE.FechaServicioHasta = Utils.GetUltimoDiaMesISO(periodo, anio)
    '    FE.FechaVencimientoPago = Utils.GetUltimoDiaMesISO(Now.Month - 1, Now.Year)
    '    FE.Moneda = AfipFactura.MONEDA_PESO
    '    FE.MonedaCotizacion = 1

    '    condicionFiscal.LoadMe(FE.CuitEmisor.ToString)
    '    FEX.DomicilioEmisor = globalConfig.DomicilioComercial
    '    FEX.RazonSocialEmisor = "CAMARA COMERCIAL E INDUSTRIAL DE BOLIVAR"

    '    Dim razonSocial As String = ""
    '    Dim domicilio As String = ""

    '    If socio.CUIT.Trim.Length = 11 Then
    '        FE.DocumentoTipo = AfipFactura.Documento.CUIT
    '        FE.DocumentoNumero = CLng(socio.CUIT)
    '        GetDatosContribuyente(socio.CUIT, globalConfig, razonSocial, domicilio)
    '        If razonSocial = "" Then
    '            Try
    '                If socio.DNI.Trim.Length >= 6 Then
    '                    FE.DocumentoTipo = AfipFactura.Documento.DNI
    '                    FE.DocumentoNumero = CLng(socio.DNI)
    '                Else
    '                    Return 0
    '                End If
    '            Catch ex As Exception
    '                Return 0
    '            End Try
    '            FEX.RazonSocialReceptor = $"{socio.Apellido} {socio.Nombre}"
    '            FEX.CondicionFiscalStringReceptor = "Consumidor Final"
    '            FEX.DomicilioReceptor = socio.Domicilio
    '        Else
    '            FEX.RazonSocialReceptor = razonSocial
    '            condicionFiscal.LoadMe(socio.CUIT)
    '            FEX.CondicionFiscalStringReceptor = GetCondicionFiscalString(condicionFiscal.Condicion)
    '            FEX.DomicilioReceptor = domicilio
    '        End If
    '    Else
    '        FE.DocumentoTipo = AfipFactura.Documento.DNI
    '        FE.DocumentoNumero = CLng(socio.DNI)
    '        FEX.RazonSocialReceptor = $"{socio.Apellido} {socio.Nombre}"
    '        FEX.CondicionFiscalStringReceptor = "Consumidor Final"
    '        FEX.DomicilioReceptor = socio.Domicilio
    '    End If

    '    FEX.CondicionContado = True

    '    Dim totalFactura As Decimal = 0

    '    Dim det As New AfipFacturaDetalle
    '    det.Codigo = "0"
    '    det.ProductoServicio = $"Cuota social {periodo + 1}/{anio}"
    '    det.Cantidad = 1
    '    det.UnidadMedida = AfipFacturaDetalle.Unidad.OTRAS_UNIDADES
    '    det.PrecioUnitario = importe
    '    det.BonificacionPercent = 0
    '    det.CuotaId = cuotaId

    '    totalFactura += (det.PrecioUnitario * det.Cantidad) - ((det.PrecioUnitario * det.Cantidad) * det.BonificacionPercent) / 100
    '    FE.Detalles.Add(det)

    '    FE.ImporteTotal = totalFactura
    '    FE.ImporteNeto = FE.ImporteTotal

    '    Dim numeroFE As Integer = 0

    '    If FE.Autorizar Then
    '        FE.Save(AfipFactura.Guardar.NUEVO)
    '        FEX.FacturaId = FE.Id
    '        facturaID = FE.Id
    '        FEX.Save(AfipFacturaEX.Guardar.NUEVO)
    '        FEX.FacturaRendered = FERender.templateFE(FE, FEX, globalConfig)
    '        numeroFE = FE.Numero

    '        If enviarPorMail Then
    '            If socio.EnviarMail And My.Computer.Name <> "ALPHA" Then
    '                EnviarFacturaMailAuto(1, socio.InternalID, numeroFE, globalConfig, homologacion, FEX.FacturaRendered)
    '            End If
    '        End If
    '    End If

    '    Return numeroFE
    'End Function


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

    ''' <summary>
    ''' Generar un comprobante por los detalles del movimiento
    ''' </summary>
    ''' <param name="tipoComprobante">Factura o recibo</param>
    ''' <param name="socioID">ID del socio a generar el comprobante</param>
    ''' <param name="detalles">Listado de detalles del comprobante</param>
    ''' <param name="gc">Configuración global</param>
    ''' <param name="localidades">Lista de localidades</param>
    ''' <returns>Id del comprobante generado si se generó correctamente, entero menor a 0 si falló</returns>
    ''' GenerarComprobante(AfipFactura.Tipo.RECIBO, c.socioID, lstDetalles, globalConfig, desde, hasta, , numeroComprobante)
    Public Function GenerarComprobante(ByVal tipoComprobante As AfipFactura.Tipo,
                                       socioID As Integer,
                                       ByVal detalles As List(Of AfipFacturaDetalle),
                                       ByVal gc As GlobalConfig,
                                       ByVal periodoDesde As Date,
                                       ByVal periodoHasta As Date,
                                        ByVal mes As Integer,
                                       ByVal anio As Integer,
                                       Optional localidades As Localidad = Nothing,
                                       Optional ByRef numeroComprobante As Integer = 0,
                                       Optional ByRef comprobanteRelacionado As AfipFactura = Nothing) As Integer

        Dim fact As New AfipFactura(SqleGlobal)
        fact.Homologacion = False
        fact.CuitEmisor = "30528257921"
        fact.PuntoVenta = 4
        fact.ComprobanteTipo = tipoComprobante

        Dim ultimoNumero = fact.GetUltimoNumeroLocal
        If ultimoNumero < 0 Then Return -1

        For Each det As AfipFacturaDetalle In detalles
            fact.Detalles.Add(det)
            fact.ImporteTotal += det.PrecioUnitario * det.Cantidad
            fact.ImporteNeto += det.PrecioUnitario * det.Cantidad
        Next

        fact.Numero = ultimoNumero + 1
        fact.Concept = AfipFactura.Concepto.SERVICIOS

        Dim socio As New SocioNT(SqleGlobal)
        If Not socio.LoadMe(socioID) Then Return -2

        If IsNumeric(socio.Cuit) Then
            fact.DocumentoNumero = socio.Cuit
            fact.DocumentoTipo = AfipFactura.Documento.CUIT
        Else
            If IsNumeric(socio.Dni) Then
                fact.DocumentoNumero = socio.Dni
                fact.DocumentoTipo = AfipFactura.Documento.DNI
            Else
                fact.DocumentoNumero = 99999999
                fact.DocumentoTipo = AfipFactura.Documento.DNI
            End If
        End If

        Dim fechaEmision As Date = New Date(anio, mes, 1)
        fact.FechaEmision = Utils.DateTo8601(Now.Date)
        fact.FechaServicioDesde = Utils.DateTo8601(periodoDesde)
        fact.FechaServicioHasta = Utils.DateTo8601(periodoHasta)
        Dim venc As New Date(anio, mes, 10)
        fact.FechaVencimiento = Utils.DateTo8601(venc)
        fact.FechaVencimientoPago = Utils.DateTo8601(venc)

        If Not fact.Save(AfipFactura.Guardar.NUEVO) Then Return -4

        Dim fx As New AfipFacturaEX(SqleGlobal)
        fx.FacturaId = fact.Id
        fx.CondicionContado = True

        If Not IsNumeric(socio.Cuit) Then
            fx.CondicionFiscalStringReceptor = "Consumidor Final"
        Else
            fx.CondicionFiscalStringReceptor = GetDatosContribuyente(socio.Cuit, gc)
        End If

        fx.DomicilioEmisor = "Las Heras 45 - Bolivar, Buenos Aires"
        fx.DomicilioReceptor = $"{ToSentenceCase(socio.Domicilio)} - {localidades.AllReverse(socio.Localidad).Split(",")(0)}, {localidades.AllReverse(socio.Localidad).Split(",")(2)}"
        fx.Operador = "0"
        fx.PuestoEmision = "PCADMIN"
        fx.RazonSocialEmisor = "CAMARA COMERCIAL E INDUSTRIAL DE BOLIVAR"
        fx.RazonSocialReceptor = socio.Nombre
        fx.Pagado = False
        fx.Save()

        numeroComprobante = fact.Numero

        Return fact.Id
    End Function

End Class
