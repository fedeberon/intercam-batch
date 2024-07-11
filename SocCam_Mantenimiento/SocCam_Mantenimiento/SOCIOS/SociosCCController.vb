Imports helix
Public Class SociosCCController
    Public Property sqle As New SQLEngine
    Public Property SearchResult As New List(Of MovimientoCuentaCorrienteSocio)

    Public Sub New()
    End Sub

    Public Sub New(ByVal sqle As SQLEngine)
        Me.sqle.RequireCredentials = sqle.RequireCredentials
        Me.sqle.Username = sqle.Username
        Me.sqle.Password = sqle.Password
        Me.sqle.dbType = sqle.dbType
        Me.sqle.Path = sqle.Path
        Me.sqle.DatabaseName = sqle.DatabaseName
        If sqle.IsStarted Then
            Me.sqle.ColdBoot()
        Else
            Me.sqle.Start()
        End If
    End Sub

    Public Function InjectSqle(ByVal sqle As SQLEngine)
        Me.sqle.RequireCredentials = sqle.RequireCredentials
        Me.sqle.Username = sqle.Username
        Me.sqle.Password = sqle.Password
        Me.sqle.dbType = sqle.dbType
        Me.sqle.Path = sqle.Path
        Me.sqle.DatabaseName = sqle.DatabaseName
        If sqle.IsStarted Then
            Return Me.sqle.ColdBoot()
        Else
            Return Me.sqle.Start()
        End If
    End Function

    Public Sub AgregarMovimiento(ByRef lst As ListView, ByVal clienteId As Integer)
        frmSociosCCMovimiento.Reset()

        Dim socio As New SocioNT(Me.sqle)
        socio.LoadMe(clienteId)

        frmSociosCCMovimiento.Text = $"Cuenta Corriente {socio.Nombre}"
        frmSociosCCMovimiento.socioID = socio.Id

        If frmSociosCCMovimiento.ShowDialog(frmMain) = DialogResult.OK Then
        End If
    End Sub

    ''' <summary>
    ''' Generar un comprobante por los detalles del movimiento
    ''' </summary>
    ''' <param name="tipoComprobante">Factura o recibo</param>
    ''' <param name="socioID">ID del socio a generar el comprobante</param>
    ''' <param name="detalles">Listado de detalles del comprobante</param>
    ''' <param name="gc">Configuración global</param>
    ''' <param name="localidades">Lista de localidades</param>
    ''' <returns>Id del comprobante generado si se generó correctamente, entero menor a 0 si falló</returns>
    Public Function GenerarComprobante(ByVal tipoComprobante As AfipFactura.Tipo,
                                       socioID As Integer,
                                       ByVal detalles As List(Of AfipFacturaDetalle),
                                       ByVal gc As GlobalConfig,
                                       ByVal periodoDesde As Date,
                                       ByVal periodoHasta As Date,
                                       Optional localidades As Localidad = Nothing,
                                       Optional ByRef numeroComprobante As Integer = 0,
                                       Optional ByRef comprobanteRelacionado As AfipFactura = Nothing) As Integer

        Dim fact As New AfipFactura(Me.sqle)
        fact.Homologacion = Not gc.Produccion
        fact.CuitEmisor = gc.Cuit
        fact.PuntoVenta = gc.PuntoVenta
        fact.ComprobanteTipo = tipoComprobante

        Dim ultimoNumero = 0

        If tipoComprobante = AfipFactura.Tipo.FACTURA_C Or tipoComprobante = AfipFactura.Tipo.NOTA_CREDITO_C Or tipoComprobante = AfipFactura.Tipo.NOTA_DEBITO_C Then
            ultimoNumero = fact.GetUltimoNumero
        Else
            ultimoNumero = fact.GetUltimoNumeroLocal
        End If

        If ultimoNumero < 0 Then Return -1

        For Each det As AfipFacturaDetalle In detalles
            fact.Detalles.Add(det)
            fact.ImporteTotal += det.PrecioUnitario * det.Cantidad
            fact.ImporteNeto += det.PrecioUnitario * det.Cantidad
        Next

        fact.Numero = ultimoNumero + 1
        fact.Concept = AfipFactura.Concepto.SERVICIOS

        Dim socio As New SocioNT(Me.sqle)
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

        fact.FechaEmision = Utils.DateTo8601(Now.Date)

        fact.FechaServicioDesde = Utils.DateTo8601(periodoDesde)
        fact.FechaServicioHasta = Utils.DateTo8601(periodoHasta)
        fact.FechaVencimiento = Utils.DateTo8601(Now.Date.AddDays(1))
        fact.FechaVencimientoPago = Utils.DateTo8601(Now.Date.AddDays(1))

        If fact.ComprobanteTipo = AfipFactura.Tipo.NOTA_CREDITO_C Or fact.ComprobanteTipo = AfipFactura.Tipo.NOTA_DEBITO_C Then
            fact.CbtesAsoc.Add(comprobanteRelacionado)
        End If

        If Not fact.Autorizar Then Return -3

        If Not fact.Save(AfipFactura.Guardar.NUEVO) Then Return -4

        Dim fx As New AfipFacturaEX(Me.sqle)

        fx.FacturaId = fact.Id
        fx.CondicionContado = True

        'If fact.DocumentoNumero = 99999999 Then
        '    fx.CondicionFiscalStringReceptor = "Consumidor Final"
        'Else
        '    fx.CondicionFiscalStringReceptor = GetCondicionFiscalString(socio.CondicionFiscal)
        'End If


        Dim frmImportarNuevoSocio As New frmImportarNuevoSocio()

        fx.CondicionFiscalStringReceptor = frmImportarNuevoSocio.getCondicionalFiscalSocio(socio.Cuit)
        MsgBox($"Condicion Fiscal: {fx.CondicionFiscalStringReceptor}", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, $"Condicion Fiscal: {fx.CondicionFiscalStringReceptor}")


        fx.DomicilioEmisor = gc.DomicilioComercial
        If IsNothing(localidades) Then
            fx.DomicilioReceptor = $"{ToSentenceCase(socio.Domicilio)} - {ToSentenceCase(frmMain.kernel.Localidades.AllReverse(socio.Localidad).Split(",")(0))}, {ToSentenceCase(frmMain.kernel.Localidades.AllReverse(socio.Localidad).Split(",")(2))}"
        Else
            fx.DomicilioReceptor = $"{ToSentenceCase(socio.Domicilio)} - {ToSentenceCase(localidades.AllReverse(socio.Localidad).Split(",")(0))}, {ToSentenceCase(localidades.AllReverse(socio.Localidad).Split(",")(2))}"
        End If

        fx.Operador = gc.UsuarioId
        fx.PuestoEmision = gc.NombrePuesto
        fx.RazonSocialEmisor = gc.RazonSocialEmisor
        fx.RazonSocialReceptor = socio.Nombre
        fx.Pagado = False

        fx.Save()

        numeroComprobante = fact.Numero

        Return fact.Id
    End Function

    Public Function LoadMovimientosList(ByRef lst As ListView, ByVal socioID As Integer) As Integer
        'lst.SuspendLayout()
        lst.Items.Clear()

        Dim movimientos As New MovimientoCuentaCorrienteSocio(Me.sqle)
        movimientos.QuickSearch(MovimientoCuentaCorrienteSocio.TABLA.CLIENTE_ID, SQLEngineQuery.OperatorCriteria.Igual, socioID, SortOrder.Ascending)

        Dim itmSaldo As New ListViewItem
        itmSaldo.Text = "Saldo"

        Dim saldo As Decimal = Me.GetSaldo(socioID, Utils.DateTo8601(Now))

        itmSaldo.UseItemStyleForSubItems = False
        If saldo >= 0 Then
            itmSaldo.SubItems.Add(Utils.ToMoneyFormat(saldo))
            If saldo = 0 Then
                itmSaldo.SubItems(1).ForeColor = Color.Black
            Else
                itmSaldo.SubItems(1).ForeColor = Color.Green
            End If
        Else
            itmSaldo.SubItems.Add(Utils.ToMoneyFormat(Math.Abs(saldo)))
            itmSaldo.SubItems(1).ForeColor = Color.Red
        End If

        lst.Items.Add(itmSaldo)

        ' Separador
        Dim itmDummy As New ListViewItem
        itmDummy.Text = ""
        itmDummy.SubItems.Add("")
        itmDummy.SubItems.Add("")
        lst.Items.Add(itmDummy)

        Dim itmHeader As New ListViewItem
        itmHeader.UseItemStyleForSubItems = False
        itmHeader.Text = $"Fecha"
        itmHeader.SubItems.Add("Debe")
        itmHeader.SubItems.Add("Haber")
        itmHeader.SubItems.Add("Comprobante")
        'itmHeader.SubItems.Add($"{Utils.GetNombreMes(balanceMensualDerechos.searchResult(balanceIndex).MesImputacion - 1)}/{balanceMensualDerechos.searchResult(balanceIndex).AnioImputacion}")
        itmHeader.BackColor = Color.WhiteSmoke
        itmHeader.SubItems(1).BackColor = Color.WhiteSmoke
        itmHeader.SubItems(2).BackColor = Color.WhiteSmoke
        itmHeader.SubItems(3).BackColor = Color.WhiteSmoke
        lst.Items.Add(itmHeader)

        For Each movimiento As MovimientoCuentaCorrienteSocio In movimientos.SearchResult
            Dim itm As New ListViewItem
            itm.Text = Utils.Int8601ToDate(movimiento.FechaIngreso)
            If movimiento.Importe < 0 Then
                itm.SubItems.Add(Utils.ToMoneyFormat(movimiento.Importe * -1))
                itm.SubItems.Add("")
                Dim f As New AfipFactura(Me.sqle)
                f.LoadMe(movimiento.ComprobanteRelacionado)

                itm.SubItems.Add(f.ToString)

            Else
                itm.SubItems.Add("")
                itm.SubItems.Add(Utils.ToMoneyFormat(movimiento.Importe))
                If movimiento.ComprobanteTipo = MovimientoCuentaCorrienteSocio.TIPO.NOTA_CREDITO_X Or
                        movimiento.ComprobanteTipo = MovimientoCuentaCorrienteSocio.TIPO.NOTA_CREDITO_C Then
                    Dim f As New AfipFactura(Me.sqle)
                    f.LoadMe(movimiento.ComprobanteRelacionado)

                    itm.SubItems.Add(f.ToString)
                Else
                    itm.SubItems.Add($"OP. N° {Utils.LeftPad(movimiento.Id, 8, "0")}")
                End If


            End If
            itm.Tag = movimiento


            If movimiento.Deleted Then
                Utils.TacharItem(itm)
            End If

            lst.Items.Add(itm)
        Next


        'lst.ResumeLayout()

        Return movimientos.SearchResult.Count
    End Function

    Public Function LoadMovimientosSocio(ByVal socioID As Integer) As Integer
        Dim movimientos As New MovimientoCuentaCorrienteSocio(Me.sqle)
        movimientos.QuickSearch(MovimientoCuentaCorrienteSocio.TABLA.CLIENTE_ID, SQLEngineQuery.OperatorCriteria.Igual, socioID, SortOrder.Ascending)
        Me.SearchResult.AddRange(movimientos.SearchResult)
        Return movimientos.SearchResult.Count
    End Function

    ''' <summary>
    ''' Devolver el saldo de la cuenta corriente a la fecha correspondiente
    ''' </summary>
    ''' <param name="socioID">ID del socio</param>
    ''' <param name="aFecha">Fecha a calcular el saldo</param>
    ''' <returns>El saldo de la cuenta corriente</returns>
    Public Function GetSaldo(ByVal socioID As Integer, ByVal aFecha As Integer) As Decimal
        Dim tmpSql As New SQLEngine
        HelperMod.InjectSqle(Me.sqle, tmpSql)
        With tmpSql.Query
            .Reset()
            .TableName = MovimientoCuentaCorrienteSocio.TABLA.TABLA_NOMBRE
            .AddSelectColumn($"SUM({MovimientoCuentaCorrienteSocio.TABLA.IMPORTE})")
            .WHEREstring = $"{MovimientoCuentaCorrienteSocio.TABLA.CLIENTE_ID} = { .p(socioID)} 
                              AND {MovimientoCuentaCorrienteSocio.TABLA.DELETED} = { .p(False)}
                              AND {MovimientoCuentaCorrienteSocio.TABLA.FECHA_INGRESO} <= { .p(aFecha)}"
            If .Query() Then
                .QueryRead()
                Return CDec(If(.GetQueryData(0).GetType.Name = "String", 0, .GetQueryData(0)))
            Else
                Err.Raise(513, "SocCam", "No se pudo acceder a la base de datos")
            End If
        End With
    End Function

    Public Function EliminarMovimiento(ByVal movimientoID As Integer, Optional ByVal noPrompt As Boolean = False) As Integer
        Dim mov As New MovimientoCuentaCorrienteSocio(Me.sqle)
        mov.LoadMe(movimientoID)
        mov.LoadDetalles()

        Dim det As String = ""
        For Each itm As ProductoSocio In mov.Productos
            det &= $"· {itm.Descripcion}{vbCrLf}"
        Next

        For Each itm As CuotaSocio In mov.CuotasSociales
            det &= $"· Cuota social {Utils.GetNombreMes(itm.Periodo)} {itm.anio}{vbCrLf}"
        Next

        Dim msg As String = ""
        Dim question As String = "¿Eliminar los siguientes items?"
        Dim accion As Integer = 0

        Select Case mov.ComprobanteTipo
            Case MovimientoCuentaCorrienteSocio.TIPO.FACTURA_C
                msg = "Además se va a crear una NCC por el mismo valor de la factura"
                accion = 1
            Case MovimientoCuentaCorrienteSocio.TIPO.RECIBO_X
                msg = "¿Continuar?"
                accion = 2
            Case MovimientoCuentaCorrienteSocio.TIPO.TICKET_X
                question = "Se va a cambiar el estado a IMPAGO de"
                Dim tmpMov As New MovimientoCuentaCorrienteSocio(Me.sqle)
                tmpMov.LoadMe(mov.ComprobanteRelacionado)
                tmpMov.LoadDetalles()

                For Each itm As ProductoSocio In tmpMov.Productos
                    det &= $"· {itm.Descripcion}{vbCrLf}"
                Next

                For Each itm As CuotaSocio In tmpMov.CuotasSociales
                    det &= $"· Cuota social {Utils.GetNombreMes(itm.Periodo)} {itm.anio}{vbCrLf}"
                Next

                msg = "¿Continuar?"
                accion = 3
            Case MovimientoCuentaCorrienteSocio.TIPO.NOTA_CREDITO_C
                msg = "Además se va a crear una NDC por el mismo valor de la nota de crédito"
            Case MovimientoCuentaCorrienteSocio.TIPO.NOTA_CREDITO_X
                msg = "Además se va a crear una NDX por el mismo valor de la nota de crédito"
        End Select

        If accion = 0 Then
            MsgBox("Seleccione un movimiento válido", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Return -1
        End If

        If Not noPrompt Then
            If MsgBox($"{question}{vbCrLf}{vbCrLf}{det}{vbCrLf}{msg}.", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Return -2
            End If
        End If


        Select Case accion
            Case 1, 2
                If EliminarCuotasSociales(mov.CuotasSociales) <> 0 Then
                    Utils.Scream("Hubo un error al eliminar las cuotas. Vuelva a intentarlo más tarde")
                    Return -1
                End If

                If EliminarProductos(mov.Productos) <> 0 Then
                    Utils.Scream("Hubo un error al eliminar los productos. Vuelva a intentarlo más tarde")
                    Return -2
                End If

                If accion = 1 Then
                    Dim comprobante As New AfipFactura(Me.sqle)
                    comprobante.LoadMe(mov.ComprobanteRelacionado)

                    If GenerarNotaCredito(mov,
                                           If(mov.ComprobanteTipo = MovimientoCuentaCorrienteSocio.TIPO.FACTURA_C, True, False),
                                           Math.Abs(mov.Importe),
                                           "Error en ",
                                           mov.ClienteId,
                                           Utils.Int8601ToDate(comprobante.FechaServicioDesde),
                                           Utils.Int8601ToDate(comprobante.FechaServicioHasta)) <> 0 Then
                        Return -3
                    End If

                    mov.Procesado = True
                    mov.Save(MovimientoCuentaCorrienteSocio.Guardar.EDITAR)
                Else
                    mov.Delete()
                End If
            Case 3
                Dim movDebe As New MovimientoCuentaCorrienteSocio(Me.sqle)
                movDebe.LoadMe(mov.ComprobanteRelacionado)
                movDebe.LoadDetalles()

                If DarImpagoCuotasSociales(movDebe.CuotasSociales) <> 0 Then
                    Utils.Scream("Hubo un error al eliminar las cuotas. Vuelva a intentarlo más tarde")
                    Return -4
                End If

                If DarImpagosProductos(movDebe.Productos) <> 0 Then
                    Utils.Scream("Hubo un error al eliminar los productos. Vuelva a intentarlo más tarde")
                    Return -5
                End If



                movDebe.Procesado = False
                movDebe.Save(MovimientoCuentaCorrienteSocio.Guardar.EDITAR)

                Dim servicio As New PagoServicio()
                HelperMod.InjectSqle(Me.sqle, servicio.sqle)
                servicio.LoadByMovimiento(mov.Id)

                For Each srv As PagoServicio In servicio.SearchResult
                    HelperMod.InjectSqle(Me.sqle, srv.sqle)
                    If Not srv.Delete() Then Return -6
                Next

                mov.Delete()
        End Select

        Return 0
    End Function

    Private Function EliminarCuotasSociales(ByVal lstCuotas As List(Of CuotaSocio)) As Integer
        For Each cuota As CuotaSocio In lstCuotas
            If Not cuota.Delete(Me.sqle, cuota.id) Then
                Return -1
            End If
        Next
        Return 0
    End Function

    Private Function DarImpagoCuotasSociales(ByVal lstCuotas As List(Of CuotaSocio)) As Integer
        For Each cuota As CuotaSocio In lstCuotas
            cuota.estado = CuotaSocio.ESTADO_SOCIO.PENDIENTE
            If cuota.Save(Me.sqle, 1) = 0 Then
                Return -1
            End If
        Next
        Return 0
    End Function

    Private Function EliminarProductos(ByVal lstProductos As List(Of ProductoSocio)) As Integer
        For Each producto As ProductoSocio In lstProductos
            If Not producto.Delete() Then
                Return -1
            End If
        Next
        Return 0
    End Function

    Private Function DarImpagosProductos(ByVal lstProductos As List(Of ProductoSocio)) As Integer
        For Each producto As ProductoSocio In lstProductos
            producto.FechaPago = 0

            If Not producto.Save(ProductoSocio.Guardar.EDITAR) Then
                Return -1
            End If
        Next
        Return 0
    End Function




    Public Function GenerarNotaCredito(ByVal mov As MovimientoCuentaCorrienteSocio,
                                       ByVal facturar As Boolean,
                                       ByVal importe As Decimal,
                                       ByVal detalle As String,
                                       ByVal clienteID As Integer,
                                       ByVal facturadoDesde As Date,
                                       ByVal facturadoHasta As Date
                                       ) As Integer
        Dim movimiento As New MovimientoCuentaCorrienteSocio(frmMain.kernel.sqle)
        movimiento.Importe = importe
        movimiento.ImporteCobrar = movimiento.Importe
        movimiento.ComprobanteTipo = If(facturar, MovimientoCuentaCorrienteSocio.TIPO.NOTA_CREDITO_C, MovimientoCuentaCorrienteSocio.TIPO.NOTA_CREDITO_X)
        movimiento.Procesado = True
        movimiento.FechaIngreso = Utils.DateTo8601(Now.Date)
        movimiento.ClienteId = clienteID



        Dim lstDetalles As New List(Of AfipFacturaDetalle)
        Dim det As New AfipFacturaDetalle(frmMain.kernel.sqle)
        det.Codigo = 0
        det.ProductoServicio = detalle
        det.Cantidad = 1
        det.UnidadMedida = AfipFacturaDetalle.Unidad.OTRAS_UNIDADES
        det.PrecioUnitario = importe

        lstDetalles.Add(det)

        Dim comprobante As New AfipFactura(Me.sqle)
        comprobante.LoadMe(mov.ComprobanteRelacionado)

        Dim numeroComprobante As Integer = 0
        Dim idComprobante As Integer = frmMain.kernel.ctrlSociosCC.GenerarComprobante(If(facturar, AfipFactura.Tipo.NOTA_CREDITO_C, AfipFactura.Tipo.NOTA_CREDITO_X),
                                                                                      clienteID,
                                                                                      lstDetalles,
                                                                                      frmMain.kernel.globalConfig,
                                                                                      facturadoDesde,
                                                                                      facturadoHasta,,
                                                                                      numeroComprobante,
                                                                                      comprobante)

        If idComprobante <= 0 Then
            Utils.Scream("No se pudo guardar el comprobante. Vuelva a intentar.")
            Return -1
        End If

        movimiento.ComprobanteRelacionado = idComprobante
        movimiento.Save(MovimientoCuentaCorrienteSocio.Guardar.NUEVO)

        Return 0
    End Function

    ''' <summary>
    ''' Genera las cuotas sociales por la campaña
    ''' </summary>
    ''' <param name="detalle">El detalle con las cuotas sociales</param>
    ''' <returns>0 Si se procesó correctamente, entero negativo si no</returns>
    Public Function GenerarCuotasSocialesCampania(ByVal detalle As DetalleCCSocio, Optional ByVal idMovimiento As Integer = 0) As Integer
        Dim errorCode As Integer = 0
        If detalle.Tipo = DetalleCCSocio.TipoDeMovimiento.SOCIOS_CUOTA_SOCIAL_CAMPANIA Then
            If detalle.ListadoCuotasSocialesVirtuales.Count > 0 Then
                Dim totalCuotasGeneradas As Integer = 0
                For Each cuota As CuotaSocio In detalle.ListadoCuotasSocialesVirtuales.Values
                    If cuota.monto = 0 Then
                        cuota.estado = CuotaSocio.ESTADO_SOCIO.AL_DIA
                        cuota.fechaPago = Now
                    Else
                        cuota.estado = CuotaSocio.ESTADO_SOCIO.PENDIENTE
                    End If
                    cuota.cobradorID = 1
                    cuota.MovimientoCC = idMovimiento
                    If cuota.Campania < 0 Then
                        cuota.observaciones = $"Cuota bonificada - {detalle.Campania.Nombre}"
                    Else
                        cuota.observaciones = $"Cuota bonificada - Campaña {detalle.Campania.Nombre}"
                    End If


                    If cuota.Save(Me.sqle, 0) = 0 Then
                        errorCode = -3
                    End If
                    totalCuotasGeneradas += 1
                Next

                If (errorCode = 0) And (totalCuotasGeneradas <> detalle.ListadoCuotasSocialesVirtuales.Count) Then
                    ' NO GENERO LA MISMA CANTIDAD DE CUOTAS
                    errorCode = -5
                End If

                If errorCode <> 0 Then
                    For Each cuota As CuotaSocio In detalle.ListadoCuotasSocialesVirtuales.Values
                        cuota.Delete(Me.sqle, cuota.id)
                    Next
                End If
            Else
                ' NO HAY CUOTAS EN LA LISTA
                errorCode = -2
            End If
        Else
            ' NO ES CORRECTO EL TIPO DE DETALLE
            errorCode = -1
        End If

        Return errorCode
    End Function

    ''' <summary>
    ''' Genera una cuota social
    ''' </summary>
    ''' <param name="detalle">Detalle</param>
    ''' <param name="idMovimiento"></param>
    ''' <returns>El ID de la cuota</returns>
    Public Function GenerarCuotaSocial(ByVal detalle As DetalleCCSocio, ByVal socioID As Integer, Optional ByVal idMovimiento As Integer = 0) As Integer
        Dim cuota As New CuotaSocio()
        If detalle.ListadoCuotasSocialesVirtuales.Count > 0 Then
            cuota = detalle.ListadoCuotasSocialesVirtuales(detalle.IdCuota)
            cuota.anio = detalle.ListadoCuotasSocialesVirtuales(detalle.IdCuota).anio
            cuota.cobradorID = 1
            cuota.estado = detalle.ListadoCuotasSocialesVirtuales(detalle.IdCuota).estado
            cuota.fechaPago = detalle.ListadoCuotasSocialesVirtuales(detalle.IdCuota).fechaPago
            cuota.monto = detalle.ListadoCuotasSocialesVirtuales(detalle.IdCuota).monto
            cuota.MovimientoCC = idMovimiento
            cuota.observaciones = detalle.ListadoCuotasSocialesVirtuales(detalle.IdCuota).observaciones
            cuota.Operacion = detalle.ListadoCuotasSocialesVirtuales(detalle.IdCuota).Operacion
            cuota.Periodicidad = detalle.ListadoCuotasSocialesVirtuales(detalle.IdCuota).Periodicidad
            cuota.Periodo = detalle.ListadoCuotasSocialesVirtuales(detalle.IdCuota).Periodo
            cuota.PlanID = detalle.ListadoCuotasSocialesVirtuales(detalle.IdCuota).PlanID
            cuota.socioID = socioID

            If Not cuota.CuotaExist(Me.sqle, cuota.Periodo, cuota.Periodicidad, cuota.anio, cuota.socioID) Then
                Return cuota.Save(Me.sqle, 0)
            End If
        End If

        Return -1
    End Function

    ''' <summary>
    ''' Carga un producto de socio según el detalle
    ''' </summary>
    ''' <param name="detalle">Detalle del producto</param>
    ''' <param name="socioID">ID del socio a guardar</param>
    ''' <param name="idMovimiento">ID del movimiento de CC que lo genera</param>
    ''' <returns></returns>
    Public Function CargarProductoSocio(ByVal detalle As DetalleCCSocio, ByVal socioID As Integer, Optional ByVal idMovimiento As Integer = 0) As Integer
        Dim prod As New ProductoSocio(Me.sqle)
        prod.SocioId = socioID
        prod.Descripcion = detalle.Descripcion
        Select Case detalle.Tipo
            Case DetalleCCSocio.TipoDeMovimiento.SOCIOS_OTROS
                prod.Tipo = ProductoSocio.Producto_tipo.OTROS
            Case DetalleCCSocio.TipoDeMovimiento.SOCIOS_PUBLICIDAD
                prod.Tipo = ProductoSocio.Producto_tipo.PUBLICIDAD
            Case DetalleCCSocio.TipoDeMovimiento.SOCIOS_BOLSIN
                prod.Tipo = ProductoSocio.Producto_tipo.BOLSIN
            Case DetalleCCSocio.TipoDeMovimiento.SOCIOS_MEDICINA
                prod.Tipo = ProductoSocio.Producto_tipo.MEDICINA_LABORAL
        End Select

        prod.Importe = detalle.Importe
        prod.Movimiento_cc = idMovimiento

        If prod.Save(ProductoSocio.Guardar.NUEVO) Then
            Return prod.Id
        End If

        Return -1
    End Function

    Public Function FacturarMovimiento(ByVal mov As MovimientoCuentaCorrienteSocio) As Integer
        Dim rec As New AfipFactura(Me.sqle)
        Dim fx As New AfipFacturaEX(Me.sqle)
        If Not rec.LoadMe(mov.ComprobanteRelacionado, fx) Then Return -1

        If MsgBox($"¿Reemplazar {rec.ToString} por una factura aprobada por AFIP?{vbCrLf}Importe a facturar: {Utils.ToMoneyFormat(rec.ImporteTotal)}", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim factura As New AfipFactura(Me.sqle)
            factura = rec
            factura.ComprobanteTipo = AfipFactura.Tipo.FACTURA_C
            factura.FechaEmision = Utils.DateTo8601(Now.Date)
            factura.Numero = factura.GetUltimoNumero + 1
            factura.FechaVencimiento = Utils.DateTo8601(Now.Date.AddDays(1))
            factura.FechaVencimientoPago = Utils.DateTo8601(Now.Date.AddDays(1))

            If Not factura.Autorizar() Then Return -2
            For Each det As AfipFacturaDetalle In factura.Detalles
                det = Utils.LimpiarDetalleFactura(det)
            Next
            factura.Save(GUARDAR_NUEVO)

            fx.FacturaId = factura.Id
            fx.Save(GUARDAR_NUEVO)


            Dim movimiento As New MovimientoCuentaCorrienteSocio(Me.sqle)
            movimiento.LoadMe(mov.Id)
            movimiento.LoadDetalles()
            movimiento.ComprobanteRelacionado = factura.Id
            movimiento.ComprobanteTipo = MovimientoCuentaCorrienteSocio.TIPO.FACTURA_C
            Try
                If Not movimiento.Save(GUARDAR_EDITAR) Then Return -3
                'movimiento.Save(GUARDAR_EDITAR)
            Catch ex As Exception
                Console.WriteLine("Error al intentar guardar el movimiento: " & ex.Message)
            End Try



            For Each cuota As CuotaSocio In movimiento.CuotasSociales
                cuota.observaciones = $"FC-{factura.Numero}"
                cuota.fechaPago = Now.Date
                cuota.estado = CuotaSocio.ESTADO_SOCIO.AL_DIA
                cuota.Save(Me.sqle, GUARDAR_EDITAR) '1
            Next

            Return 0
        End If

        Return 0
    End Function

    Public Function BonificacionesMasivas(ByVal nombreCampania As String,
                                          ByVal porcentajeDescuento As Decimal,
                                          ByVal cantidadCuotas As Integer,
                                          ByVal incluirCofres As Boolean,
                                          ByVal periodoAplicacion As Integer,
                                          ByVal anioAplicacion As Integer,
                                          ByVal listaBonificados As List(Of Integer)(),
                                          ByVal facturar As Boolean
                                          ) As Integer

        If listaBonificados.Count > 0 Then
            For i = 0 To 2
                For Each segmento In listaBonificados(i)
                    Dim sociosSegmento As New SocioNT(Me.sqle)
                    sociosSegmento.LoadAll(i + 1, segmento)

                    For Each socio As SocioNT In sociosSegmento.SearchResult
                        HelperMod.InjectSqle(Me.sqle, socio.Sqle)
                        If socio.TieneCofreSeguridad() And Not incluirCofres Then Continue For

                        Dim campania As New Campania(Me.sqle)
                        Dim detalleCC As New DetalleCCSocio


                        campania.Id = 1
                        campania.CuotasBonificadas = cantidadCuotas
                        campania.Descripcion = nombreCampania
                        campania.Nombre = nombreCampania
                        campania.PorcentajeBonificado = porcentajeDescuento
                        Dim importeCuotas As Decimal = campania.CargarListaDeCuotas(socio.Id)

                        detalleCC.Campania = campania
                        detalleCC.Descripcion = nombreCampania
                        detalleCC.Importe = importeCuotas
                        detalleCC.Tipo = DetalleCCSocio.TipoDeMovimiento.SOCIOS_CUOTA_SOCIAL_CAMPANIA

                        Dim facturadoDesde As Date
                        Dim facturadoHasta As Date

                        Dim lstDetalles As New List(Of AfipFacturaDetalle)

                        Dim movimiento As New MovimientoCuentaCorrienteSocio(frmMain.kernel.sqle)

                        Dim detalle As New AfipFacturaDetalle(frmMain.kernel.sqle)
                        If facturar Then
                            detalle.ProductoServicio = $"{nombreCampania}"
                        Else
                            detalle.ProductoServicio = $"{nombreCampania}: <b>{Utils.ToMoneyFormat(importeCuotas)}</b>"
                        End If

                        detalle.Cantidad = campania.CuotasBonificadas
                        detalle.UnidadMedida = AfipFacturaDetalle.Unidad.OTRAS_UNIDADES
                        detalle.PrecioUnitario = importeCuotas / campania.CuotasBonificadas
                        detalle.Codigo = 0
                        detalle.BonificacionPercent = campania.PorcentajeBonificado


                        lstDetalles.Add(detalle)

                        Dim numeroComprobante As Integer = 0
                        Dim idComprobante As Integer = frmMain.kernel.ctrlSociosCC.GenerarComprobante(If(facturar, AfipFactura.Tipo.FACTURA_C, AfipFactura.Tipo.RECIBO), socio.Id, lstDetalles, frmMain.kernel.globalConfig, facturadoDesde, facturadoHasta, , numeroComprobante)

                        If idComprobante <= 0 Then
                            Utils.Scream("No se pudo guardar el comprobante. Vuelva a intentar.")
                            Return -1
                        End If

                        movimiento.ClienteId = socio.Id
                        movimiento.ComprobanteRelacionado = idComprobante
                        movimiento.FechaIngreso = Utils.DateTo8601(Now.Date)
                        movimiento.Importe = (0 * -1)
                        movimiento.ImporteCobrar = movimiento.Importe
                        movimiento.ComprobanteTipo = If(facturar, MovimientoCuentaCorrienteSocio.TIPO.FACTURA_C, MovimientoCuentaCorrienteSocio.TIPO.RECIBO_X)


                        movimiento.Save(MovimientoCuentaCorrienteSocio.Guardar.NUEVO)


                        For Each cuota As CuotaSocio In campania.ListaCuotasCampaña
                            cuota.Campania = -1
                            detalleCC.ListadoCuotasSocialesVirtuales.Add(Utils.ComponerIso8601(0, cuota.Periodo, cuota.anio), cuota)
                        Next

                        Me.GenerarCuotasSocialesCampania(detalleCC, movimiento.Id)
                    Next
                Next
            Next

        End If


        Return 0
    End Function

End Class
