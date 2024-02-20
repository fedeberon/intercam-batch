Imports helix
Public Class AfipFactura

    Public Enum Tipo As Integer
        TODOS = 0
        FACTURA_A = 1
        NOTA_DEBITO_A = 2
        NOTA_CREDITO_A = 3
        RECIBO_A = 4
        NOTA_VENTA_CONTADO_A = 5
        FACTURA_B = 6
        NOTA_DEBITO_B = 7
        NOTA_CREDITO_B = 8
        RECIBO_B = 9
        NOTA_VENTA_CONTADO_B = 10
        FACTURA_C = 11
        NOTA_DEBITO_C = 12
        NOTA_CREDITO_C = 13
        RECIBO_C = 15
        NOTA_VENTA_CONTADO_C = 16
        COMPROBANTES_COMPRA_BIENES_USADOS = 30
        OTROS_COMPROBANTES_CUMPLEN_1415 = 30
        OTROS_COMPROBANTES_NO_CUMPLEN_1415 = 99
        RECIBO = 999
        NOTA_CREDITO_X = 1000
        NOTA_DEBITO_X = 1001
    End Enum

    Public Enum Concepto As Integer
        PRODUCTOS = 1
        SERVICIOS = 2
        PRODUCTOS_Y_SERVICIOS = 3
    End Enum

    Public Enum Estado As Integer
        APROBADO = 0
        RECHAZADO = 1
        PARCIAL = 2
        SIN_SOLICITAR = 3
    End Enum

    Public Enum Documento As Integer
        CUIT = 80
        CUIL = 86
        CDI = 87
        LE = 89
        LC = 90
        CI_EXTRANJERA = 91
        EN_TRAMITE = 92
        ACTA_NACIMIENTO = 93
        PASPORTE = 94
        CI_BA_RNP = 95
        DNI = 96
        VENTA_GLOBAL_DIARIA = 99
        CERTIFICADO_MIGRACION = 30
        USADO_ANSES_PADRON = 88
    End Enum

    Public Const MONEDA_OTRA As String = "000"
    Public Const MONEDA_PESO As String = "PES"
    Public Const MONEDA_DOLAR As String = "DOL"

    Public Property Sqle As New SQLEngine
    Public Property SearchResult As New List(Of AfipFactura)

    Public Property Id As Long = 0
    Public Property CuitEmisor As Long = 0
    Public Property ComprobanteTipo As Tipo = Tipo.FACTURA_C
    Public Property PuntoVenta As Integer = 1
    Public Property Numero As Integer = 0
    Public Property Concept As Concepto = Concepto.PRODUCTOS
    Public Property DocumentoTipo As Documento = Documento.CUIT
    Public Property DocumentoNumero As Long = 0
    Public Property FechaEmision As Long = 0
    Public Property ImporteTotal As Decimal = 0
    Public Property ImporteTotalConc As Decimal = 0
    Public Property ImporteNeto As Decimal = 0
    Public Property ImporteOpExcento As Decimal = 0
    Public Property ImporteIva As Decimal = 0
    Public Property ImporteTributo As Decimal = 0
    Public Property FechaServicioDesde As Long = 0
    Public Property FechaServicioHasta As Long = 0
    Public Property FechaVencimientoPago As Long = 0
    Public Property Moneda As String = MONEDA_PESO
    Public Property MonedaCotizacion As Decimal = 1
    Public Property Homologacion As Boolean = False
    Public Property EstadoSolicitud As Estado = Estado.SIN_SOLICITAR
    Public Property MensajeAfip As String = ""
    Public Property Cae As String = ""
    Public Property FechaVencimiento As Integer = 0
    Public Property Deleted As Boolean = False
    Public Property Modificado As Date = Now
    Public Property CbtesAsoc As New List(Of AfipFactura)


    Public Property Detalles As New List(Of AfipFacturaDetalle)


    Public Enum Guardar
        NUEVO = 0
        EDITAR = 1
    End Enum


    Public Structure TABLA
        Const TABLA_NOMBRE As String = "AfipFactura"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const CUIT_EMISOR As String = TABLA_NOMBRE & "_cuitEmisor"
        Const COMPROBANTE_TIPO As String = TABLA_NOMBRE & "_comprobanteTipo"
        Const PUNTO_VENTA As String = TABLA_NOMBRE & "_puntoVenta"
        Const NUMERO As String = TABLA_NOMBRE & "_numero"
        Const CONCEPT As String = TABLA_NOMBRE & "_concept"
        Const DOCUMENTO_TIPO As String = TABLA_NOMBRE & "_documentoTipo"
        Const DOCUMENTO_NUMERO As String = TABLA_NOMBRE & "_documentoNumero"
        Const FECHA_EMISION As String = TABLA_NOMBRE & "_fechaEmision"
        Const IMPORTE_TOTAL As String = TABLA_NOMBRE & "_importeTotal"
        Const IMPORTE_TOTAL_CONC As String = TABLA_NOMBRE & "_importeTotalConc"
        Const IMPORTE_NETO As String = TABLA_NOMBRE & "_importeNeto"
        Const IMPORTE_OP_EXCENTO As String = TABLA_NOMBRE & "_importeOpExcento"
        Const IMPORTE_IVA As String = TABLA_NOMBRE & "_importeIva"
        Const IMPORTE_TRIBUTO As String = TABLA_NOMBRE & "_importeTributo"
        Const FECHA_SERVICIO_DESDE As String = TABLA_NOMBRE & "_fechaServicioDesde"
        Const FECHA_SERVICIO_HASTA As String = TABLA_NOMBRE & "_fechaServicioHasta"
        Const FECHA_VENCIMIENTO_PAGO As String = TABLA_NOMBRE & "_fechaVencimientoPago"
        Const MONEDA As String = TABLA_NOMBRE & "_moneda"
        Const MONEDA_COTIZACION As String = TABLA_NOMBRE & "_monedaCotizacion"
        Const HOMOLOGACION As String = TABLA_NOMBRE & "_homologacion"
        Const ESTADO As String = TABLA_NOMBRE & "_estado"
        Const MENSAJE_AFIP As String = TABLA_NOMBRE & "_mensajeAfip"
        Const CAE As String = TABLA_NOMBRE & "_cae"
        Const FECHA_VENCIMIENTO As String = TABLA_NOMBRE & "_fechaVencimiento"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & CUIT_EMISOR & ", " & COMPROBANTE_TIPO & ", " & PUNTO_VENTA & ", " & NUMERO & ", " & CONCEPT & ", " & DOCUMENTO_TIPO & ", " & DOCUMENTO_NUMERO & ", " & FECHA_EMISION & ", " & IMPORTE_TOTAL & ", " & IMPORTE_TOTAL_CONC & ", " & IMPORTE_NETO & ", " & IMPORTE_OP_EXCENTO & ", " & IMPORTE_IVA & ", " & IMPORTE_TRIBUTO & ", " & FECHA_SERVICIO_DESDE & ", " & FECHA_SERVICIO_HASTA & ", " & FECHA_VENCIMIENTO_PAGO & ", " & MONEDA & ", " & MONEDA_COTIZACION & ", " & HOMOLOGACION & ", " & ESTADO & ", " & MENSAJE_AFIP & ", " & CAE & ", " & FECHA_VENCIMIENTO & ", " & DELETED & ", " & MODIFICADO
    End Structure



    Public Sub New()
    End Sub


    Public Sub New(ByVal iSqle As SQLEngine)
        Me.Sqle.RequireCredentials = iSqle.RequireCredentials
        Me.Sqle.Username = iSqle.Username
        Me.Sqle.Password = iSqle.Password
        Me.Sqle.dbType = iSqle.dbType
        Me.Sqle.Path = iSqle.Path
        Me.Sqle.DatabaseName = iSqle.DatabaseName
        If iSqle.IsStarted Then
            Me.Sqle.ColdBoot()
        Else
            Me.Sqle.Start()
        End If
    End Sub

    Public Overrides Function ToString() As String
        Dim tipoString As String
        Select Case Me.ComprobanteTipo
            Case Tipo.FACTURA_C
                tipoString = "FC"
            Case Tipo.RECIBO
                tipoString = "RX"
            Case Tipo.RECIBO_C
                tipoString = "RC"
            Case Tipo.NOTA_CREDITO_C
                tipoString = "NCC"
            Case Tipo.NOTA_DEBITO_C
                tipoString = "NDC"
            Case Tipo.NOTA_CREDITO_X
                tipoString = "NCX"
            Case Tipo.NOTA_DEBITO_X
                tipoString = "NDX"
            Case Else
                tipoString = "CMP"
        End Select

        Return tipoString & " " & Utils.ComponerNumeroComprobante(Me.PuntoVenta, Me.Numero)
    End Function


    Public Function LoadMe(ByVal myID As Integer) As Boolean
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .SimpleSearch(TABLA.ID, SQLEngineQuery.OperatorCriteria.Igual, myID)
            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    Id = CLng(.GetQueryData(TABLA.ID))
                    CuitEmisor = CLng(.GetQueryData(TABLA.CUIT_EMISOR))
                    ComprobanteTipo = CInt(.GetQueryData(TABLA.COMPROBANTE_TIPO))
                    PuntoVenta = CInt(.GetQueryData(TABLA.PUNTO_VENTA))
                    Numero = CInt(.GetQueryData(TABLA.NUMERO))
                    Concept = CInt(.GetQueryData(TABLA.CONCEPT))
                    DocumentoTipo = CInt(.GetQueryData(TABLA.DOCUMENTO_TIPO))
                    DocumentoNumero = CLng(.GetQueryData(TABLA.DOCUMENTO_NUMERO))
                    FechaEmision = CLng(.GetQueryData(TABLA.FECHA_EMISION))
                    ImporteTotal = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL))
                    ImporteTotalConc = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL_CONC))
                    ImporteNeto = CDec(.GetQueryData(TABLA.IMPORTE_NETO))
                    ImporteOpExcento = CDec(.GetQueryData(TABLA.IMPORTE_OP_EXCENTO))
                    ImporteIva = CDec(.GetQueryData(TABLA.IMPORTE_IVA))
                    ImporteTributo = CDec(.GetQueryData(TABLA.IMPORTE_TRIBUTO))
                    FechaServicioDesde = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_DESDE))
                    FechaServicioHasta = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_HASTA))
                    FechaVencimientoPago = CLng(.GetQueryData(TABLA.FECHA_VENCIMIENTO_PAGO))
                    Moneda = CStr(.GetQueryData(TABLA.MONEDA))
                    MonedaCotizacion = CDec(.GetQueryData(TABLA.MONEDA_COTIZACION))
                    Homologacion = CBool(.GetQueryData(TABLA.HOMOLOGACION))
                    EstadoSolicitud = CInt(.GetQueryData(TABLA.ESTADO))
                    MensajeAfip = CStr(.GetQueryData(TABLA.MENSAJE_AFIP))
                    Cae = CStr(.GetQueryData(TABLA.CAE))
                    FechaVencimiento = CInt(.GetQueryData(TABLA.FECHA_VENCIMIENTO))
                    Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))


                    Dim tmpDet As New AfipFacturaDetalle(Sqle)
                    tmpDet.QuickSearch(AfipFacturaDetalle.TABLA.FACTURA_ID, SQLEngineQuery.OperatorCriteria.Igual, Id)
                    Detalles.Clear()
                    Detalles.AddRange(tmpDet.SearchResult)


                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End With
    End Function

    Public Function LoadMe(ByVal factPVenta As Integer, ByVal factNumero As Long, ByRef fx As AfipFacturaEX) As Boolean
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .WHEREstring = $"{TABLA.PUNTO_VENTA} = ? AND {TABLA.NUMERO} = ? AND {TABLA.HOMOLOGACION} = ?  AND {TABLA.COMPROBANTE_TIPO} = ?"
            .AddWHEREparam(factPVenta)
            .AddWHEREparam(factNumero)
            .AddWHEREparam(Me.Homologacion)
            .AddWHEREparam(Tipo.FACTURA_C)
            If .Query Then
                If .RecordCount = 1 Then
                    .QueryRead()
                    Id = CLng(.GetQueryData(TABLA.ID))
                    CuitEmisor = CLng(.GetQueryData(TABLA.CUIT_EMISOR))
                    ComprobanteTipo = CInt(.GetQueryData(TABLA.COMPROBANTE_TIPO))
                    PuntoVenta = CInt(.GetQueryData(TABLA.PUNTO_VENTA))
                    Numero = CInt(.GetQueryData(TABLA.NUMERO))
                    Concept = CInt(.GetQueryData(TABLA.CONCEPT))
                    DocumentoTipo = CInt(.GetQueryData(TABLA.DOCUMENTO_TIPO))
                    DocumentoNumero = CLng(.GetQueryData(TABLA.DOCUMENTO_NUMERO))
                    FechaEmision = CLng(.GetQueryData(TABLA.FECHA_EMISION))
                    ImporteTotal = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL))
                    ImporteTotalConc = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL_CONC))
                    ImporteNeto = CDec(.GetQueryData(TABLA.IMPORTE_NETO))
                    ImporteOpExcento = CDec(.GetQueryData(TABLA.IMPORTE_OP_EXCENTO))
                    ImporteIva = CDec(.GetQueryData(TABLA.IMPORTE_IVA))
                    ImporteTributo = CDec(.GetQueryData(TABLA.IMPORTE_TRIBUTO))
                    FechaServicioDesde = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_DESDE))
                    FechaServicioHasta = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_HASTA))
                    FechaVencimientoPago = CLng(.GetQueryData(TABLA.FECHA_VENCIMIENTO_PAGO))
                    Moneda = CStr(.GetQueryData(TABLA.MONEDA))
                    MonedaCotizacion = CDec(.GetQueryData(TABLA.MONEDA_COTIZACION))
                    Homologacion = CBool(.GetQueryData(TABLA.HOMOLOGACION))
                    EstadoSolicitud = CInt(.GetQueryData(TABLA.ESTADO))
                    MensajeAfip = CStr(.GetQueryData(TABLA.MENSAJE_AFIP))
                    Cae = CStr(.GetQueryData(TABLA.CAE))
                    FechaVencimiento = CInt(.GetQueryData(TABLA.FECHA_VENCIMIENTO))
                    Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))


                    Dim tmpDet As New AfipFacturaDetalle(Sqle)
                    tmpDet.QuickSearch(AfipFacturaDetalle.TABLA.FACTURA_ID, SQLEngineQuery.OperatorCriteria.Igual, Id)
                    Detalles.Clear()
                    Detalles.AddRange(tmpDet.SearchResult)

                    fx.LoadMe(Me.Id, True)

                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End With
    End Function


    Public Function LoadMe(ByVal factPVenta As Integer, ByVal factNumero As Long, Optional tipoComprobante As AfipFactura.Tipo = AfipFactura.Tipo.FACTURA_C) As Boolean
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)

            .WHEREstring = $"{TABLA.PUNTO_VENTA} = ? AND {TABLA.NUMERO} = ? AND {TABLA.HOMOLOGACION} = ? AND {TABLA.COMPROBANTE_TIPO} = ?"
            .AddWHEREparam(factPVenta)
            .AddWHEREparam(factNumero)
            'If Not frmMain.kernel.DEBUG_MODE Then .AddWHEREparam(Not frmMain.kernel.globalConfig.Produccion)
            .AddWHEREparam(tipoComprobante)

            If .Query Then
                If .RecordCount = 1 Then
                    .QueryRead()
                    Id = CLng(.GetQueryData(TABLA.ID))
                    CuitEmisor = CLng(.GetQueryData(TABLA.CUIT_EMISOR))
                    ComprobanteTipo = CInt(.GetQueryData(TABLA.COMPROBANTE_TIPO))
                    PuntoVenta = CInt(.GetQueryData(TABLA.PUNTO_VENTA))
                    Numero = CInt(.GetQueryData(TABLA.NUMERO))
                    Concept = CInt(.GetQueryData(TABLA.CONCEPT))
                    DocumentoTipo = CInt(.GetQueryData(TABLA.DOCUMENTO_TIPO))
                    DocumentoNumero = CLng(.GetQueryData(TABLA.DOCUMENTO_NUMERO))
                    FechaEmision = CLng(.GetQueryData(TABLA.FECHA_EMISION))
                    ImporteTotal = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL))
                    ImporteTotalConc = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL_CONC))
                    ImporteNeto = CDec(.GetQueryData(TABLA.IMPORTE_NETO))
                    ImporteOpExcento = CDec(.GetQueryData(TABLA.IMPORTE_OP_EXCENTO))
                    ImporteIva = CDec(.GetQueryData(TABLA.IMPORTE_IVA))
                    ImporteTributo = CDec(.GetQueryData(TABLA.IMPORTE_TRIBUTO))
                    FechaServicioDesde = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_DESDE))
                    FechaServicioHasta = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_HASTA))
                    FechaVencimientoPago = CLng(.GetQueryData(TABLA.FECHA_VENCIMIENTO_PAGO))
                    Moneda = CStr(.GetQueryData(TABLA.MONEDA))
                    MonedaCotizacion = CDec(.GetQueryData(TABLA.MONEDA_COTIZACION))
                    Homologacion = CBool(.GetQueryData(TABLA.HOMOLOGACION))
                    EstadoSolicitud = CInt(.GetQueryData(TABLA.ESTADO))
                    MensajeAfip = CStr(.GetQueryData(TABLA.MENSAJE_AFIP))
                    Cae = CStr(.GetQueryData(TABLA.CAE))
                    FechaVencimiento = CInt(.GetQueryData(TABLA.FECHA_VENCIMIENTO))
                    Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))


                    Dim tmpDet As New AfipFacturaDetalle(Sqle)
                    tmpDet.QuickSearch(AfipFacturaDetalle.TABLA.FACTURA_ID, SQLEngineQuery.OperatorCriteria.Igual, Id)
                    Detalles.Clear()
                    Detalles.AddRange(tmpDet.SearchResult)


                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End With
    End Function

    Public Function LoadMe(ByVal facturaId As Integer, ByRef extraData As AfipFacturaEX) As Boolean
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .WHEREstring = $"{TABLA.ID} = ?"
            .AddWHEREparam(facturaId)
            If .Query Then
                If .RecordCount = 1 Then
                    .QueryRead()
                    Id = CLng(.GetQueryData(TABLA.ID))
                    CuitEmisor = CLng(.GetQueryData(TABLA.CUIT_EMISOR))
                    ComprobanteTipo = CInt(.GetQueryData(TABLA.COMPROBANTE_TIPO))
                    PuntoVenta = CInt(.GetQueryData(TABLA.PUNTO_VENTA))
                    Numero = CInt(.GetQueryData(TABLA.NUMERO))
                    Concept = CInt(.GetQueryData(TABLA.CONCEPT))
                    DocumentoTipo = CInt(.GetQueryData(TABLA.DOCUMENTO_TIPO))
                    DocumentoNumero = CLng(.GetQueryData(TABLA.DOCUMENTO_NUMERO))
                    FechaEmision = CLng(.GetQueryData(TABLA.FECHA_EMISION))
                    ImporteTotal = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL))
                    ImporteTotalConc = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL_CONC))
                    ImporteNeto = CDec(.GetQueryData(TABLA.IMPORTE_NETO))
                    ImporteOpExcento = CDec(.GetQueryData(TABLA.IMPORTE_OP_EXCENTO))
                    ImporteIva = CDec(.GetQueryData(TABLA.IMPORTE_IVA))
                    ImporteTributo = CDec(.GetQueryData(TABLA.IMPORTE_TRIBUTO))
                    FechaServicioDesde = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_DESDE))
                    FechaServicioHasta = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_HASTA))
                    FechaVencimientoPago = CLng(.GetQueryData(TABLA.FECHA_VENCIMIENTO_PAGO))
                    Moneda = CStr(.GetQueryData(TABLA.MONEDA))
                    MonedaCotizacion = CDec(.GetQueryData(TABLA.MONEDA_COTIZACION))
                    Homologacion = CBool(.GetQueryData(TABLA.HOMOLOGACION))
                    EstadoSolicitud = CInt(.GetQueryData(TABLA.ESTADO))
                    MensajeAfip = CStr(.GetQueryData(TABLA.MENSAJE_AFIP))
                    Cae = CStr(.GetQueryData(TABLA.CAE))
                    FechaVencimiento = CInt(.GetQueryData(TABLA.FECHA_VENCIMIENTO))
                    Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))


                    Dim tmpDet As New AfipFacturaDetalle(Sqle)
                    tmpDet.QuickSearch(AfipFacturaDetalle.TABLA.FACTURA_ID, SQLEngineQuery.OperatorCriteria.Igual, Id)
                    Detalles.Clear()
                    Detalles.AddRange(tmpDet.SearchResult)

                    extraData.LoadMe(Id, True)


                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End With
    End Function


    Public Function LoadAll(ByRef dt As DataTable) As Boolean
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .SimpleSearch(TABLA.DELETED, SQLEngineQuery.OperatorCriteria.Igual, False)
            Return .Query(True, dt)
        End With
    End Function



    Public Function Save(ByVal editMode As Guardar) As Boolean
        Select Case editMode
            Case 0
                With Sqle.Insert
                    .Reset()
                    .TableName = TABLA.TABLA_NOMBRE
                    .AddColumnValue(TABLA.CUIT_EMISOR, CuitEmisor)
                    .AddColumnValue(TABLA.COMPROBANTE_TIPO, ComprobanteTipo)
                    .AddColumnValue(TABLA.PUNTO_VENTA, PuntoVenta)
                    .AddColumnValue(TABLA.NUMERO, Numero)
                    .AddColumnValue(TABLA.CONCEPT, Concept)
                    .AddColumnValue(TABLA.DOCUMENTO_TIPO, DocumentoTipo)
                    .AddColumnValue(TABLA.DOCUMENTO_NUMERO, DocumentoNumero)
                    .AddColumnValue(TABLA.FECHA_EMISION, FechaEmision)
                    .AddColumnValue(TABLA.IMPORTE_TOTAL, ImporteTotal)
                    .AddColumnValue(TABLA.IMPORTE_TOTAL_CONC, ImporteTotalConc)
                    .AddColumnValue(TABLA.IMPORTE_NETO, ImporteNeto)
                    .AddColumnValue(TABLA.IMPORTE_OP_EXCENTO, ImporteOpExcento)
                    .AddColumnValue(TABLA.IMPORTE_IVA, ImporteIva)
                    .AddColumnValue(TABLA.IMPORTE_TRIBUTO, ImporteTributo)
                    .AddColumnValue(TABLA.FECHA_SERVICIO_DESDE, FechaServicioDesde)
                    .AddColumnValue(TABLA.FECHA_SERVICIO_HASTA, FechaServicioHasta)
                    .AddColumnValue(TABLA.FECHA_VENCIMIENTO_PAGO, FechaVencimientoPago)
                    .AddColumnValue(TABLA.MONEDA, Moneda)
                    .AddColumnValue(TABLA.MONEDA_COTIZACION, MonedaCotizacion)
                    .AddColumnValue(TABLA.HOMOLOGACION, Homologacion)
                    .AddColumnValue(TABLA.ESTADO, EstadoSolicitud)
                    .AddColumnValue(TABLA.MENSAJE_AFIP, MensajeAfip)
                    .AddColumnValue(TABLA.CAE, Cae)
                    .AddColumnValue(TABLA.FECHA_VENCIMIENTO, FechaVencimiento)
                    .AddColumnValue(TABLA.DELETED, Deleted)
                    .AddColumnValue(TABLA.MODIFICADO, Now)


                    Dim lastID As Integer = 0
                    If .Insert(lastID) Then
                        Me.Id = lastID
                        For Each detalle As AfipFacturaDetalle In Detalles
                            HelperMod.InjectSqle(Sqle, detalle.Sqle)
                            detalle.FacturaId = lastID
                            detalle.Save(AfipFacturaDetalle.Guardar.NUEVO)
                        Next
                        Return True
                    Else
                        Return False
                    End If
                End With
            Case 1
                With Sqle.Update
                    .Reset()
                    .TableName = TABLA.TABLA_NOMBRE
                    .AddColumnValue(TABLA.CUIT_EMISOR, CuitEmisor)
                    .AddColumnValue(TABLA.COMPROBANTE_TIPO, ComprobanteTipo)
                    .AddColumnValue(TABLA.PUNTO_VENTA, PuntoVenta)
                    .AddColumnValue(TABLA.NUMERO, Numero)
                    .AddColumnValue(TABLA.CONCEPT, Concept)
                    .AddColumnValue(TABLA.DOCUMENTO_TIPO, DocumentoTipo)
                    .AddColumnValue(TABLA.DOCUMENTO_NUMERO, DocumentoNumero)
                    .AddColumnValue(TABLA.FECHA_EMISION, FechaEmision)
                    .AddColumnValue(TABLA.IMPORTE_TOTAL, ImporteTotal)
                    .AddColumnValue(TABLA.IMPORTE_TOTAL_CONC, ImporteTotalConc)
                    .AddColumnValue(TABLA.IMPORTE_NETO, ImporteNeto)
                    .AddColumnValue(TABLA.IMPORTE_OP_EXCENTO, ImporteOpExcento)
                    .AddColumnValue(TABLA.IMPORTE_IVA, ImporteIva)
                    .AddColumnValue(TABLA.IMPORTE_TRIBUTO, ImporteTributo)
                    .AddColumnValue(TABLA.FECHA_SERVICIO_DESDE, FechaServicioDesde)
                    .AddColumnValue(TABLA.FECHA_SERVICIO_HASTA, FechaServicioHasta)
                    .AddColumnValue(TABLA.FECHA_VENCIMIENTO_PAGO, FechaVencimientoPago)
                    .AddColumnValue(TABLA.MONEDA, Moneda)
                    .AddColumnValue(TABLA.MONEDA_COTIZACION, MonedaCotizacion)
                    .AddColumnValue(TABLA.HOMOLOGACION, Homologacion)
                    .AddColumnValue(TABLA.ESTADO, EstadoSolicitud)
                    .AddColumnValue(TABLA.MENSAJE_AFIP, MensajeAfip)
                    .AddColumnValue(TABLA.CAE, Cae)
                    .AddColumnValue(TABLA.FECHA_VENCIMIENTO, FechaVencimiento)
                    .AddColumnValue(TABLA.DELETED, Deleted)
                    .AddColumnValue(TABLA.MODIFICADO, Now)

                    .SimpleSearch(TABLA.ID, SQLEngineQuery.OperatorCriteria.Igual, Id)
                    Return .Update
                End With
            Case Else
                Return False
        End Select
    End Function

    Public Function Delete(Optional ByVal hard As Boolean = False) As Boolean
        If hard Then
            With Sqle.Delete
                .Reset()
                .TableName = TABLA.TABLA_NOMBRE
                .SimpleSearch(TABLA.ID, SQLEngineUpdate.OperatorCriteria.Igual, Id)
                Return .Delete
            End With
        Else
            With Sqle.Update
                .Reset()
                .TableName = TABLA.TABLA_NOMBRE
                .AddColumnValue(TABLA.DELETED, True)
                .AddColumnValue(TABLA.MODIFICADO, Now)
                .SimpleSearch(TABLA.ID, SQLEngineUpdate.OperatorCriteria.Igual, Id)
                Return .Update
            End With
        End If
    End Function

    Public Function FastSearch(ByVal tipoComprobante As Tipo, ByVal columna As String, ByVal value As Object) As Integer
        Return 0
    End Function

    Public Function Search(ByVal tipoComprobante As Tipo, ByVal columna As String, ByVal value As Object) As Integer
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            If columna = "" Then
                .AddFirstJoin(TABLA.TABLA_NOMBRE, AfipFacturaEX.TABLA.TABLA_NOMBRE, TABLA.ID, AfipFacturaEX.TABLA.FACTURA_ID)
                columna = AfipFacturaEX.TABLA.RAZON_SOCIAL_RECEPTOR
                .WHEREstring = $"{columna} LIKE ? AND {TABLA.HOMOLOGACION} = ?"
                value = $"%{value}%"
            Else
                .WHEREstring = $"{columna} = ? AND {TABLA.HOMOLOGACION} = ?"
            End If


            If columna = TABLA.DOCUMENTO_NUMERO And IsNumeric(value) Then
                .AddWHEREparam(CLng(value))
            Else
                .AddWHEREparam(value)
            End If

            .AddWHEREparam(Me.Homologacion)

            If tipoComprobante <> 0 Then
                .WHEREstring &= $" AND {TABLA.COMPROBANTE_TIPO} = ?"
                .AddWHEREparam(tipoComprobante)
            End If

            If .Query() Then
                Me.SearchResult.Clear()
                While .QueryRead
                    Dim tmp As New AfipFactura
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.CuitEmisor = CLng(.GetQueryData(TABLA.CUIT_EMISOR))
                    tmp.ComprobanteTipo = CInt(.GetQueryData(TABLA.COMPROBANTE_TIPO))
                    tmp.PuntoVenta = CInt(.GetQueryData(TABLA.PUNTO_VENTA))
                    tmp.Numero = CInt(.GetQueryData(TABLA.NUMERO))
                    tmp.Concept = CInt(.GetQueryData(TABLA.CONCEPT))
                    tmp.DocumentoTipo = CInt(.GetQueryData(TABLA.DOCUMENTO_TIPO))
                    tmp.DocumentoNumero = CLng(.GetQueryData(TABLA.DOCUMENTO_NUMERO))
                    tmp.FechaEmision = CLng(.GetQueryData(TABLA.FECHA_EMISION))
                    tmp.ImporteTotal = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL))
                    tmp.ImporteTotalConc = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL_CONC))
                    tmp.ImporteNeto = CDec(.GetQueryData(TABLA.IMPORTE_NETO))
                    tmp.ImporteOpExcento = CDec(.GetQueryData(TABLA.IMPORTE_OP_EXCENTO))
                    tmp.ImporteIva = CDec(.GetQueryData(TABLA.IMPORTE_IVA))
                    tmp.ImporteTributo = CDec(.GetQueryData(TABLA.IMPORTE_TRIBUTO))
                    tmp.FechaServicioDesde = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_DESDE))
                    tmp.FechaServicioHasta = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_HASTA))
                    tmp.FechaVencimientoPago = CLng(.GetQueryData(TABLA.FECHA_VENCIMIENTO_PAGO))
                    tmp.Moneda = CStr(.GetQueryData(TABLA.MONEDA))
                    tmp.MonedaCotizacion = CDec(.GetQueryData(TABLA.MONEDA_COTIZACION))
                    tmp.Homologacion = CBool(.GetQueryData(TABLA.HOMOLOGACION))
                    tmp.EstadoSolicitud = CInt(.GetQueryData(TABLA.ESTADO))
                    tmp.MensajeAfip = CStr(.GetQueryData(TABLA.MENSAJE_AFIP))
                    tmp.Cae = CStr(.GetQueryData(TABLA.CAE))
                    tmp.FechaVencimiento = CInt(.GetQueryData(TABLA.FECHA_VENCIMIENTO))
                    tmp.Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    tmp.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))

                    Dim tmpDet As New AfipFacturaDetalle(Sqle)
                    tmpDet.QuickSearch(AfipFacturaDetalle.TABLA.FACTURA_ID, SQLEngineQuery.OperatorCriteria.Igual, tmp.Id)
                    Detalles.Clear()
                    tmp.Detalles.AddRange(tmpDet.SearchResult)

                    SearchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function

    Public Function SearchByDate(ByVal tipoComp As Integer, ByVal desde As Integer, ByVal hasta As Integer, Optional caja As String = "", Optional dt As DataTable = Nothing, Optional soloPagos As Boolean = False) As Integer
        With Sqle.Query
            .Reset()
            .AddSelectColumn(TABLA.ID)
            .AddFirstJoin(TABLA.TABLA_NOMBRE, AfipFacturaEX.TABLA.TABLA_NOMBRE, TABLA.ID, AfipFacturaEX.TABLA.FACTURA_ID)
            .WHEREstring = $"{TABLA.COMPROBANTE_TIPO} = ? AND {TABLA.HOMOLOGACION} = ?"
            .AddWHEREparam(tipoComp)
            .AddWHEREparam(Me.Homologacion)

            If caja.Length > 0 Then
                .WHEREstring &= $" AND {AfipFacturaEX.TABLA.PUESTO_EMISION} LIKE ?"
                .AddWHEREparam(caja)
            End If

            If soloPagos Then
                .WHEREstring &= $" AND {AfipFacturaEX.TABLA.PAGADO} = ? AND ({AfipFacturaEX.TABLA.FECHA_PAGO} >= ? AND {AfipFacturaEX.TABLA.FECHA_PAGO} <= ?)"
                .AddWHEREparam(soloPagos)
            Else
                .WHEREstring &= $" AND ({TABLA.FECHA_EMISION} >= ? AND {TABLA.FECHA_EMISION} <= ?)"
            End If
            .AddWHEREparam(desde)
            .AddWHEREparam(hasta)

            If IsNothing(dt) Then
                If .Query Then
                    SearchResult.Clear()

                    While .QueryRead
                        Dim fa As New AfipFactura(Sqle)
                        fa.LoadMe(.GetQueryData(TABLA.ID))

                        SearchResult.Add(fa)
                    End While

                    Return .RecordCount
                Else
                    Return 0
                End If
            Else
                If .Query(True, dt) Then
                    Return .RecordCount
                Else
                    Return 0
                End If

            End If
        End With
    End Function

    Public Function GetNCCPorDiaPorCodigo(ByRef dt As DataTable, ByVal desdeFecha As Date, ByVal hastaFecha As Date, ByVal codigoDetalleBuscar As Integer) As Integer
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.FECHA_EMISION)
            .AddSelectColumn($"SUM({AfipFactura.TABLA.IMPORTE_TOTAL}) AS importe")
            .AddSelectColumn($"COUNT({AfipFactura.TABLA.IMPORTE_TOTAL}) AS cantidad")
            .AddFirstJoin(AfipFactura.TABLA.TABLA_NOMBRE, AfipFacturaEX.TABLA.TABLA_NOMBRE, AfipFactura.TABLA.ID, AfipFacturaEX.TABLA.FACTURA_ID)
            .AddNestedJoin(AfipFacturaDetalle.TABLA.TABLA_NOMBRE, AfipFacturaDetalle.TABLA.FACTURA_ID, AfipFactura.TABLA.ID)
            .WHEREstring = $"{AfipFactura.TABLA.COMPROBANTE_TIPO} = ? AND {AfipFactura.TABLA.HOMOLOGACION} = ? AND 
                             ({AfipFactura.TABLA.FECHA_EMISION} >= ? AND {AfipFactura.TABLA.FECHA_EMISION} <= ?) AND
                             {AfipFacturaDetalle.TABLA.CODIGO} = ?
                             GROUP BY {AfipFactura.TABLA.FECHA_EMISION}"
            .AddOrderColumn(AfipFactura.TABLA.FECHA_EMISION, SQLEngineQuery.sortOrder.ascending)
            .AddWHEREparam(13)
            .AddWHEREparam(Me.Homologacion)
            .AddWHEREparam(Utils.DateTo8601(desdeFecha.Date))
            .AddWHEREparam(Utils.DateTo8601(hastaFecha.Date))
            .AddWHEREparam(codigoDetalleBuscar.ToString)
            If .Query(True, dt) Then
                Return .RecordCount
            Else
                Return -1
            End If
        End With
    End Function



    Public Function QuickSearch(ByVal columna As String, ByVal operador As SQLEngineQuery.OperatorCriteria, ByVal value As Object) As Integer
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .SimpleSearch(columna.ToString, operador, value)
            If .Query() Then
                Me.SearchResult.Clear()
                While .QueryRead
                    Dim tmp As New AfipFactura
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.CuitEmisor = CLng(.GetQueryData(TABLA.CUIT_EMISOR))
                    tmp.ComprobanteTipo = CInt(.GetQueryData(TABLA.COMPROBANTE_TIPO))
                    tmp.PuntoVenta = CInt(.GetQueryData(TABLA.PUNTO_VENTA))
                    tmp.Numero = CInt(.GetQueryData(TABLA.NUMERO))
                    tmp.Concept = CInt(.GetQueryData(TABLA.CONCEPT))
                    tmp.DocumentoTipo = CInt(.GetQueryData(TABLA.DOCUMENTO_TIPO))
                    tmp.DocumentoNumero = CLng(.GetQueryData(TABLA.DOCUMENTO_NUMERO))
                    tmp.FechaEmision = CLng(.GetQueryData(TABLA.FECHA_EMISION))
                    tmp.ImporteTotal = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL))
                    tmp.ImporteTotalConc = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL_CONC))
                    tmp.ImporteNeto = CDec(.GetQueryData(TABLA.IMPORTE_NETO))
                    tmp.ImporteOpExcento = CDec(.GetQueryData(TABLA.IMPORTE_OP_EXCENTO))
                    tmp.ImporteIva = CDec(.GetQueryData(TABLA.IMPORTE_IVA))
                    tmp.ImporteTributo = CDec(.GetQueryData(TABLA.IMPORTE_TRIBUTO))
                    tmp.FechaServicioDesde = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_DESDE))
                    tmp.FechaServicioHasta = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_HASTA))
                    tmp.FechaVencimientoPago = CLng(.GetQueryData(TABLA.FECHA_VENCIMIENTO_PAGO))
                    tmp.Moneda = CStr(.GetQueryData(TABLA.MONEDA))
                    tmp.MonedaCotizacion = CDec(.GetQueryData(TABLA.MONEDA_COTIZACION))
                    tmp.Homologacion = CBool(.GetQueryData(TABLA.HOMOLOGACION))
                    tmp.EstadoSolicitud = CInt(.GetQueryData(TABLA.ESTADO))
                    tmp.MensajeAfip = CStr(.GetQueryData(TABLA.MENSAJE_AFIP))
                    tmp.Cae = CStr(.GetQueryData(TABLA.CAE))
                    tmp.FechaVencimiento = CInt(.GetQueryData(TABLA.FECHA_VENCIMIENTO))
                    tmp.Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    tmp.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))

                    Dim tmpDet As New AfipFacturaDetalle(Sqle)
                    tmpDet.QuickSearch(AfipFacturaDetalle.TABLA.FACTURA_ID, SQLEngineQuery.OperatorCriteria.Igual, tmp.Id)
                    Detalles.Clear()
                    tmp.Detalles.AddRange(tmpDet.SearchResult)

                    SearchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function

    Public Function LoadAllPorTipo(ByVal tipoFE As Tipo, ByVal homo As Boolean,
                                   Optional ByVal dtDesde As Date = Nothing,
                                   Optional ByVal dtHasta As Date = Nothing) As Integer
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn($"{TABLA.ALL}")
            If tipoFE = -1 Then
                .WHEREstring = $"{TABLA.COMPROBANTE_TIPO} >= ? AND {TABLA.HOMOLOGACION} = ?"
                .AddWHEREparam(0)
                .AddWHEREparam(homo)
            Else
                .WHEREstring = $"{TABLA.COMPROBANTE_TIPO} = ? AND {TABLA.HOMOLOGACION} = ?"
                .AddWHEREparam(tipoFE)
                .AddWHEREparam(homo)
            End If

            If dtDesde.Year > 2000 And dtHasta.Year > 2000 Then
                If dtDesde.Date = dtHasta.Date Then
                    .WHEREstring &= $" AND ({TABLA.FECHA_EMISION} = ?)"
                    .AddWHEREparam(ToISO8601(dtDesde))
                Else
                    .WHEREstring &= $" AND ({TABLA.FECHA_EMISION} BETWEEN ? AND ?)"
                    .AddWHEREparam(ToISO8601(dtDesde))
                    .AddWHEREparam(ToISO8601(dtHasta))
                End If
            End If
            .AddOrderColumn(TABLA.MODIFICADO, SQLEngineQuery.sortOrder.descending)

            If .Query() Then
                Me.SearchResult.Clear()
                While .QueryRead
                    Dim tmp As New AfipFactura
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.CuitEmisor = CLng(.GetQueryData(TABLA.CUIT_EMISOR))
                    tmp.ComprobanteTipo = CInt(.GetQueryData(TABLA.COMPROBANTE_TIPO))
                    tmp.PuntoVenta = CInt(.GetQueryData(TABLA.PUNTO_VENTA))
                    tmp.Numero = CInt(.GetQueryData(TABLA.NUMERO))
                    tmp.Concept = CInt(.GetQueryData(TABLA.CONCEPT))
                    tmp.DocumentoTipo = CInt(.GetQueryData(TABLA.DOCUMENTO_TIPO))
                    tmp.DocumentoNumero = CLng(.GetQueryData(TABLA.DOCUMENTO_NUMERO))
                    tmp.FechaEmision = CLng(.GetQueryData(TABLA.FECHA_EMISION))
                    tmp.ImporteTotal = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL))
                    tmp.ImporteTotalConc = CDec(.GetQueryData(TABLA.IMPORTE_TOTAL_CONC))
                    tmp.ImporteNeto = CDec(.GetQueryData(TABLA.IMPORTE_NETO))
                    tmp.ImporteOpExcento = CDec(.GetQueryData(TABLA.IMPORTE_OP_EXCENTO))
                    tmp.ImporteIva = CDec(.GetQueryData(TABLA.IMPORTE_IVA))
                    tmp.ImporteTributo = CDec(.GetQueryData(TABLA.IMPORTE_TRIBUTO))
                    tmp.FechaServicioDesde = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_DESDE))
                    tmp.FechaServicioHasta = CLng(.GetQueryData(TABLA.FECHA_SERVICIO_HASTA))
                    tmp.FechaVencimientoPago = CLng(.GetQueryData(TABLA.FECHA_VENCIMIENTO_PAGO))
                    tmp.Moneda = CStr(.GetQueryData(TABLA.MONEDA))
                    tmp.MonedaCotizacion = CDec(.GetQueryData(TABLA.MONEDA_COTIZACION))
                    tmp.Homologacion = CBool(.GetQueryData(TABLA.HOMOLOGACION))
                    tmp.EstadoSolicitud = CInt(.GetQueryData(TABLA.ESTADO))
                    tmp.MensajeAfip = CStr(.GetQueryData(TABLA.MENSAJE_AFIP))
                    tmp.Cae = CStr(.GetQueryData(TABLA.CAE))
                    tmp.FechaVencimiento = CInt(.GetQueryData(TABLA.FECHA_VENCIMIENTO))
                    tmp.Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    tmp.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))

                    SearchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function

    Public Function Autorizar() As Boolean
        Dim afip As New Afip(Homologacion)
        Dim cert As New AfipCert(Sqle)
        Dim auth As New AfipAuth(Sqle)
        Dim login As New AfipLogin(Afip.SERVICIO_FE, afip.AUTH_URL)

        cert.Homologacion = Homologacion
        auth.Homologacion = Homologacion

        If Me.ComprobanteTipo <> Tipo.RECIBO Then
            If cert.LoadActive Then
                If Not auth.LoadActive(login.Serv, Now.Ticks) Then
                    Dim loginError As String = ""
                    If Not login.Login(cert.Certificado, loginError) Then
                        Me.MensajeAfip = loginError
                        Me.EstadoSolicitud = Estado.SIN_SOLICITAR
                        Return False
                    Else
                        auth.Sign = login.Sign
                        auth.Token = login.Token
                        auth.Req = login.XDocRequest.ToString
                        auth.Res = login.XDocResponse.ToString
                        auth.GenerationTime = login.GenerationTime.Ticks
                        auth.ExpirationTime = login.ExpirationTime.Ticks
                        auth.Servicio = login.Serv

                        If Not auth.Save(AfipAuth.Guardar.NUEVO) Then
                            Me.MensajeAfip = "Error al almacenar sesión"
                            Me.EstadoSolicitud = Estado.SIN_SOLICITAR
                            Return False
                        End If
                    End If
                End If
            End If
        End If

        Dim req As New WSFEHOMO.FEAuthRequest
        req.Cuit = Me.CuitEmisor
        req.Sign = auth.Sign
        req.Token = auth.Token

        Dim feServicio As New WSFEHOMO.Service
        feServicio.Url = afip.FE_SOLICITAR_URL

        Dim feReq As New WSFEHOMO.FECAERequest
        Dim fe As New WSFEHOMO.FECAECabRequest
        fe.CantReg = 1
        fe.CbteTipo = Me.ComprobanteTipo
        fe.PtoVta = Me.PuntoVenta
        feReq.FeCabReq = fe

        Dim fed As New WSFEHOMO.FECAEDetRequest

        With fed
            .Concepto = Me.Concept
            .DocTipo = Me.DocumentoTipo
            .DocNro = Me.DocumentoNumero
            .CbteDesde = Me.Numero
            .CbteHasta = Me.Numero
            .CbteFch = Me.FechaEmision.ToString
            .ImpTotal = Me.ImporteTotal
            .ImpTotConc = Me.ImporteTotalConc
            .ImpNeto = Me.ImporteNeto
            .ImpOpEx = Me.ImporteOpExcento
            .ImpTrib = Me.ImporteTributo
            If Me.Concept = Concepto.SERVICIOS Or Me.Concept = Concepto.PRODUCTOS_Y_SERVICIOS Then
                .FchServDesde = Me.FechaServicioDesde.ToString
                .FchServHasta = Me.FechaServicioHasta.ToString
                .FchVtoPago = Me.FechaVencimientoPago.ToString
            End If
            .MonId = Me.Moneda
            .MonCotiz = Me.MonedaCotizacion
        End With

        ' Si es RECIBO X completar los datos y salir
        If Me.ComprobanteTipo >= Tipo.RECIBO Then
            Me.Cae = "00000000000000"
            Me.FechaVencimiento = Utils.DateTo8601(Now)
            Me.EstadoSolicitud = Estado.APROBADO
            Return True
        End If

        If Me.ComprobanteTipo = Tipo.NOTA_CREDITO_C Or
            Me.ComprobanteTipo = Tipo.NOTA_CREDITO_A Or
            Me.ComprobanteTipo = Tipo.NOTA_CREDITO_B Or
            Me.ComprobanteTipo = Tipo.NOTA_DEBITO_C Then


            Dim array As New List(Of WSFEHOMO.CbteAsoc)
            For Each comprobanteAsociado As AfipFactura In Me.CbtesAsoc
                Dim cbteasoc As New WSFEHOMO.CbteAsoc
                cbteasoc.Cuit = comprobanteAsociado.CuitEmisor
                'cbteasoc.CbteFch = comprobanteAsociado.FechaEmision
                cbteasoc.Nro = comprobanteAsociado.Numero
                cbteasoc.PtoVta = comprobanteAsociado.PuntoVenta
                cbteasoc.Tipo = comprobanteAsociado.ComprobanteTipo
                array.Add(cbteasoc)
            Next
            fed.CbtesAsoc = array.ToArray

        End If


        feReq.FeDetReq = {fed}

        Dim feRes As New WSFEHOMO.FECAEResponse
        Try
            feRes = feServicio.FECAESolicitar(req, feReq)
            If IsNothing(feRes.Errors) Then
                Select Case feRes.FeCabResp.Resultado
                    Case "A"
                        Me.Cae = feRes.FeDetResp(0).CAE
                        Me.FechaVencimiento = CInt(feRes.FeDetResp(0).CAEFchVto)
                        Me.EstadoSolicitud = Estado.APROBADO
                        Return True
                    Case Else
                        If Not IsNothing(feRes.Errors) Then
                            For Each obs In feRes.Errors
                                Me.MensajeAfip &= obs.Msg & vbCrLf
                            Next
                        End If


                        Me.EstadoSolicitud = Estado.RECHAZADO
                        Return False
                End Select
            Else
                For Each obs In feRes.Errors
                    Me.MensajeAfip &= obs.Msg & vbCrLf
                Next
                Me.EstadoSolicitud = Estado.RECHAZADO
                Return False
            End If
        Catch ex As Exception
            Me.MensajeAfip = ex.Message
            Me.EstadoSolicitud = Estado.RECHAZADO
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Devuelve el ultimo numero de comprobante
    ''' </summary>
    ''' <param name="auth">Autorizacion al servicio</param>
    ''' <param name="cuit">Cuit del emisor</param>
    ''' <param name="puntoVenta">Punto de venta del emisor</param>
    ''' <param name="comprobanteTipo">Tipo de comprobante: factura, recibo, etc</param>
    ''' <returns></returns>
    Public Shared Function GetUltimoNumero(ByVal auth As AfipAuth, ByVal cuit As Long, ByVal puntoVenta As Integer, ByVal comprobanteTipo As Tipo, ByVal homologacion As Boolean) As Integer
        Dim afip As New Afip(homologacion)
        Dim feServicio As New WSFEHOMO.Service
        feServicio.Url = afip.ULTIMO_NUMERO_URL
        Dim req As New WSFEHOMO.FEAuthRequest
        req.Cuit = cuit
        req.Sign = auth.Sign
        req.Token = auth.Token
        Try
            Dim feRes As New WSFEHOMO.FERecuperaLastCbteResponse
            feRes = feServicio.FECompUltimoAutorizado(req, puntoVenta, comprobanteTipo)
            Return CInt(feRes.CbteNro)


        Catch ex As Exception
            Return -1
        End Try
    End Function

    Public Function GetUltimoNumero() As Integer
        Dim afip As New Afip(Homologacion)
        Dim cert As New AfipCert(Sqle)
        Dim auth As New AfipAuth(Sqle)
        Dim login As New AfipLogin(Afip.SERVICIO_FE, afip.AUTH_URL)

        cert.Homologacion = Homologacion
        auth.Homologacion = Homologacion

        If cert.LoadActive Then
            If Not auth.LoadActive(login.Serv, Now.Ticks) Then
                Dim loginError As String = ""
                If Not login.Login(cert.Certificado, loginError) Then
                    Me.MensajeAfip = loginError
                    Me.EstadoSolicitud = Estado.SIN_SOLICITAR
                    Return False
                Else
                    auth.Sign = login.Sign
                    auth.Token = login.Token
                    auth.Req = login.XDocRequest.ToString
                    auth.Res = login.XDocResponse.ToString
                    auth.GenerationTime = login.GenerationTime.Ticks
                    auth.ExpirationTime = login.ExpirationTime.Ticks
                    auth.Servicio = login.Serv

                    If Not auth.Save(AfipAuth.Guardar.NUEVO) Then
                        Me.MensajeAfip = "Error al almacenar sesión"
                        Me.EstadoSolicitud = Estado.SIN_SOLICITAR
                        Return False
                    End If
                End If
            End If
        End If
        Dim req As New WSFEHOMO.FEAuthRequest
        req.Cuit = Me.CuitEmisor
        req.Sign = auth.Sign
        req.Token = auth.Token

        Dim feServicio As New WSFEHOMO.Service
        feServicio.Url = afip.ULTIMO_NUMERO_URL

        Try
            Dim feRes As New WSFEHOMO.FERecuperaLastCbteResponse
            feRes = feServicio.FECompUltimoAutorizado(req, Me.PuntoVenta, Me.ComprobanteTipo)
            Return CInt(feRes.CbteNro)


        Catch ex As Exception
            Return -1
        End Try
    End Function

    Public Function GetUltimoNumeroLocal() As Integer
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .WHEREstring = $"{TABLA.HOMOLOGACION} = ? AND {TABLA.COMPROBANTE_TIPO} = ? AND {TABLA.PUNTO_VENTA} = ?  AND {TABLA.DELETED} = ?"
            .AddWHEREparam(Homologacion)
            .AddWHEREparam(ComprobanteTipo)
            .AddWHEREparam(PuntoVenta)
            .AddWHEREparam(False)
            .AddOrderColumn(TABLA.NUMERO, SQLEngineQuery.sortOrder.descending)

            If .Query Then
                If .RecordCount > 0 Then
                    .QueryRead()
                    Return CInt(.GetQueryData(TABLA.NUMERO))
                Else
                    Return 0
                End If
            Else
                Return -1
            End If
        End With
    End Function

    Public Function LoadUltimoComprobante() As Boolean
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .WHEREstring = $"{TABLA.HOMOLOGACION} = ? AND {TABLA.COMPROBANTE_TIPO} = ? AND {TABLA.PUNTO_VENTA} = ?  AND {TABLA.DELETED} = ?"
            .AddWHEREparam(Homologacion)
            .AddWHEREparam(ComprobanteTipo)
            .AddWHEREparam(PuntoVenta)
            .AddWHEREparam(False)
            .AddOrderColumn(TABLA.NUMERO, SQLEngineQuery.sortOrder.descending)

            If .Query Then
                Dim res As New AfipFactura(Me.Sqle)
                If .RecordCount > 0 Then
                    .QueryRead()
                    Dim ultimoId As Integer = CInt(.GetQueryData(TABLA.ID))
                    If Me.LoadMe(ultimoId) Then
                        Return True
                    Else
                        Return False
                    End If
                End If
                Me.Numero = 0
                Me.FechaEmision = ToISO8601(Now.AddDays(-5))
                Return True
            Else
                Return False
            End If
        End With
    End Function

    ''' <summary>
    ''' Chequear que la factura contiene un item con detalle buscado
    ''' </summary>
    ''' <param name="codigoDetalle">El codigo de detalle a buscar</param>
    ''' <returns></returns>
    Public Function TieneDetalleCodigo(ByVal codigoDetalle As String) As Boolean
        Dim res As Boolean = False
        For Each detalle As AfipFacturaDetalle In Detalles
            If detalle.Codigo = codigoDetalle Then
                Return True
            End If
        Next

        Return res
    End Function





End Class

