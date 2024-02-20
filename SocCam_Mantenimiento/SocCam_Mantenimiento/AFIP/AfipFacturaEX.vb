Imports helix

Public Class AfipFacturaEX
    Public Property Sqle As New SQLEngine
    Public Property SearchResult As New List(Of AfipFacturaEX)

    Public Property Id As Long = 0
    Public Property FacturaId As Long = 0
    Public Property CondicionContado As Boolean = False
    Public Property CondicionTarjetaDebito As Boolean = False
    Public Property CondicionTarjetaCredito As Boolean = False
    Public Property CondicionCuentaCorriente As Boolean = False
    Public Property CondicionCheque As Boolean = False
    Public Property CondicionTicket As Boolean = False
    Public Property CondicionOtra As Boolean = False
    Public Property CondicionOtraDet As String = ""
    Public Property RazonSocialEmisor As String = ""
    Public Property RazonSocialReceptor As String = ""
    Public Property DomicilioEmisor As String = ""
    Public Property DomicilioReceptor As String = ""
    Public Property CondicionFiscalStringReceptor As String = ""
    Public Property FacturaRendered As String = ""
    Public Property Operador As String = ""
    Public Property PuestoEmision As String = ""
    Public Property Pagado As Boolean = False
    Public Property FechaPago As Long = 0
    Public Property Deleted As Boolean = False
    Public Property Modificado As Date = Now


    Public Enum Guardar
        NUEVO = 0
        EDITAR = 1
    End Enum


    Public Structure TABLA
        Const TABLA_NOMBRE As String = "AfipFacturaEx"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const FACTURA_ID As String = TABLA_NOMBRE & "_facturaId"
        Const CONDICION_CONTADO As String = TABLA_NOMBRE & "_condicionContado"
        Const CONDICION_TARJETA_DEBITO As String = TABLA_NOMBRE & "_condicionTarjetaDebito"
        Const CONDICION_TARJETA_CREDITO As String = TABLA_NOMBRE & "_condicionTarjetaCredito"
        Const CONDICION_CUENTA_CORRIENTE As String = TABLA_NOMBRE & "_condicionCuentaCorriente"
        Const CONDICION_CHEQUE As String = TABLA_NOMBRE & "_condicionCheque"
        Const CONDICION_TICKET As String = TABLA_NOMBRE & "_condicionTicket"
        Const CONDICION_OTRA As String = TABLA_NOMBRE & "_condicionOtra"
        Const CONDICION_OTRA_DET As String = TABLA_NOMBRE & "_condicionOtraDet"
        Const RAZON_SOCIAL_EMISOR As String = TABLA_NOMBRE & "_razonSocialEmisor"
        Const RAZON_SOCIAL_RECEPTOR As String = TABLA_NOMBRE & "_razonSocialReceptor"
        Const DOMICILIO_EMISOR As String = TABLA_NOMBRE & "_domicilioEmisor"
        Const DOMICILIO_RECEPTOR As String = TABLA_NOMBRE & "_domicilioReceptor"
        Const CONDICION_FISCAL_STRING_RECEPTOR As String = TABLA_NOMBRE & "_condicionFiscalStringReceptor"
        Const FACTURA_RENDERED As String = TABLA_NOMBRE & "_facturaRendered"
        Const OPERADOR As String = TABLA_NOMBRE & "_operador"
        Const PUESTO_EMISION As String = TABLA_NOMBRE & "_puestoEmision"
        Const PAGADO As String = TABLA_NOMBRE & "_pagado"
        Const FECHA_PAGO As String = TABLA_NOMBRE & "_fechaPago"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & FACTURA_ID & ", " & CONDICION_CONTADO & ", " & CONDICION_TARJETA_DEBITO & ", " & CONDICION_TARJETA_CREDITO & ", " & CONDICION_CUENTA_CORRIENTE & ", " & CONDICION_CHEQUE & ", " & CONDICION_TICKET & ", " & CONDICION_OTRA & ", " & CONDICION_OTRA_DET & ", " & RAZON_SOCIAL_EMISOR & ", " & RAZON_SOCIAL_RECEPTOR & ", " & DOMICILIO_EMISOR & ", " & DOMICILIO_RECEPTOR & ", " & CONDICION_FISCAL_STRING_RECEPTOR & ", " & FACTURA_RENDERED & ", " & OPERADOR & ", " & PUESTO_EMISION & ", " & PAGADO & ", " & FECHA_PAGO & ", " & DELETED & ", " & MODIFICADO
    End Structure



    Public Sub New()
    End Sub


    Public Sub New(ByVal sqle As SQLEngine)
        Me.Sqle.RequireCredentials = sqle.RequireCredentials
        Me.Sqle.Username = sqle.Username
        Me.Sqle.Password = sqle.Password
        Me.Sqle.dbType = sqle.dbType
        Me.Sqle.Path = sqle.Path
        Me.Sqle.DatabaseName = sqle.DatabaseName
        If sqle.IsStarted Then
            Me.Sqle.ColdBoot()
        Else
            Me.Sqle.Start()
        End If
    End Sub


    Public Function LoadMe(ByVal myID As Integer, Optional esFacturaID As Boolean = False) As Boolean
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            If esFacturaID Then
                .SimpleSearch(TABLA.FACTURA_ID, SQLEngineQuery.OperatorCriteria.Igual, myID)
            Else
                .SimpleSearch(TABLA.ID, SQLEngineQuery.OperatorCriteria.Igual, myID)
            End If
            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    Id = CLng(.GetQueryData(TABLA.ID))
                    FacturaId = CLng(.GetQueryData(TABLA.FACTURA_ID))
                    CondicionContado = CBool(.GetQueryData(TABLA.CONDICION_CONTADO))
                    CondicionTarjetaDebito = CBool(.GetQueryData(TABLA.CONDICION_TARJETA_DEBITO))
                    CondicionTarjetaCredito = CBool(.GetQueryData(TABLA.CONDICION_TARJETA_CREDITO))
                    CondicionCuentaCorriente = CBool(.GetQueryData(TABLA.CONDICION_CUENTA_CORRIENTE))
                    CondicionCheque = CBool(.GetQueryData(TABLA.CONDICION_CHEQUE))
                    CondicionTicket = CBool(.GetQueryData(TABLA.CONDICION_TICKET))
                    CondicionOtra = CBool(.GetQueryData(TABLA.CONDICION_OTRA))
                    CondicionOtraDet = CStr(.GetQueryData(TABLA.CONDICION_OTRA_DET))
                    RazonSocialEmisor = CStr(.GetQueryData(TABLA.RAZON_SOCIAL_EMISOR))
                    RazonSocialReceptor = CStr(.GetQueryData(TABLA.RAZON_SOCIAL_RECEPTOR))
                    DomicilioEmisor = CStr(.GetQueryData(TABLA.DOMICILIO_EMISOR))
                    DomicilioReceptor = CStr(.GetQueryData(TABLA.DOMICILIO_RECEPTOR))
                    CondicionFiscalStringReceptor = CStr(.GetQueryData(TABLA.CONDICION_FISCAL_STRING_RECEPTOR))
                    FacturaRendered = CStr(.GetQueryData(TABLA.FACTURA_RENDERED))
                    Operador = CStr(.GetQueryData(TABLA.OPERADOR))
                    PuestoEmision = CStr(.GetQueryData(TABLA.PUESTO_EMISION))
                    If Not IsDBNull(.GetQueryData(TABLA.PAGADO)) Then
                        Pagado = False
                        FechaPago = 0
                    Else
                        Pagado = CBool(.GetQueryData(TABLA.PAGADO))
                        FechaPago = CLng(.GetQueryData(TABLA.FECHA_PAGO))
                    End If
                    Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))
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
                    .AddColumnValue(TABLA.FACTURA_ID, FacturaId)
                    .AddColumnValue(TABLA.CONDICION_CONTADO, CondicionContado)
                    .AddColumnValue(TABLA.CONDICION_TARJETA_DEBITO, CondicionTarjetaDebito)
                    .AddColumnValue(TABLA.CONDICION_TARJETA_CREDITO, CondicionTarjetaCredito)
                    .AddColumnValue(TABLA.CONDICION_CUENTA_CORRIENTE, CondicionCuentaCorriente)
                    .AddColumnValue(TABLA.CONDICION_CHEQUE, CondicionCheque)
                    .AddColumnValue(TABLA.CONDICION_TICKET, CondicionTicket)
                    .AddColumnValue(TABLA.CONDICION_OTRA, CondicionOtra)
                    .AddColumnValue(TABLA.CONDICION_OTRA_DET, CondicionOtraDet)
                    .AddColumnValue(TABLA.RAZON_SOCIAL_EMISOR, RazonSocialEmisor)
                    .AddColumnValue(TABLA.RAZON_SOCIAL_RECEPTOR, RazonSocialReceptor)
                    .AddColumnValue(TABLA.DOMICILIO_EMISOR, DomicilioEmisor)
                    .AddColumnValue(TABLA.DOMICILIO_RECEPTOR, DomicilioReceptor)
                    .AddColumnValue(TABLA.CONDICION_FISCAL_STRING_RECEPTOR, CondicionFiscalStringReceptor)
                    .AddColumnValue(TABLA.FACTURA_RENDERED, FacturaRendered)
                    .AddColumnValue(TABLA.OPERADOR, Operador)
                    .AddColumnValue(TABLA.PUESTO_EMISION, PuestoEmision)
                    .AddColumnValue(TABLA.PAGADO, Pagado)
                    .AddColumnValue(TABLA.FECHA_PAGO, FechaPago)
                    .AddColumnValue(TABLA.DELETED, Deleted)
                    .AddColumnValue(TABLA.MODIFICADO, Now)
                    Dim lastID As Integer = 0
                    If .Insert(lastID) Then
                        Id = lastID
                        Return True
                    Else
                        Return False
                    End If
                End With
            Case 1
                With Sqle.Update
                    .Reset()
                    .TableName = TABLA.TABLA_NOMBRE
                    .AddColumnValue(TABLA.FACTURA_ID, FacturaId)
                    .AddColumnValue(TABLA.CONDICION_CONTADO, CondicionContado)
                    .AddColumnValue(TABLA.CONDICION_TARJETA_DEBITO, CondicionTarjetaDebito)
                    .AddColumnValue(TABLA.CONDICION_TARJETA_CREDITO, CondicionTarjetaCredito)
                    .AddColumnValue(TABLA.CONDICION_CUENTA_CORRIENTE, CondicionCuentaCorriente)
                    .AddColumnValue(TABLA.CONDICION_CHEQUE, CondicionCheque)
                    .AddColumnValue(TABLA.CONDICION_TICKET, CondicionTicket)
                    .AddColumnValue(TABLA.CONDICION_OTRA, CondicionOtra)
                    .AddColumnValue(TABLA.CONDICION_OTRA_DET, CondicionOtraDet)
                    .AddColumnValue(TABLA.RAZON_SOCIAL_EMISOR, RazonSocialEmisor)
                    .AddColumnValue(TABLA.RAZON_SOCIAL_RECEPTOR, RazonSocialReceptor)
                    .AddColumnValue(TABLA.DOMICILIO_EMISOR, DomicilioEmisor)
                    .AddColumnValue(TABLA.DOMICILIO_RECEPTOR, DomicilioReceptor)
                    .AddColumnValue(TABLA.CONDICION_FISCAL_STRING_RECEPTOR, CondicionFiscalStringReceptor)
                    .AddColumnValue(TABLA.FACTURA_RENDERED, FacturaRendered)
                    .AddColumnValue(TABLA.OPERADOR, Operador)
                    .AddColumnValue(TABLA.PUESTO_EMISION, PuestoEmision)
                    .AddColumnValue(TABLA.PAGADO, Pagado)
                    .AddColumnValue(TABLA.FECHA_PAGO, FechaPago)
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


    Public Function QuickSearch(ByVal columna As String, ByVal operador As SQLEngineQuery.OperatorCriteria, ByVal value As Object) As Integer
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .SimpleSearch(columna.ToString, operador, value)
            If .Query() Then
                Me.SearchResult.Clear()
                While .QueryRead
                    Dim tmp As New AfipFacturaEX
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.FacturaId = CLng(.GetQueryData(TABLA.FACTURA_ID))
                    tmp.CondicionContado = CBool(.GetQueryData(TABLA.CONDICION_CONTADO))
                    tmp.CondicionTarjetaDebito = CBool(.GetQueryData(TABLA.CONDICION_TARJETA_DEBITO))
                    tmp.CondicionTarjetaCredito = CBool(.GetQueryData(TABLA.CONDICION_TARJETA_CREDITO))
                    tmp.CondicionCuentaCorriente = CBool(.GetQueryData(TABLA.CONDICION_CUENTA_CORRIENTE))
                    tmp.CondicionCheque = CBool(.GetQueryData(TABLA.CONDICION_CHEQUE))
                    tmp.CondicionTicket = CBool(.GetQueryData(TABLA.CONDICION_TICKET))
                    tmp.CondicionOtra = CBool(.GetQueryData(TABLA.CONDICION_OTRA))
                    tmp.CondicionOtraDet = CStr(.GetQueryData(TABLA.CONDICION_OTRA_DET))
                    tmp.RazonSocialEmisor = CStr(.GetQueryData(TABLA.RAZON_SOCIAL_EMISOR))
                    tmp.RazonSocialReceptor = CStr(.GetQueryData(TABLA.RAZON_SOCIAL_RECEPTOR))
                    tmp.DomicilioEmisor = CStr(.GetQueryData(TABLA.DOMICILIO_EMISOR))
                    tmp.DomicilioReceptor = CStr(.GetQueryData(TABLA.DOMICILIO_RECEPTOR))
                    tmp.CondicionFiscalStringReceptor = CStr(.GetQueryData(TABLA.CONDICION_FISCAL_STRING_RECEPTOR))
                    tmp.FacturaRendered = CStr(.GetQueryData(TABLA.FACTURA_RENDERED))
                    tmp.Operador = CStr(.GetQueryData(TABLA.OPERADOR))
                    tmp.PuestoEmision = CStr(.GetQueryData(TABLA.PUESTO_EMISION))
                    tmp.Pagado = CBool(.GetQueryData(TABLA.PAGADO))
                    tmp.FechaPago = CLng(.GetQueryData(TABLA.FECHA_PAGO))
                    tmp.Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    tmp.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))
                    SearchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function

End Class
