Imports System.Data.SqlClient
Imports helix
Public Class MovimientoCuentaCorrienteSocio

    Public Enum TIPO As Integer
        FACTURA_C = 0
        RECIBO_X = 1
        TICKET_X = 2
        NOTA_CREDITO_C = 3
        NOTA_CREDITO_X = 4
    End Enum

    Public Property Sqle As New SQLEngine
    Public ReadOnly Property SearchResult As New List(Of MovimientoCuentaCorrienteSocio)

    ' Listado de productos contenidos en el movimiento
    Public ReadOnly Property CuotasSociales As New List(Of CuotaSocio)
    Public ReadOnly Property Productos As New List(Of ProductoSocio)


    Public Property Id As Long = 0
    Public Property ClienteId As Long = 0
    Public Property FechaIngreso As Integer = 0
    Public Property TipoMovimiento As Integer = 0

    ''' <summary>
    ''' El id del comprobante de la operación
    ''' </summary>
    ''' <returns>El id del recibo / factura si es debe, el ID de operación si es pago</returns>
    Public Property ComprobanteRelacionado As Long = 0
    Public Property ComprobanteTipo As TIPO = TIPO.RECIBO_X
    Public Property Importe As Decimal = 0
    Public Property ImporteCobrar As Decimal = 0
    Public Property Procesado As Boolean = False
    Public Property Deleted As Boolean = False
    Public Property Modificado As Date = Now

    Public Enum Guardar As Integer
        NUEVO = 0
        EDITAR = 1
    End Enum

    Public Enum TipoDeMovimiento As Integer
        SOCIOS_CUOTA_SOCIAL = 0
        SOCIOS_OTROS = 1
        SOCIOS_PUBLICIDAD = 2
        SOCIOS_BOLSIN = 3
        SOCIOS_MEDICINA = 4
    End Enum

    Private Enum Modalidad As Integer
        TODO = 0
        DEBE = 1
        HABER = 2
    End Enum

    Public Structure TABLA
        Const TABLA_NOMBRE As String = "MovimientoCuentaCorrienteSocio"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const CLIENTE_ID As String = TABLA_NOMBRE & "_clienteId"
        Const FECHA_INGRESO As String = TABLA_NOMBRE & "_fechaIngreso"
        Const TIPO_MOVIMIENTO As String = TABLA_NOMBRE & "_tipoMovimiento"
        Const COMPROBANTE_RELACIONADO As String = TABLA_NOMBRE & "_comprobanteRelacionado"
        Const COMPROBANTE_TIPO As String = TABLA_NOMBRE & "_comprobanteTipo"
        Const IMPORTE As String = TABLA_NOMBRE & "_importe"
        Const IMPORTE_COBRAR As String = TABLA_NOMBRE & "_importeCobrar"
        Const PROCESADO As String = TABLA_NOMBRE & "_procesado"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & CLIENTE_ID & ", " & FECHA_INGRESO & ", " & TIPO_MOVIMIENTO & ", " & COMPROBANTE_RELACIONADO & ", " & COMPROBANTE_TIPO & ", " & IMPORTE & ", " & IMPORTE_COBRAR & ", " & PROCESADO & ", " & DELETED & ", " & MODIFICADO
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

    Public Sub InjectSQL(ByVal iSqle As SQLEngine)
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
                    ClienteId = CLng(.GetQueryData(TABLA.CLIENTE_ID))
                    FechaIngreso = CInt(.GetQueryData(TABLA.FECHA_INGRESO))
                    TipoMovimiento = CInt(.GetQueryData(TABLA.TIPO_MOVIMIENTO))
                    ComprobanteRelacionado = CLng(.GetQueryData(TABLA.COMPROBANTE_RELACIONADO))
                    ComprobanteTipo = CInt(.GetQueryData(TABLA.COMPROBANTE_TIPO))
                    Importe = CDec(.GetQueryData(TABLA.IMPORTE))
                    ImporteCobrar = CDec(.GetQueryData(TABLA.IMPORTE_COBRAR))
                    Procesado = CBool(.GetQueryData(TABLA.PROCESADO))
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

    Public Function LoadByFactura(ByVal facturaID As Integer) As Boolean
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .WHEREstring = $"{TABLA.COMPROBANTE_RELACIONADO} = { .p(facturaID)}  AND {TABLA.DELETED} = { .p(False)}"
            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    Id = CLng(.GetQueryData(TABLA.ID))
                    ClienteId = CLng(.GetQueryData(TABLA.CLIENTE_ID))
                    FechaIngreso = CInt(.GetQueryData(TABLA.FECHA_INGRESO))
                    TipoMovimiento = CInt(.GetQueryData(TABLA.TIPO_MOVIMIENTO))
                    ComprobanteRelacionado = CLng(.GetQueryData(TABLA.COMPROBANTE_RELACIONADO))
                    ComprobanteTipo = CInt(.GetQueryData(TABLA.COMPROBANTE_TIPO))
                    Importe = CDec(.GetQueryData(TABLA.IMPORTE))
                    ImporteCobrar = CDec(.GetQueryData(TABLA.IMPORTE_COBRAR))
                    Procesado = CBool(.GetQueryData(TABLA.PROCESADO))
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
                    .AddColumnValue(TABLA.CLIENTE_ID, ClienteId)
                    .AddColumnValue(TABLA.FECHA_INGRESO, FechaIngreso)
                    .AddColumnValue(TABLA.TIPO_MOVIMIENTO, TipoMovimiento)
                    .AddColumnValue(TABLA.COMPROBANTE_RELACIONADO, ComprobanteRelacionado)
                    .AddColumnValue(TABLA.COMPROBANTE_TIPO, ComprobanteTipo)
                    .AddColumnValue(TABLA.IMPORTE, Importe)
                    .AddColumnValue(TABLA.IMPORTE_COBRAR, ImporteCobrar)
                    .AddColumnValue(TABLA.PROCESADO, Procesado)
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
                    .AddColumnValue(TABLA.CLIENTE_ID, ClienteId)
                    .AddColumnValue(TABLA.FECHA_INGRESO, FechaIngreso)
                    .AddColumnValue(TABLA.TIPO_MOVIMIENTO, TipoMovimiento)
                    .AddColumnValue(TABLA.COMPROBANTE_RELACIONADO, ComprobanteRelacionado)
                    .AddColumnValue(TABLA.COMPROBANTE_TIPO, ComprobanteTipo)
                    .AddColumnValue(TABLA.IMPORTE, Importe)
                    .AddColumnValue(TABLA.IMPORTE_COBRAR, ImporteCobrar)
                    .AddColumnValue(TABLA.PROCESADO, Procesado)
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

    Private Sub LoadQueryData(ByVal currentSqle As SQLEngine, ByRef obj As MovimientoCuentaCorrienteSocio)
        With currentSqle.Query
            obj.Id = CLng(.GetQueryData(TABLA.ID))
            obj.ClienteId = CLng(.GetQueryData(TABLA.CLIENTE_ID))
            obj.FechaIngreso = CInt(.GetQueryData(TABLA.FECHA_INGRESO))
            obj.TipoMovimiento = CInt(.GetQueryData(TABLA.TIPO_MOVIMIENTO))
            obj.ComprobanteRelacionado = CLng(.GetQueryData(TABLA.COMPROBANTE_RELACIONADO))
            obj.ComprobanteTipo = CInt(.GetQueryData(TABLA.COMPROBANTE_TIPO))
            obj.Importe = CDec(.GetQueryData(TABLA.IMPORTE))
            obj.ImporteCobrar = CDec(.GetQueryData(TABLA.IMPORTE_COBRAR))
            obj.Procesado = CBool(.GetQueryData(TABLA.PROCESADO))
            obj.Deleted = CBool(.GetQueryData(TABLA.DELETED))
            obj.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))
        End With
    End Sub

    Public Function QuickSearch(ByVal columna As String, ByVal operador As SQLEngineQuery.OperatorCriteria, ByVal value As Object, Optional order As SortOrder = SortOrder.Ascending) As Integer
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .SimpleSearch(columna.ToString, operador, value)
            .AddOrderColumn(TABLA.FECHA_INGRESO, order)
            .AddOrderColumn(TABLA.ID, order)
            If .Query() Then
                Me.SearchResult.Clear()
                While .QueryRead
                    Dim tmp As New MovimientoCuentaCorrienteSocio
                    LoadQueryData(Sqle, tmp)
                    SearchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function

    ''' <summary>
    ''' Carga los movimientos impagos por socio
    ''' </summary>
    ''' <param name="idSocio">ID del socio a buscar</param>
    ''' <returns>El numero de movimientos recolectados, entero negativo si fallo</returns>
    Public Function CargarMovimientosImpagosSocios(ByVal idSocio As Integer) As Integer
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .WHEREstring = $"{TABLA.CLIENTE_ID} = { .p(idSocio)} AND 
							 {TABLA.DELETED} =  { .p(False)} AND
							 {TABLA.IMPORTE} < { .p(0)} AND
							 {TABLA.PROCESADO} = { .p(False)}"


            .AddOrderColumn(TABLA.FECHA_INGRESO, SQLEngineQuery.sortOrder.descending)
            .AddOrderColumn(TABLA.ID, SQLEngineQuery.sortOrder.descending)
            If .Query() Then
                Me.SearchResult.Clear()
                While .QueryRead
                    Dim tmp As New MovimientoCuentaCorrienteSocio
                    LoadQueryData(Sqle, tmp)
                    SearchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return -1
    End Function
End Class

