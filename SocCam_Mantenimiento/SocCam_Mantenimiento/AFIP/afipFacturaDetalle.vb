Imports helix

Public Class AfipFacturaDetalle
    Public Property Sqle As New SQLEngine
    Public Property SearchResult As New List(Of AfipFacturaDetalle)

    Public Property Id As Long = 0
    Public Property FacturaId As Long = 0
    Public Property Codigo As String = ""
    Public Property ProductoServicio As String = ""
    Public Property Cantidad As Integer = 0
    Public Property UnidadMedida As Unidad = Unidad.OTRAS_UNIDADES
    Public Property PrecioUnitario As Decimal = 0
    Public Property BonificacionPercent As Decimal = 0
    Public Property CuotaId As Long = 0
    Public Property Deleted As Boolean = False
    Public Property Modificado As Date = Now

    Public Enum Guardar
        NUEVO = 0
        EDITAR = 1
    End Enum

    Public Enum Unidad As Integer
        KILOGRAMOS = 1
        METROS = 2
        METROS_CUADRADOS = 3
        METROS_CUBICOS = 4
        LITROS = 5
        KWH = 6
        UNIDADES = 7
        PARES = 8
        DOCENAS = 9
        QUILATES = 10
        MILLARES = 11
        GRAMOS = 14
        MILIMETROS = 15
        MM_CUBICOS = 16
        KILOMETROS = 17
        HECTOLITROS = 18
        CENTIMETROS = 20
        JGO_PQT_MAZO_NAIPES = 25
        CM_CUBICOS = 27
        TONELADAS = 29
        DAM_CUBICOS = 30
        HM_CUBICOS = 31
        KM_CUBICOS = 32
        MICROGRAMOS = 33
        NANOGRAMOS = 34
        PICOGRAMOS = 35
        MILIGRAMOS = 41
        MILILITROS = 47
        CURIE = 48
        MILICURIE = 49
        MICROCURIE = 50
        UIACTHOR = 51
        MUIACTHOR = 52
        KG_BASE = 53
        GRUESA = 54
        KG_BRUTO = 61
        UIACTANT = 62
        MUIACTANT = 63
        UIACTIG = 64
        MUIACTIG = 65
        KG_ACTIVO = 66
        GRAMO_ACTIVO = 67
        GRAMO_BASE = 68
        PACKS = 96
        OTRAS_UNIDADES = 98
    End Enum


    Public Structure TABLA
        Const TABLA_NOMBRE As String = "afipFacturaDetalle"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const FACTURA_ID As String = TABLA_NOMBRE & "_facturaId"
        Const CODIGO As String = TABLA_NOMBRE & "_codigo"
        Const PRODUCTO_SERVICIO As String = TABLA_NOMBRE & "_productoServicio"
        Const CANTIDAD As String = TABLA_NOMBRE & "_cantidad"
        Const UNIDAD_MEDIDA As String = TABLA_NOMBRE & "_unidadMedida"
        Const PRECIO_UNITARIO As String = TABLA_NOMBRE & "_precioUnitario"
        Const BONIFICACION_PERCENT As String = TABLA_NOMBRE & "_bonificacionPercent"
        Const CUOTA_ID As String = TABLA_NOMBRE & "_cuotaId"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & FACTURA_ID & ", " & CODIGO & ", " & PRODUCTO_SERVICIO & ", " & CANTIDAD & ", " & UNIDAD_MEDIDA & ", " & PRECIO_UNITARIO & ", " & BONIFICACION_PERCENT & ", " & CUOTA_ID & ", " & DELETED & ", " & MODIFICADO
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
                    FacturaId = CLng(.GetQueryData(TABLA.FACTURA_ID))
                    Codigo = CStr(.GetQueryData(TABLA.CODIGO))
                    ProductoServicio = CStr(.GetQueryData(TABLA.PRODUCTO_SERVICIO))
                    Cantidad = CInt(.GetQueryData(TABLA.CANTIDAD))
                    UnidadMedida = CInt(.GetQueryData(TABLA.UNIDAD_MEDIDA))
                    PrecioUnitario = CDec(.GetQueryData(TABLA.PRECIO_UNITARIO))
                    BonificacionPercent = CDec(.GetQueryData(TABLA.BONIFICACION_PERCENT))
                    CuotaId = CLng(.GetQueryData(TABLA.CUOTA_ID))
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
                    .AddColumnValue(TABLA.CODIGO, Codigo)
                    .AddColumnValue(TABLA.PRODUCTO_SERVICIO, ProductoServicio)
                    .AddColumnValue(TABLA.CANTIDAD, Cantidad)
                    .AddColumnValue(TABLA.UNIDAD_MEDIDA, UnidadMedida)
                    .AddColumnValue(TABLA.PRECIO_UNITARIO, PrecioUnitario)
                    .AddColumnValue(TABLA.BONIFICACION_PERCENT, BonificacionPercent)
                    .AddColumnValue(TABLA.CUOTA_ID, CuotaId)
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
                    .AddColumnValue(TABLA.CODIGO, Codigo)
                    .AddColumnValue(TABLA.PRODUCTO_SERVICIO, ProductoServicio)
                    .AddColumnValue(TABLA.CANTIDAD, Cantidad)
                    .AddColumnValue(TABLA.UNIDAD_MEDIDA, UnidadMedida)
                    .AddColumnValue(TABLA.PRECIO_UNITARIO, PrecioUnitario)
                    .AddColumnValue(TABLA.BONIFICACION_PERCENT, BonificacionPercent)
                    .AddColumnValue(TABLA.CUOTA_ID, CuotaId)
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
                    Dim tmp As New AfipFacturaDetalle
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.FacturaId = CLng(.GetQueryData(TABLA.FACTURA_ID))
                    tmp.Codigo = CStr(.GetQueryData(TABLA.CODIGO))
                    tmp.ProductoServicio = CStr(.GetQueryData(TABLA.PRODUCTO_SERVICIO))
                    tmp.Cantidad = CInt(.GetQueryData(TABLA.CANTIDAD))
                    tmp.UnidadMedida = CInt(.GetQueryData(TABLA.UNIDAD_MEDIDA))
                    tmp.PrecioUnitario = CDec(.GetQueryData(TABLA.PRECIO_UNITARIO))
                    tmp.BonificacionPercent = CDec(.GetQueryData(TABLA.BONIFICACION_PERCENT))
                    tmp.CuotaId = CLng(.GetQueryData(TABLA.CUOTA_ID))
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



