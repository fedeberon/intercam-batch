Imports System.Windows.Forms
Imports helix

Public Class Campania
    Public Property Sqle As New SQLEngine
    Public Property SearchResult As New List(Of Campania)

    Public Property Id As Long = 0
    Public Property Nombre As String = ""
    Public Property CuotasBonificadas As Integer = 0
    Public Property PorcentajeBonificado As Decimal = 0
    Public Property Deleted As Boolean = False
    Public Property Modificado As Date = Now
    Public Property Descripcion As String

    Public Property ListaCuotasCampaña As New List(Of CuotaSocio)

    Public Enum Guardar
        NUEVO = 0
        EDITAR = 1
    End Enum


    Public Structure TABLA
        Const TABLA_NOMBRE As String = "Campanias"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const NOMBRE As String = TABLA_NOMBRE & "_nombre"
        Const CUOTAS_BONIFICADAS As String = TABLA_NOMBRE & "_cuotasBonificadas"
        Const PORCENTAJE_BONIFICADO As String = TABLA_NOMBRE & "_porcentajeBonificado"
        Const DESCRIPCION As String = TABLA_NOMBRE & "_descripcion"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & NOMBRE & ", " & CUOTAS_BONIFICADAS & ", " & PORCENTAJE_BONIFICADO & ", " & DELETED & ", " & MODIFICADO & ", " & DESCRIPCION
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

    Public Function InjectSql(ByVal iSqle As SQLEngine) As Boolean
        Me.Sqle.RequireCredentials = iSqle.RequireCredentials
        Me.Sqle.Username = iSqle.Username
        Me.Sqle.Password = iSqle.Password
        Me.Sqle.dbType = iSqle.dbType
        Me.Sqle.Path = iSqle.Path
        Me.Sqle.DatabaseName = iSqle.DatabaseName
        If iSqle.IsStarted Then
            Return Me.Sqle.ColdBoot()
        Else
            Return Me.Sqle.Start()
        End If

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
                    Nombre = CStr(.GetQueryData(TABLA.NOMBRE))
                    CuotasBonificadas = CInt(.GetQueryData(TABLA.CUOTAS_BONIFICADAS))
                    PorcentajeBonificado = CDec(.GetQueryData(TABLA.PORCENTAJE_BONIFICADO))
                    Descripcion = CStr(.GetQueryData(TABLA.DESCRIPCION))
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
                    .AddColumnValue(TABLA.NOMBRE, Nombre)
                    .AddColumnValue(TABLA.CUOTAS_BONIFICADAS, CuotasBonificadas)
                    .AddColumnValue(TABLA.PORCENTAJE_BONIFICADO, PorcentajeBonificado)
                    .AddColumnValue(TABLA.DESCRIPCION, Descripcion)
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
                    .AddColumnValue(TABLA.NOMBRE, Nombre)
                    .AddColumnValue(TABLA.CUOTAS_BONIFICADAS, CuotasBonificadas)
                    .AddColumnValue(TABLA.PORCENTAJE_BONIFICADO, PorcentajeBonificado)
                    .AddColumnValue(TABLA.DESCRIPCION, Descripcion)
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
                    Dim tmp As New Campania
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.Nombre = CStr(.GetQueryData(TABLA.NOMBRE))
                    tmp.CuotasBonificadas = CInt(.GetQueryData(TABLA.CUOTAS_BONIFICADAS))
                    tmp.PorcentajeBonificado = CDec(.GetQueryData(TABLA.PORCENTAJE_BONIFICADO))
                    tmp.Descripcion = CStr(.GetQueryData(TABLA.DESCRIPCION))
                    tmp.Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    tmp.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))
                    SearchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function


    ''' <summary>
    ''' Carga un listado de campanias en un combobox. El tag es una lista de campañas. 
    ''' El primer elemento siempre es NINGUNA y el ID del primer objeto es 0
    ''' </summary>
    ''' <param name="cmb">El combo a cargar el listado</param>
    ''' <returns>-1 si falló, un entero con la cantidad de items cargados</returns>
    Public Function LoadCombo(ByRef cmb As ComboBox, ByVal cargarNinguna As Boolean) As Integer
        Dim lst As New List(Of Campania)
        cmb.Items.Clear()

        If cargarNinguna Then
            cmb.Items.Add("NINGUNA")
            lst.Add(New Campania)
        End If


        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .WHEREstring = $"{TABLA.DELETED} = { .p(False)}"
            If .Query() Then
                While .QueryRead
                    Dim tmp As New Campania
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.Nombre = CStr(.GetQueryData(TABLA.NOMBRE))
                    tmp.CuotasBonificadas = CInt(.GetQueryData(TABLA.CUOTAS_BONIFICADAS))
                    tmp.PorcentajeBonificado = CDec(.GetQueryData(TABLA.PORCENTAJE_BONIFICADO))
                    tmp.Descripcion = CStr(.GetQueryData(TABLA.DESCRIPCION))
                    tmp.Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    tmp.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))

                    lst.Add(tmp)
                    cmb.Items.Add(tmp.Nombre)
                    cmb.Tag = lst

                End While
                Return lst.Count
            End If
        End With

        Return -1
    End Function

    ''' <summary>
    ''' Genera la lista de cuotas bonificadas segun la campaña
    ''' </summary>
    ''' <param name="socioID">El ID de socio a generar la campaña</param>
    ''' <returns>El importe total de todas las cuotas, un entero negativo si falló</returns>
    Public Function CargarListaDeCuotas(ByVal socioID As Integer) As Decimal
        If Me.Id <> 0 Then
            Dim cuota As New CuotaSocio(Me.Sqle)
            Dim totalImporteCuotas As Decimal = 0
            If cuota.BuscarUltimaCuota(socioID) = 0 Then
                For i = 1 To Me.CuotasBonificadas
                    Dim nuevaCuota As New CuotaSocio
                    cuota.IrProximoPeriodo()
                    nuevaCuota.anio = cuota.anio
                    nuevaCuota.cobradorID = cuota.cobradorID
                    nuevaCuota.deleted = cuota.deleted
                    nuevaCuota.estado = cuota.estado
                    nuevaCuota.fechaPago = cuota.fechaPago
                    nuevaCuota.id = cuota.id
                    nuevaCuota.monto = cuota.monto
                    nuevaCuota.observaciones = cuota.observaciones
                    nuevaCuota.Operacion = cuota.Operacion
                    nuevaCuota.Periodicidad = cuota.Periodicidad
                    nuevaCuota.Periodo = cuota.Periodo
                    nuevaCuota.PlanID = cuota.PlanID
                    nuevaCuota.SearchResult = cuota.SearchResult
                    nuevaCuota.socioID = cuota.socioID

                    nuevaCuota.monto = Utils.CalcularMenosPorcentaje(cuota.monto, Me.PorcentajeBonificado)
                    totalImporteCuotas += nuevaCuota.monto
                    Me.ListaCuotasCampaña.Add(nuevaCuota)
                Next

                Return totalImporteCuotas
            Else
                Return -2
            End If
        Else
            Return -1
        End If

    End Function
End Class