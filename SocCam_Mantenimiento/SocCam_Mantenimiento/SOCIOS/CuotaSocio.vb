Imports System.Data.SqlClient
Imports helix

Public Class CuotaSocio

    Dim ConsoleOut As New ConsoleOut
    Public Property id As Integer = 0
    Public Property socioID As Integer = 0
    Public Property PlanID As Integer = 0
    Public Property Operacion As Integer = 0

    ''' <summary>
    ''' Periodicidad de la cuota, si es mensual, bimestral, etc
    ''' </summary>
    ''' <value>Byte indicando la periodicidad</value>
    ''' <returns>La periodicidad de la cuota</returns>
    ''' <remarks></remarks>
    Public Property Periodicidad As Byte = 0

    ''' <summary>
    ''' Periodo al que corresponde la cuota
    ''' </summary>
    ''' <value>Byte indicando el periodo</value>
    ''' <returns>El periodo al que corresponde la cuota</returns>
    ''' <remarks></remarks>
    Public Property Periodo As Byte = 0

    ''' <summary>
    ''' Año al que corresponde la cuota
    ''' </summary>
    ''' <value>Entero con indicando el año</value>
    ''' <returns>El año al que corresponde</returns>
    ''' <remarks></remarks>
    Public Property anio As Integer = Now.Year

    ''' <summary>
    ''' Estado en el que se encuentra la cuota
    ''' </summary>
    ''' <value>Uno de los 3 estados que puede tener la cuota, al dia, pendiente o moroso</value>
    ''' <returns>El estado de la cuota</returns>
    ''' <remarks></remarks>
    Public Property estado As ESTADO_SOCIO = ESTADO_SOCIO.AL_DIA

    ''' <summary>
    ''' Db ID del cobrador que cobra/genera la cuota
    ''' </summary>
    ''' <value>Entero con el Db ID del cobrador</value>
    ''' <returns>Entero con el Db ID</returns>
    ''' <remarks></remarks>
    Public Property cobradorID As Integer = 0

    ''' <summary>
    ''' Informacion adicional que se le agrega a la cuota
    ''' </summary>
    ''' <value>Cadena con comentario</value>
    ''' <returns></returns>
    ''' <remarks>El comentario</remarks>
    Public Property observaciones As String = ""

    ''' <summary>
    ''' Flag indicando si la cuota fue eliminada
    ''' </summary>
    ''' <value>TRUE si esta borrada, FALSE si no</value>
    ''' <returns>El estado de la cuota</returns>
    ''' <remarks></remarks>
    Public Property deleted As Boolean = False

    ''' <summary>
    ''' Monto de la cuota
    ''' </summary>
    ''' <value>Decimal con el valor de la cuota por periodo</value>
    ''' <returns>El valor de la cuota</returns>
    ''' <remarks></remarks>
    Public Property monto As Decimal = 0

    ''' <summary>
    ''' ID de la campaña a la que pertenece la cuota
    ''' </summary>
    ''' <returns></returns>
    Public Property Campania As Integer = 0

    ''' <summary>
    ''' ID del movimiento de CC
    ''' </summary>
    ''' <returns></returns>
    Public Property MovimientoCC As Decimal = 0

    ''' <summary>
    ''' El motor de base datos
    ''' </summary>
    ''' <returns></returns>
    Public Property sqle As New SQLEngine

    Public Property SearchResult As New List(Of CuotaSocio)

    Public Sub New()
    End Sub

    Public Sub New(ByVal tSqle As SQLEngine)
        Me.sqle.RequireCredentials = tSqle.RequireCredentials
        Me.sqle.Username = tSqle.Username
        Me.sqle.Password = tSqle.Password
        Me.sqle.dbType = tSqle.dbType
        Me.sqle.Path = tSqle.Path
        Me.sqle.DatabaseName = tSqle.DatabaseName
        If tSqle.IsStarted Then
            Me.sqle.ColdBoot()
        Else
            Me.sqle.Start()
        End If
    End Sub

    Public Sub InjectSql(ByVal tSqle As SQLEngine)
        Me.sqle.RequireCredentials = tSqle.RequireCredentials
        Me.sqle.Username = tSqle.Username
        Me.sqle.Password = tSqle.Password
        Me.sqle.dbType = tSqle.dbType
        Me.sqle.Path = tSqle.Path
        Me.sqle.DatabaseName = tSqle.DatabaseName
        If tSqle.IsStarted Then
            Me.sqle.ColdBoot()
        Else
            Me.sqle.Start()
        End If
    End Sub


    Public Sub Reset()
        monto = 0
        deleted = False
        observaciones = ""
        cobradorID = 0
        estado = ESTADO_SOCIO.AL_DIA
        anio = Now.Year
        Periodo = 0
        Periodicidad = 0
        PlanID = 0
        socioID = 0
        Campania = 0
        id = 0
        MovimientoCC = 0
    End Sub


    ''' <summary>
    ''' Fecha en que la cuota pasa de estado pendiente a vencida
    ''' </summary>
    ''' <value>Fecha de vencimiento de la cuota</value>
    ''' <returns>La fecha de vencimiento</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property fechaVencimiento As Date
        Get
            Return GenerateVencimiento(Periodo, anio, Periodicidad)
        End Get
    End Property

    Public Property fechaPago As Date

    ''' <summary>
    ''' Tipos de estados que se puede encontrar socio/cuota
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ESTADO_SOCIO
        AL_DIA = 0
        PENDIENTE = 1
        MOROSO = 2
    End Enum

    ''' <summary>
    ''' Guarda la cuota en la base de datos
    ''' </summary>
    ''' <param name="sqle">Motor de base de datos</param>
    ''' <param name="editMode">Modo de guardado, nuevo o edicion</param>
    ''' <returns>El ultimo</returns>
    ''' <remarks></remarks>
    Public Function Save(ByVal sqle As SQLEngine, ByVal editMode As Byte) As Integer
        Try
            Select Case editMode
                Case 0
                    With sqle.Insert
                        .Reset()
                        .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
                        .AddColumnValue(TABLA_PAGO_SOCIOS.SOCIO, socioID)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.ANIO, anio)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.PERIODO, Periodo)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.PERIODICIDAD, Periodicidad)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.FECHA_VENCIMIENTO, fechaVencimiento)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.PLAN, PlanID)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.OBSERVACIONES, observaciones)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.ESTADO, estado)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.OPERACION, Operacion)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.BUSQUEDA_VENCIMIENTO, GetVencimientoSearchable(fechaVencimiento))

                        If estado = ESTADO_SOCIO.AL_DIA Then
                            .AddColumnValue(TABLA_PAGO_SOCIOS.FECHA_PAGO, fechaPago)
                            .AddColumnValue(TABLA_PAGO_SOCIOS.BUSQUEDA_PAGO, GetVencimientoSearchable(fechaPago))
                        Else
                            .AddColumnValue(TABLA_PAGO_SOCIOS.FECHA_PAGO, DBNull.Value)
                        End If

                        .AddColumnValue(TABLA_PAGO_SOCIOS.COBRADOR, cobradorID)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.MODIFICADO, Now)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.DELETED, deleted)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.MONTO, monto)

                        .AddColumnValue(TABLA_PAGO_SOCIOS.MOVIMIENTO_CC, MovimientoCC)

                        Dim newIndex As Integer

                        If .Insert(newIndex) Then
                            Me.id = newIndex
                            Return newIndex     ' Si guardo bien retornar el ultimo ID
                        Else
                            Return 0           ' Si no, retornar flag de error
                        End If
                    End With
                Case 1
                    With sqle.Update
                        .Reset()
                        .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
                        .AddColumnValue(TABLA_PAGO_SOCIOS.SOCIO, socioID)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.ANIO, anio)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.PERIODO, Periodo)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.PERIODICIDAD, Periodicidad)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.FECHA_VENCIMIENTO, fechaVencimiento)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.PLAN, PlanID)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.OBSERVACIONES, observaciones)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.ESTADO, estado)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.OPERACION, Operacion)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.BUSQUEDA_VENCIMIENTO, GetVencimientoSearchable(fechaVencimiento))

                        If estado = ESTADO_SOCIO.AL_DIA Then
                            .AddColumnValue(TABLA_PAGO_SOCIOS.FECHA_PAGO, fechaPago)
                            .AddColumnValue(TABLA_PAGO_SOCIOS.BUSQUEDA_PAGO, GetVencimientoSearchable(fechaPago))
                        Else
                            .AddColumnValue(TABLA_PAGO_SOCIOS.FECHA_PAGO, DBNull.Value)
                        End If

                        .AddColumnValue(TABLA_PAGO_SOCIOS.COBRADOR, cobradorID)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.MODIFICADO, Now)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.DELETED, deleted)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.MONTO, monto)
                        .AddColumnValue(TABLA_PAGO_SOCIOS.MOVIMIENTO_CC, MovimientoCC)

                        .WHEREstring = TABLA_PAGO_SOCIOS.ID & " = ?"
                        .AddWHEREparam(id)

                        If .Update() Then Return id Else Return 0
                    End With
                Case Else
                    Return 0
            End Select
        Catch ex As Exception
            ConsoleOut.Print("Error al guardar la cuota: " & ex.Message)
        End Try
    End Function

    Public Function Update(ByVal sqle As SQLEngine) As Integer
        Try
            Dim connectionString As String = "Server= " + My.Computer.Name & "\" & "SQLEXPRESS" + ";Database=" + "soccam" + ";User Id=" + "soccam_user" + ";Password=" + "1Aleonardo" + ";"
            Dim query As String = "UPDATE " + TABLA_PAGO_SOCIOS.TABLA_NOMBRE + "
    SET 
    " + TABLA_PAGO_SOCIOS.MOVIMIENTO_CC + " = " + MovimientoCC.ToString + "
    WHERE " + TABLA_PAGO_SOCIOS.ID + " = " + id.ToString + ";"
            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    connection.Open()
                    command.ExecuteReader()
                    connection.Close()
                End Using
            End Using
        Catch ex As Exception
            ConsoleOut.Print("Error al guardar la cuota: " & ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Busca todas las cuotas de un usuario
    ''' </summary>
    ''' <param name="sqle">El motor de base datos</param>
    ''' <param name="userID">El ID del usuario que se quiere recuperar las cuotas</param>
    ''' <returns>True si la busqueda fue exitosa, FALSE si fallo</returns>
    Public Function GetCuotasByUser(ByRef sqle As SQLEngine, ByVal userID As Integer) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE

            ' Unir las tablas pagos, cobradores y plan
            .AddFirstJoin(TABLA_PAGO_SOCIOS.TABLA_NOMBRE, TABLA_COBRADORES.TABLA_NOMBRE, TABLA_PAGO_SOCIOS.COBRADOR, TABLA_COBRADORES.ID)
            '.AddFirstJoin(TABLA_PAGO_SOCIOS.TABLA_NOMBRE, TABLA_COBRADORES.TABLA_NOMBRE, TABLA_PAGO_SOCIOS.COBRADOR, TABLA_COBRADORES.ID)
            '.AddNestedJoin(TABLA_TIPO_SOCIO.TABLA_NOMBRE, TABLA_PAGO_SOCIOS.PLAN, TABLA_TIPO_SOCIO.ID)

            .AddSelectColumn(TABLA_PAGO_SOCIOS.ID)
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ANIO)
            .AddSelectColumn(TABLA_PAGO_SOCIOS.PERIODO)
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ESTADO)
            .AddSelectColumn(TABLA_PAGO_SOCIOS.FECHA_PAGO)
            .AddSelectColumn(TABLA_PAGO_SOCIOS.OBSERVACIONES)
            .AddSelectColumn(TABLA_PAGO_SOCIOS.MONTO)
            '.AddSelectColumn(TABLA_TIPO_SOCIO.PERIODICIDAD)
            .AddSelectColumn(TABLA_PAGO_SOCIOS.PERIODICIDAD)
            .AddSelectColumn(TABLA_COBRADORES.NOMBRE)
            .AddSelectColumn(TABLA_COBRADORES.APELLIDO)
            .AddSelectColumn(TABLA_PAGO_SOCIOS.MOVIMIENTO_CC)
            ' Buscar los que no esten borrados y los que correspondan al socio
            .WHEREstring = TABLA_PAGO_SOCIOS.DELETED & " = ? AND " & TABLA_PAGO_SOCIOS.SOCIO & " = ?"

            .AddOrderColumn(TABLA_PAGO_SOCIOS.ANIO, SQLEngineQuery.sortOrder.descending)
            .AddOrderColumn(TABLA_PAGO_SOCIOS.PERIODO, SQLEngineQuery.sortOrder.descending)

            .AddWHEREparam(0)
            .AddWHEREparam(userID)


            Return .Query

        End With
    End Function

    Public Function GetMontoImpago(ByVal sqle As SQLEngine, ByVal userID As Integer) As Decimal
        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ALL)
            .WHEREstring = TABLA_PAGO_SOCIOS.DELETED & " = ? AND "
            .WHEREstring &= TABLA_PAGO_SOCIOS.SOCIO & " = ? AND "
            .WHEREstring &= TABLA_PAGO_SOCIOS.ESTADO & " > ?"
            .AddWHEREparam(False)
            .AddWHEREparam(userID)
            .AddWHEREparam(0)

            Dim importe As Decimal = -1
            If .Query Then
                importe = 0
                While .QueryRead
                    importe += .GetQueryData(TABLA_PAGO_SOCIOS.MONTO)
                End While
            End If

            Return importe
        End With
    End Function

    Public Function GetListaCuotasImpagas(ByVal sqle As SQLEngine, ByVal userID As Integer) As List(Of CuotaSocio)
        Dim lstResult As New List(Of CuotaSocio)

        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ALL)
            .WHEREstring = TABLA_PAGO_SOCIOS.DELETED & " = ? AND "
            .WHEREstring &= TABLA_PAGO_SOCIOS.SOCIO & " = ? AND "
            .WHEREstring &= TABLA_PAGO_SOCIOS.ESTADO & " > ?"
            .AddWHEREparam(False)
            .AddWHEREparam(userID)
            .AddWHEREparam(1)
            .AddOrderColumn(TABLA_PAGO_SOCIOS.BUSQUEDA_VENCIMIENTO, SQLEngineQuery.sortOrder.ascending)


            Dim importe As Decimal = -1
            If .Query Then
                importe = 0
                Dim lstIndexes As New List(Of Integer)
                While .QueryRead
                    lstIndexes.Add(.GetQueryData(TABLA_PAGO_SOCIOS.ID))
                End While
                ' FIX DE MIERDA que resetea el queryread
                For Each indx In lstIndexes
                    Dim cuota As New CuotaSocio
                    cuota.LoadMe(sqle, indx)
                    lstResult.Add(cuota)
                Next
            End If

            Return lstResult
        End With

    End Function




    ''' <summary>
    ''' Genera la fecha de vencimiento de un cuota
    ''' </summary>
    ''' <param name="periodo">Periodo al que pertenece la cuota</param>
    ''' <param name="anio">Año al que pertenece la cuota</param>
    ''' <param name="periodicidad">Periodicidad de cobro de la cuota</param>
    ''' <returns>La fecha de cuando se va a vencer la cuota</returns>
    ''' <remarks>Para no tener problema con la fecha, arrancamos el primero de año y se suma los meses para estar al 1º del proximo periodo
    ''' y luego se retrocede 1 dia para estar en el ultimo dia del periodo correspondiente</remarks>
    Public Function GenerateVencimiento(ByVal periodo As Byte, ByVal anio As Integer, ByVal periodicidad As Byte) As Date

        Dim tmpDay As Byte = 1
        Dim tmpMonth As Byte = 1
        Dim tmpYear As Integer = anio

        Dim resultDate As Date

        If Not Date.TryParse("01/01/" & tmpYear, resultDate) Then
            Return Nothing
        End If

        Select Case periodicidad

            Case 0 ' Mensual
                resultDate = resultDate.AddMonths(periodo + 1).AddDays(-1)
            Case 1 ' Bimestral
                ' El vencimiento es todos los meses pares y no hay cambio de año
                resultDate = resultDate.AddMonths(periodo + 2).AddDays(-1)
            Case 2 ' Trimestral
                resultDate = resultDate.AddMonths(periodo + 3).AddDays(-1)
            Case 3 ' Cuatrimestral
                resultDate = resultDate.AddMonths(periodo + 4).AddDays(-1)
            Case 4 ' Semestral
                resultDate = resultDate.AddMonths(periodo + 6).AddDays(-1)
            Case 5 ' Anual
                ' Si la cuota es anual, el vencimiento es en febrero
                resultDate = resultDate.AddMonths(periodo + 1).AddDays(-1)
        End Select


        Return resultDate
    End Function

    Public Function GetVencimientoSearchable(ByVal vencimiento As Date) As Integer

        Dim tmpMes As String = ""
        Dim tmpDia As String = ""

        If vencimiento.Month < 10 Then
            tmpMes = "0" & vencimiento.Month.ToString
        Else
            tmpMes = vencimiento.Month.ToString
        End If

        If vencimiento.Day < 10 Then
            tmpDia = "0" & vencimiento.Day.ToString
        Else
            tmpDia = vencimiento.Day.ToString
        End If

        Return CInt(vencimiento.Year.ToString & tmpMes & tmpDia)
    End Function

    ''' <summary>
    ''' Elimina una cuota seleccionada
    ''' </summary>
    ''' <param name="sqle"></param>
    ''' <param name="cuotaID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Delete(ByVal sqle As SQLEngine, ByVal cuotaID As Integer) As Boolean
        With sqle.Update
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddColumnValue(TABLA_PAGO_SOCIOS.DELETED, True)
            .AddColumnValue(TABLA_PAGO_SOCIOS.MODIFICADO, Now())

            .WHEREstring = TABLA_PAGO_SOCIOS.ID & " = ?"
            .AddWHEREparam(cuotaID)

            Return .Update
        End With
    End Function

    ''' <summary>
    ''' Determina el periodo que corresponde a una cuota segun fecha y periodicidad
    ''' </summary>
    ''' <param name="fromDate">Fecha a determinar el periodo</param>
    ''' <param name="periodicidad">Periodicidad de la cuota</param>
    ''' <returns>El periodo al que corresponde la cuota</returns>
    ''' <remarks></remarks>
    Public Function DeterminePeriodFromDate(ByVal fromDate As Date, ByVal periodicidad As Byte) As Byte
        Select Case periodicidad
            Case 0
                ' Mensual
                Return fromDate.Month - 1
            Case 1
                ' Bimestral
                Return Math.Abs((fromDate.Month / 2) - (1 / 2))
            Case 2
                ' Trimestral
                Return Math.Abs((fromDate.Month / 3) - (1 / 3))
            Case 3
                ' Cuatrimestral
                Return Math.Abs((fromDate.Month / 4) - (1 / 4))
            Case 4
                ' Semestral
                Return Math.Abs((fromDate.Month / 6) - (1 / 6))
            Case Else
                ' Anual
                Return 0
        End Select
    End Function

    ''' <summary>
    ''' Verifica que la cuota no haya sido creada con anterioridad
    ''' </summary>
    ''' <param name="sqle">El motor de base de datos</param>
    ''' <param name="periodo">Periodo de la cuota a verificar</param>
    ''' <param name="periodicidad">Periodicidad de la cuota a verificar</param>
    ''' <param name="anio">Anio de la cuota a verificar</param>
    ''' <param name="socioID">ID de la cuota del socio a verificar</param>
    ''' <returns>TRUE si ya existe una cuota igual, FALSE si no</returns>
    ''' <remarks></remarks>
    Public Function CuotaExist(ByVal sqle As SQLEngine, ByVal periodo As Byte, ByVal periodicidad As Byte, ByVal anio As Integer, ByVal socioID As Integer, Optional ByRef cuotaId As Integer = 0) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ID)

            .WHEREstring = TABLA_PAGO_SOCIOS.PERIODO & " = ? AND " & TABLA_PAGO_SOCIOS.PERIODICIDAD & " = ? AND " & TABLA_PAGO_SOCIOS.ANIO & " = ? AND " & TABLA_PAGO_SOCIOS.SOCIO & " = ? AND " & TABLA_PAGO_SOCIOS.DELETED & " = ?"
            .AddWHEREparam(periodo)
            .AddWHEREparam(periodicidad)
            .AddWHEREparam(anio)
            .AddWHEREparam(socioID)
            .AddWHEREparam(False)

            If .Query Then
                If .RecordCount > 0 Then
                    .QueryRead()
                    cuotaId = .GetQueryData(TABLA_PAGO_SOCIOS.ID)
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End With
    End Function

    Public Function LoadMe(ByVal sqle As SQLEngine, ByVal myId As Integer) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ALL)
            .SimpleSearch(TABLA_PAGO_SOCIOS.ID, SQLEngineQuery.OperatorCriteria.Igual, myId)
            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    id = myId
                    socioID = .GetQueryData(TABLA_PAGO_SOCIOS.SOCIO)
                    PlanID = .GetQueryData(TABLA_PAGO_SOCIOS.PLAN)
                    anio = .GetQueryData(TABLA_PAGO_SOCIOS.ANIO)
                    Periodo = .GetQueryData(TABLA_PAGO_SOCIOS.PERIODO)
                    Periodicidad = .GetQueryData(TABLA_PAGO_SOCIOS.PERIODICIDAD)
                    If CStr(.GetQueryData(TABLA_PAGO_SOCIOS.FECHA_PAGO)) <> "" Then
                        fechaPago = .GetQueryData(TABLA_PAGO_SOCIOS.FECHA_PAGO)
                    Else
                        fechaPago = Nothing
                    End If

                    If CStr(.GetQueryData(TABLA_PAGO_SOCIOS.OPERACION)).Length > 0 Then
                        Operacion = .GetQueryData(TABLA_PAGO_SOCIOS.OPERACION)
                    Else
                        Operacion = 0
                    End If
                    observaciones = .GetQueryData(TABLA_PAGO_SOCIOS.OBSERVACIONES)
                    estado = .GetQueryData(TABLA_PAGO_SOCIOS.ESTADO)
                    cobradorID = .GetQueryData(TABLA_PAGO_SOCIOS.COBRADOR)
                    monto = .GetQueryData(TABLA_PAGO_SOCIOS.MONTO)
                    MovimientoCC = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.MOVIMIENTO_CC))
                    Return True
                End If
            End If
        End With
        Return False
    End Function

    Public Function LoadMe(ByVal sqle As SQLEngine, ByVal socioId As Integer, ByVal cuotaPeriodo As Byte, ByVal cuotaAnio As Integer) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ALL)
            .WHEREstring = TABLA_PAGO_SOCIOS.SOCIO & " = ? AND " & TABLA_PAGO_SOCIOS.PERIODO & " = ? AND " &
                           TABLA_PAGO_SOCIOS.ANIO & " = ? AND " & TABLA_PAGO_SOCIOS.DELETED & " = ?"
            .AddWHEREparam(socioId)
            .AddWHEREparam(cuotaPeriodo)
            .AddWHEREparam(cuotaAnio)
            .AddWHEREparam(False)

            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    id = .GetQueryData(TABLA_PAGO_SOCIOS.ID)
                    Me.socioID = .GetQueryData(TABLA_PAGO_SOCIOS.SOCIO)
                    PlanID = .GetQueryData(TABLA_PAGO_SOCIOS.PLAN)
                    anio = .GetQueryData(TABLA_PAGO_SOCIOS.ANIO)
                    Periodo = .GetQueryData(TABLA_PAGO_SOCIOS.PERIODO)
                    Periodicidad = .GetQueryData(TABLA_PAGO_SOCIOS.PERIODICIDAD)
                    If CStr(.GetQueryData(TABLA_PAGO_SOCIOS.FECHA_PAGO)) <> "" Then
                        fechaPago = .GetQueryData(TABLA_PAGO_SOCIOS.FECHA_PAGO)
                    Else
                        fechaPago = Nothing
                    End If
                    observaciones = .GetQueryData(TABLA_PAGO_SOCIOS.OBSERVACIONES)
                    estado = .GetQueryData(TABLA_PAGO_SOCIOS.ESTADO)
                    cobradorID = .GetQueryData(TABLA_PAGO_SOCIOS.COBRADOR)
                    monto = .GetQueryData(TABLA_PAGO_SOCIOS.MONTO)
                    If CStr(.GetQueryData(TABLA_PAGO_SOCIOS.OPERACION)).Length > 0 Then
                        Operacion = .GetQueryData(TABLA_PAGO_SOCIOS.OPERACION)
                    Else
                        Operacion = 0
                    End If
                    MovimientoCC = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.MOVIMIENTO_CC))
                    Return True
                End If
            End If
        End With
        Return False
    End Function

    Public Function LoadAll(ByVal sqle As SQLEngine, ByRef dt As DataTable) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ALL)
            .WHEREstring = TABLA_PAGO_SOCIOS.DELETED & " = ?"
            .AddWHEREparam(False)
            Return .Query(True, dt)
        End With
    End Function

    Public Function LoadAll(ByVal sqle As SQLEngine, ByRef dt As DataTable, ByVal estadoOperator As String,
                            ByVal cuotaEstado As ESTADO_SOCIO, ByVal socioID As Integer) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ALL)
            .WHEREstring = TABLA_PAGO_SOCIOS.DELETED & " = ? AND " & TABLA_PAGO_SOCIOS.SOCIO & " = ? AND " & TABLA_PAGO_SOCIOS.ESTADO & " " & estadoOperator & " ?"
            .AddWHEREparam(False)
            .AddWHEREparam(socioID)
            .AddWHEREparam(cuotaEstado)

            Return .Query(True, dt)
        End With
    End Function

    Public Function GetTotalPagarSemestre(ByVal sqle As SQLEngine,
                                          ByVal semestre As Byte, ByVal anio As Integer,
                                          ByVal socioID As Integer, Optional ByRef cantidadCuotas As Integer = 0,
                                          Optional ByVal ignorarFacturadas As Boolean = False,
                                          Optional ByVal importeCuotaSocial As Decimal = -1) As Decimal

        Dim mesBase As Byte = 0
        Dim mesHasta As Byte = 5

        If semestre = 1 Then
            mesBase = 6
            mesHasta = 11
        End If

        Dim soc As New Socio
        soc.LoadMe(sqle, socioID)

        Dim socTipo As New SocioTipo
        socTipo.sqle = sqle
        socTipo.LoadMe(soc.Tipo)

        Dim total As Decimal = 0


        For i = mesBase To mesHasta
            Dim cuota As New CuotaSocio
            If cuota.CuotaExist(sqle, i, 0, anio, socioID, cuota.id) Then
                cuota.LoadMe(sqle, cuota.id)
                If ignorarFacturadas Then
                    If cuota.observaciones.StartsWith("FC") Then Continue For
                End If
                If cuota.estado >= ESTADO_SOCIO.PENDIENTE Then
                    total += If(importeCuotaSocial > 0, importeCuotaSocial, cuota.monto)
                    cantidadCuotas += 1
                End If
            Else
                total += socTipo.importe
                cantidadCuotas += 1
            End If
        Next

        Return total
    End Function




    Public Function LoadSemestre(ByVal semestre As Integer, ByVal anio As Integer, Optional idSocio As Integer = -1) As List(Of CuotaSocio)
        Dim tmpLst As New List(Of CuotaSocio)
        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .WHEREstring = TABLA_PAGO_SOCIOS.PERIODO & " >= ? AND " & TABLA_PAGO_SOCIOS.PERIODO & " <= ? AND "
            If semestre = 0 Then
                .AddWHEREparam(0)
                .AddWHEREparam(5)
            Else
                .AddWHEREparam(6)
                .AddWHEREparam(11)
            End If
            .WHEREstring &= TABLA_PAGO_SOCIOS.ANIO & " = ? AND " & TABLA_PAGO_SOCIOS.DELETED & " = ?"
            .AddWHEREparam(anio)
            .AddWHEREparam(False)

            If idSocio > 0 Then
                .WHEREstring &= $" AND {TABLA_PAGO_SOCIOS.SOCIO} = ?"
                .AddWHEREparam(idSocio)
            End If

            .AddOrderColumn(TABLA_PAGO_SOCIOS.PERIODO, SQLEngineQuery.sortOrder.ascending)

            Dim dt As New DataTable
            Dim dtr As DataTableReader

            If .Query(True, dt) Then
                dtr = dt.CreateDataReader
                While dtr.Read
                    Dim tmpCuota As New CuotaSocio
                    tmpCuota.LoadMe(sqle, dtr(TABLA_PAGO_SOCIOS.ID))
                    tmpLst.Add(tmpCuota)
                End While
                Return tmpLst
            End If

        End With

        Return tmpLst
    End Function



    Public Overrides Function ToString() As String
        Dim out As String = "|  "
        out &= Me.id & "  |  "
        out &= Me.socioID & "  |  "
        out &= Me.PlanID & "  |  "
        out &= Me.anio & "  |  "
        out &= Me.Periodo & "  |  "
        out &= Me.Periodicidad & "  |  "
        out &= Me.fechaVencimiento & "  |  "
        out &= Me.fechaPago & "  |  "
        out &= Me.monto & "  |  "
        out &= Me.deleted & "  |  "
        out &= Me.Operacion & "  |"

        Return out

    End Function

    ''' <summary>
    ''' Indica si el mes esta dentro de un periodo dado
    ''' </summary>
    ''' <param name="periodicity">Periodicidad de la cuota</param>
    ''' <param name="period">Periodo a comparar</param>
    ''' <param name="mes">Mes a comparar</param>
    ''' <returns>True si el mes esta dentro del periodo, False si no</returns>
    Public Shared Function PeriodoContieneMes(ByVal periodicity As Integer, ByVal period As Integer, ByVal mes As Integer) As Boolean
        ' Si es anual logicamente esta dentro del periodo
        If periodicity = 5 Then Return True
        mes += 1

        ' Mensual
        Select Case periodicity
            Case 0
                ' Mensual
                If period = (mes - 1) Then
                    Return True
                Else
                    Return False
                End If
            Case 1
                ' Bimestral
                If period = Math.Truncate((Math.Abs((mes / 2) - (1 / 2)))) Then
                    Return True
                Else
                    Return False
                End If
            Case 2
                ' Trimestral
                If period = Math.Truncate((Math.Abs((mes / 3) - (1 / 3)))) Then
                    Return True
                Else
                    Return False
                End If
            Case 3
                ' Cuatrimestral
                If period = Math.Truncate((Math.Abs((mes / 4) - (1 / 4)))) Then
                    Return True
                Else
                    Return False
                End If
            Case 4
                ' Semestral
                If period = Math.Truncate((Math.Abs((mes / 6) - (1 / 6)))) Then
                    Return True
                Else
                    Return False
                End If
        End Select


        Return False

    End Function


    Public Function GetTotalFacturas(ByVal desde As Date, ByVal hasta As Date, Optional ByVal caja As String = "", Optional ByVal pagas As Boolean = True) As Integer
        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddFirstJoin(TABLA_PAGO_SOCIOS.TABLA_NOMBRE, TABLA_PAGO_IMPUESTO.TABLA_NOMBRE, TABLA_PAGO_SOCIOS.OPERACION, TABLA_PAGO_IMPUESTO.ID)
            .WHEREstring &= $"({TABLA_PAGO_SOCIOS.BUSQUEDA_PAGO} BETWEEN ? AND ?) AND {TABLA_PAGO_SOCIOS.DELETED} = ?"
            .AddWHEREparam(Utils.DateTo8601(desde))
            .AddWHEREparam(Utils.DateTo8601(hasta))
            .AddWHEREparam(False)
            If caja.Length > 0 Then
                .WHEREstring &= $" AND {TABLA_PAGO_IMPUESTO.CAJA} LIKE = ?"
                .AddWHEREparam(caja)
            End If
            .WHEREstring &= $" GROUP BY {TABLA_PAGO_SOCIOS.OBSERVACIONES}"
            If .Query Then
                Return .RecordCount
            End If
        End With

        Return 0
    End Function

    Public Function GetNumeroFactura() As Long
        If Me.observaciones.StartsWith("FC") Then
            Return CLng(Me.observaciones.Split("-")(1))
        Else
            Return 0
        End If
    End Function


    ''' <summary>
    ''' Busca todas las cuotas de un periodo e importe especifico
    ''' </summary>
    ''' <param name="mes">Mes en base 0</param>
    ''' <param name="anio">Anio</param>
    ''' <param name="importe">Importe a buscar</param>
    ''' <returns>Listado de cuotas</returns>
    Public Function LoadPeriodo(ByVal mes As Integer, ByVal anio As Integer, ByVal importe As Decimal, ByRef lst As List(Of CuotaSocio)) As Integer
        lst.Clear()

        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .WHEREstring = $"{TABLA_PAGO_SOCIOS.PERIODICIDAD} = ? AND {TABLA_PAGO_SOCIOS.PERIODO} = ? AND {TABLA_PAGO_SOCIOS.ANIO} = ? AND {TABLA_PAGO_SOCIOS.DELETED} = ? AND {TABLA_PAGO_SOCIOS.ESTADO} > ? AND {TABLA_PAGO_SOCIOS.MONTO} = ?"
            .AddWHEREparam(0)
            .AddWHEREparam(mes)
            .AddWHEREparam(anio)
            .AddWHEREparam(False)
            .AddWHEREparam(0)
            .AddWHEREparam(importe)


            .AddOrderColumn(TABLA_PAGO_SOCIOS.PERIODO, SQLEngineQuery.sortOrder.ascending)

            Dim dt As New DataTable
            Dim dtr As DataTableReader

            If .Query(True, dt) Then
                dtr = dt.CreateDataReader
                While dtr.Read
                    Dim tmpCuota As New CuotaSocio
                    tmpCuota.LoadMe(sqle, dtr(TABLA_PAGO_SOCIOS.ID))
                    lst.Add(tmpCuota)
                End While
                Return lst.Count
            End If

        End With

        Return Nothing
    End Function

    Public Function LoadCuotasMismaFactura(ByVal numFactura As Integer, ByVal anioFactura As Integer) As Boolean
        With Me.sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ID)
            .WHEREstring = $"{TABLA_PAGO_SOCIOS.OBSERVACIONES} = ?  AND {TABLA_PAGO_SOCIOS.ANIO} = ? AND {TABLA_PAGO_SOCIOS.DELETED} = ?"
            .AddWHEREparam($"FC-{numFactura}")
            .AddWHEREparam(anioFactura)
            .AddWHEREparam(False)
            If .Query Then
                Me.SearchResult.Clear()
                While .QueryRead
                    Dim c As New CuotaSocio()
                    c.LoadMe(Me.sqle, CInt(.GetQueryData(0)))
                    Me.SearchResult.Add(c)
                End While
                Return True
            End If
        End With

        Return False
    End Function


    ''' <summary>
    ''' Carga una cuota virtual que corresponde al siguiente periodo al último generado
    ''' </summary>
    ''' <param name="idSocio">El ID del socio que corresponde la cuota</param>
    ''' <returns>0 si la pudo generar, menor que 0 si hubo un error</returns>
    Public Function CrearUltimaCuotaSinGenerar(Optional IdSocio As Integer = 0) As Integer
        Dim idSocioABuscar As Integer
        If IdSocio > 0 Then
            idSocioABuscar = IdSocio
        Else
            idSocioABuscar = Me.id
        End If

        With Me.sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn($"TOP 1 {TABLA_PAGO_SOCIOS.ALL}")
            .WHEREstring = $"{TABLA_PAGO_SOCIOS.SOCIO} = { .p(idSocioABuscar)} AND {TABLA_PAGO_SOCIOS.DELETED} = { .p(False)}"
            .AddOrderColumn(TABLA_PAGO_SOCIOS.BUSQUEDA_VENCIMIENTO, SQLEngineQuery.sortOrder.descending)

            If .Query Then
                If .RecordCount = 1 Then
                    .QueryRead()
                    Me.socioID = idSocioABuscar
                    Me.cobradorID = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.COBRADOR))
                    Me.Periodicidad = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.PERIODICIDAD))
                    Me.PlanID = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.PLAN))
                    Me.monto = CDec(.GetQueryData(TABLA_PAGO_SOCIOS.MONTO))


                    Me.id = 0
                    Me.deleted = False
                    Me.estado = 1
                    Me.fechaPago = Nothing
                    Me.observaciones = ""
                    Me.Operacion = 0
                    Me.MovimientoCC = 0

                    Me.Periodo = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.PERIODO))
                    Me.anio = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.ANIO))

                    CalcularProximoPeriodo()
                Else
                    ' Es la primera cuota
                    Dim soc As New SocioNT(Me.sqle)
                    If soc.LoadMe(idSocioABuscar) Then
                        Dim sector As New Cobrador(Me.sqle)
                        sector.LoadMe(Me.sqle, soc.Sector, True)

                        Dim planSocio As New SocioTipo(Me.sqle)
                        planSocio.LoadMe(soc.TipoSocio)

                        Me.socioID = idSocioABuscar
                        Me.cobradorID = sector.ID
                        Me.Periodicidad = planSocio.categoria
                        Me.PlanID = planSocio.id
                        Me.monto = planSocio.importe


                        Me.id = 0
                        Me.deleted = False
                        Me.estado = 1
                        Me.fechaPago = Nothing
                        Me.observaciones = ""
                        Me.Operacion = 0
                        Me.MovimientoCC = 0



                        Me.Periodo = Utils.GetPeriodoFromFecha(Now, Utils.GetMesesFromPeriodicidad(planSocio.periodicidad))
                        Me.anio = Now.Year

                        CalcularProximoPeriodo()

                    Else
                        Return -2
                    End If
                End If

                ' No hubo ningun error
                Return 0
            Else
                ' No pudo ejecutar la consulta
                Return -1
            End If
        End With
    End Function


    ''' <summary>
    ''' Carga o crea una cuota virtual a la cuota cargada
    ''' </summary>
    ''' <returns>0 si cargó una cuota existente, 1 si creo una cuota virtual. Un entero menor que 0 si hubo un error</returns>
    Public Function IrProximoPeriodo() As Integer
        CalcularProximoPeriodo()
        Dim tmpCuotaID As Integer
        If CuotaExist(Me.sqle, Me.Periodo, Me.Periodicidad, Me.anio, Me.socioID, tmpCuotaID) Then
            If Not Me.LoadMe(Me.sqle, tmpCuotaID) Then Return -1
            Return 0
        Else
            Return 1
        End If
    End Function

    ''' <summary>
    ''' Mueve el periodo o año segun periodicidad y periodo cargado
    ''' </summary>
    Private Sub CalcularProximoPeriodo()
        Select Case Me.Periodicidad
            Case 0
                ' MENSUAL
                If Me.Periodo = 11 Then
                    Me.Periodo = 0
                    Me.anio += 1
                Else
                    Me.Periodo += 1
                End If
            Case 1
                ' BIMESTRAL
                If Me.Periodo = 5 Then
                    Me.Periodo = 0
                    Me.anio += 1
                Else
                    Me.Periodo += 1
                End If
            Case 2
                ' TRIMESTRAL
                If Me.Periodo = 3 Then
                    Me.Periodo = 0
                    Me.anio += 1
                Else
                    Me.Periodo += 1
                End If
            Case 3
                ' CUATRIMESTRAL
                If Me.Periodo = 2 Then
                    Me.Periodo = 0
                    Me.anio += 1
                Else
                    Me.Periodo += 1
                End If
            Case 4
                ' SEMESTRAL
                If Me.Periodo = 1 Then
                    Me.Periodo = 0
                    Me.anio += 1
                Else
                    Me.Periodo += 1
                End If
            Case 5
                ' ANUAL
                Me.Periodo = 0
                Me.anio += 1
        End Select
    End Sub


    Public Function BuscarUltimaCuota(Optional IdSocio As Integer = 0) As Integer
        Dim idSocioABuscar As Integer
        If IdSocio > 0 Then
            idSocioABuscar = IdSocio
        Else
            idSocioABuscar = Me.id
        End If

        With Me.sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn($"TOP 1 {TABLA_PAGO_SOCIOS.ALL}")
            .WHEREstring = $"{TABLA_PAGO_SOCIOS.SOCIO} = { .p(idSocioABuscar)} AND {TABLA_PAGO_SOCIOS.DELETED} = { .p(False)}"
            .AddOrderColumn(TABLA_PAGO_SOCIOS.BUSQUEDA_VENCIMIENTO, SQLEngineQuery.sortOrder.descending)

            If .Query Then
                If .RecordCount = 1 Then
                    .QueryRead()
                    Me.socioID = idSocioABuscar
                    Me.cobradorID = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.COBRADOR))
                    Me.Periodicidad = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.PERIODICIDAD))
                    Me.PlanID = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.PLAN))
                    Me.monto = CDec(.GetQueryData(TABLA_PAGO_SOCIOS.MONTO))


                    Me.id = 0
                    Me.deleted = False
                    Me.estado = 1
                    Me.fechaPago = Nothing
                    Me.observaciones = ""
                    Me.Operacion = 0
                    Me.MovimientoCC = 0

                    Me.Periodo = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.PERIODO))
                    Me.anio = CInt(.GetQueryData(TABLA_PAGO_SOCIOS.ANIO))

                Else
                    ' Es la primera cuota
                    Dim soc As New SocioNT(Me.sqle)
                    If soc.LoadMe(idSocioABuscar) Then
                        Dim sector As New Cobrador(Me.sqle)
                        sector.LoadMe(Me.sqle, soc.Sector, True)

                        Dim planSocio As New SocioTipo(Me.sqle)
                        planSocio.LoadMe(soc.TipoSocio)

                        Me.socioID = idSocioABuscar
                        Me.cobradorID = sector.ID
                        Me.Periodicidad = planSocio.categoria
                        Me.PlanID = planSocio.id
                        Me.monto = planSocio.importe


                        Me.id = 0
                        Me.deleted = False
                        Me.estado = 1
                        Me.fechaPago = Nothing
                        Me.observaciones = ""
                        Me.Operacion = 0
                        Me.MovimientoCC = 0



                        Me.Periodo = Utils.GetPeriodoFromFecha(Now, Utils.GetMesesFromPeriodicidad(planSocio.periodicidad))
                        Me.anio = Now.Year
                    Else
                        Return -2
                    End If
                End If

                ' No hubo ningun error
                Return 0
            Else
                ' No pudo ejecutar la consulta
                Return -1
            End If
        End With
    End Function


    Public Function LoadPorMovimiento(ByVal movimientoID As Integer) As List(Of CuotaSocio)
        Dim res As New List(Of CuotaSocio)
        With Me.sqle.Query
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ID)
            .WHEREstring = $"{TABLA_PAGO_SOCIOS.MOVIMIENTO_CC} = { .p(movimientoID)} AND
                             {TABLA_PAGO_SOCIOS.DELETED} = { .p(False)}"
            If .Query Then
                While .QueryRead
                    Dim c As New CuotaSocio(Me.sqle)
                    c.LoadMe(c.sqle, .GetQueryData(TABLA_PAGO_SOCIOS.ID))
                    res.Add(c)
                End While
            End If
        End With

        Return res
    End Function

    Public Function LoadImpagas() As Integer
        Dim res As Integer = -1
        With Me.sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ID)
            .WHEREstring = $"{TABLA_PAGO_SOCIOS.SOCIO} = { .p(Me.socioID)} AND
                             {TABLA_PAGO_SOCIOS.ESTADO} <> { .p(0)} AND
                             {TABLA_PAGO_SOCIOS.DELETED} = { .p(False)}"
            .AddOrderColumn(TABLA_PAGO_SOCIOS.ID, SQLEngineQuery.sortOrder.descending)
            If .Query Then
                SearchResult.Clear()

                While .QueryRead
                    Dim c As New CuotaSocio(Me.sqle)
                    c.LoadMe(c.sqle, .GetQueryData(TABLA_PAGO_SOCIOS.ID))
                    SearchResult.Add(c)
                End While

                Return SearchResult.Count
            Else
                Err.Raise(515, "CuotaSocio", "No pudo buscar la DB")
            End If
        End With

        Return res
    End Function


End Class
