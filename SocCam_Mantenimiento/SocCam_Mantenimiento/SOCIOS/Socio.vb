Imports helix
''' <summary>
''' Clase administradora de socio
''' </summary>
''' <remarks></remarks>
Public Class Socio
    Public Property Nombre As String = ""

    Public Property Apellido As String = ""

    Public Property Nacionalidad As String = ""

    Public Property DNI As String = ""

    Public Property FechaNacimiento As Date

    Public Property Mail As String = ""

    Public Property CUIT As String = ""

    Public Property Firma As String = ""

    Public Property TipoFirma As String = ""

    Public Property Domicilio As String = ""

    Public Property Localidad As Integer = 813

    Public Property Telefono As String = ""

    Public Property Celular As String = ""

    Public Property OtroTelefono As String = ""

    Public Property FechaAceptacion As Date

    Public Property ActaNumero As Integer = 0

    Public Property Numero As Integer = 0

    Public Property PresentadoPor1 As Integer = 0

    Public Property PresentadoPor2 As Integer = 0

    Public Property Tipo As Integer = 0

    Public Property TipoString As String = ""

    Public Property Sector As Integer = 0

    Public Property MotivoBaja As String = ""

    Public Property InternalID As Integer

    Public Property Deleted As Boolean = False

    Public Property TieneCajaSeguridad As Boolean = False

    Public Property Gestion As Integer = 0

    Public Property Segmento As Integer = 0

    Public Property Rubro As Integer = 0

    Public Property Habilitacion As Integer = 0

    Public Property CondicionFiscal As Integer = 0

    Public Property TarjetaEntregada As Boolean = False

    Public Property TarjetaFechaEntrega As Long = 0

    Public Property Campania As Long = 0

    Public Property EnviarMail As Boolean = False


    ''' <summary>
    ''' Estado en que se encuentra: 0 = al dia, 1 = sin pagar el mes en curso, 2 = con deuda (30 dias o mas)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Estado As Byte = 0

    ''' <summary>
    ''' Fecha de ultima modificacion
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Modificado As Date = Now()




    ''' <summary>
    ''' Reinicia los valores por defecto
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()
        _Nombre = ""
        _Apellido = ""
        _Nacionalidad = ""
        _DNI = ""
        _FechaNacimiento = Now()
        _Mail = ""
        _CUIT = ""
        _Firma = ""
        _TipoFirma = ""
        _Domicilio = ""
        _Telefono = ""
        _Celular = ""
        _OtroTelefono = ""
        _FechaAceptacion = Now()
        _ActaNumero = 0
        _Numero = 0
        _PresentadoPor1 = 0
        _PresentadoPor2 = 0
        _Tipo = 0
        _TipoString = ""
        _Sector = 0
        _InternalID = 0
        _Deleted = False
        _TieneCajaSeguridad = False
        _MotivoBaja = ""
        _Estado = 0
        _Modificado = Now()
        _Gestion = 0
        _Segmento = 0
        _Rubro = 0
        _Habilitacion = 0
        _CondicionFiscal = 0
        _Localidad = 813
        _TarjetaEntregada = False
        _TarjetaFechaEntrega = 0
        _Campania = 0
        _EnviarMail = False
    End Sub


    ''' <summary>
    ''' Guarda la instancia actual en la base de datos
    ''' </summary>
    ''' <param name="sqlEngine">Motor de base de datos que se va a usar para guardar</param>
    ''' <returns>True si se guardo correctamente, False si no</returns>
    ''' <remarks>El motor pasa por valor para que todos los cambios que se realizen no afecte al resto del sistema</remarks>
    Public Function Save(ByVal sqlEngine As helix.SQLEngine, Optional getLastID As Boolean = True) As Boolean
        With sqlEngine.Insert
            .Reset()
            .TableName = TABLA_SOCIO.TABLA_NOMBRE
            .AddColumnValue(TABLA_SOCIO.NOMBRE, _Nombre)
            .AddColumnValue(TABLA_SOCIO.APELLIDO, _Apellido)
            .AddColumnValue(TABLA_SOCIO.NACIONALIDAD, _Nacionalidad)
            .AddColumnValue(TABLA_SOCIO.DNI, _DNI)
            .AddColumnValue(TABLA_SOCIO.FECHA_NACIMIENTO, _FechaNacimiento)
            .AddColumnValue(TABLA_SOCIO.CUIT, _CUIT)
            .AddColumnValue(TABLA_SOCIO.MAIL, _Mail)
            .AddColumnValue(TABLA_SOCIO.FIRMA, _Firma)
            .AddColumnValue(TABLA_SOCIO.TIPO_EMPRESA, _TipoFirma)
            .AddColumnValue(TABLA_SOCIO.DOMICILIO, _Domicilio)
            .AddColumnValue(TABLA_SOCIO.TELEFONO, _Telefono)
            .AddColumnValue(TABLA_SOCIO.TIPO_SOCIO, _Tipo)
            .AddColumnValue(TABLA_SOCIO.NUMERO, _Numero)
            .AddColumnValue(TABLA_SOCIO.FECHA_APROBACION, _FechaAceptacion)
            .AddColumnValue(TABLA_SOCIO.ACTA, _ActaNumero)
            .AddColumnValue(TABLA_SOCIO.SOCIO_PADRINO1, _PresentadoPor1)
            .AddColumnValue(TABLA_SOCIO.SOCIO_PADRINO2, _PresentadoPor2)
            .AddColumnValue(TABLA_SOCIO.SECTOR, _Sector)
            .AddColumnValue(TABLA_SOCIO.DELETED, _Deleted)
            .AddColumnValue(TABLA_SOCIO.CAJA_SEGURIDAD, _TieneCajaSeguridad)
            .AddColumnValue(TABLA_SOCIO.MOTIVO_BAJA, _MotivoBaja)
            .AddColumnValue(TABLA_SOCIO.ESTADO, _Estado)
            .AddColumnValue(TABLA_SOCIO.MODIFICADO, _Modificado)
            ' HACK TURNOS
            If Not sqlEngine.dbType = SQLEngine.dataBaseType.MYSQL Then
                .AddColumnValue(TABLA_SOCIO.CELULAR, _Celular)
                .AddColumnValue(TABLA_SOCIO.OTRO_TELEFONO, _OtroTelefono)
                .AddColumnValue(TABLA_SOCIO.GESTION, Me.Gestion)
                .AddColumnValue(TABLA_SOCIO.SEGMENTO, Me.Segmento)
                .AddColumnValue(TABLA_SOCIO.RUBRO, Me.Rubro)
                .AddColumnValue(TABLA_SOCIO.HABILITACION, Me.Habilitacion)
                .AddColumnValue(TABLA_SOCIO.LOCALIDAD, Me.Localidad)
                .AddColumnValue(TABLA_SOCIO.CONDICION_FISCAL, Me.CondicionFiscal)
                .AddColumnValue(TABLA_SOCIO.TARJETA_ENTREGADA, Me.TarjetaEntregada)
                .AddColumnValue(TABLA_SOCIO.TARJETA_FECHA_ENTREGA, Me.TarjetaFechaEntrega)
                .AddColumnValue(TABLA_SOCIO.CAMPANIA, Me.Campania)
                .AddColumnValue(TABLA_SOCIO.ENVIAR_MAIL, Me.EnviarMail)
            End If

            If getLastID Then
                Return .Insert(Me.InternalID)
            Else
                ' ---------------------------
                ' HACK TURNOS
                If sqlEngine.dbType = SQLEngine.dataBaseType.MYSQL Then
                    .AddColumnValue(TABLA_SOCIO.TIPO, "SOCIO")
                End If
                ' ---------------------------

                .AddColumnValue(TABLA_SOCIO.ID, Me.InternalID)
                Return .Insert()
            End If


        End With

    End Function


    ''' <summary>
    ''' Actualiza los datos de un socio
    ''' </summary>
    ''' <param name="sqlEngine">El motor de la base de datos</param>
    ''' <returns>True si la operacion se realizo correctamente, False si no</returns>
    ''' <remarks></remarks>
    Public Function Update(ByVal sqlEngine As helix.SQLEngine, Optional actualizarModificado As Boolean = True) As Boolean
        With sqlEngine.Update
            .Reset()
            .TableName = TABLA_SOCIO.TABLA_NOMBRE
            .AddColumnValue(TABLA_SOCIO.NOMBRE, _Nombre)
            .AddColumnValue(TABLA_SOCIO.APELLIDO, _Apellido)
            .AddColumnValue(TABLA_SOCIO.NACIONALIDAD, _Nacionalidad)
            .AddColumnValue(TABLA_SOCIO.DNI, _DNI)
            .AddColumnValue(TABLA_SOCIO.FECHA_NACIMIENTO, _FechaNacimiento)
            .AddColumnValue(TABLA_SOCIO.CUIT, _CUIT)
            .AddColumnValue(TABLA_SOCIO.MAIL, _Mail)
            .AddColumnValue(TABLA_SOCIO.FIRMA, _Firma)
            .AddColumnValue(TABLA_SOCIO.TIPO_EMPRESA, _TipoFirma)
            .AddColumnValue(TABLA_SOCIO.DOMICILIO, _Domicilio)
            .AddColumnValue(TABLA_SOCIO.TELEFONO, _Telefono)
            .AddColumnValue(TABLA_SOCIO.TIPO_SOCIO, _Tipo)
            .AddColumnValue(TABLA_SOCIO.NUMERO, _Numero)
            .AddColumnValue(TABLA_SOCIO.FECHA_APROBACION, _FechaAceptacion)
            .AddColumnValue(TABLA_SOCIO.ACTA, _ActaNumero)
            .AddColumnValue(TABLA_SOCIO.SOCIO_PADRINO1, _PresentadoPor1)
            .AddColumnValue(TABLA_SOCIO.SOCIO_PADRINO2, _PresentadoPor2)
            .AddColumnValue(TABLA_SOCIO.SECTOR, _Sector)
            .AddColumnValue(TABLA_SOCIO.CAJA_SEGURIDAD, _TieneCajaSeguridad)
            .AddColumnValue(TABLA_SOCIO.DELETED, _Deleted)
            .AddColumnValue(TABLA_SOCIO.MOTIVO_BAJA, _MotivoBaja)
            .AddColumnValue(TABLA_SOCIO.ESTADO, _Estado)
            If actualizarModificado Then
                .AddColumnValue(TABLA_SOCIO.MODIFICADO, Now())
            Else
                .AddColumnValue(TABLA_SOCIO.MODIFICADO, Modificado)
            End If
            ' ---------------------------
            ' HACK TURNOS
            If sqlEngine.dbType = SQLEngine.dataBaseType.MYSQL Then
                .AddColumnValue(TABLA_SOCIO.TIPO, "SOCIO")
            Else
                .AddColumnValue(TABLA_SOCIO.CELULAR, _Celular)
                .AddColumnValue(TABLA_SOCIO.OTRO_TELEFONO, _OtroTelefono)
                .AddColumnValue(TABLA_SOCIO.GESTION, Me.Gestion)
                .AddColumnValue(TABLA_SOCIO.SEGMENTO, Me.Segmento)
                .AddColumnValue(TABLA_SOCIO.RUBRO, Me.Rubro)
                .AddColumnValue(TABLA_SOCIO.HABILITACION, Me.Habilitacion)
                .AddColumnValue(TABLA_SOCIO.LOCALIDAD, Me.Localidad)
                .AddColumnValue(TABLA_SOCIO.CONDICION_FISCAL, Me.CondicionFiscal)
                .AddColumnValue(TABLA_SOCIO.TARJETA_ENTREGADA, Me.TarjetaEntregada)
                .AddColumnValue(TABLA_SOCIO.TARJETA_FECHA_ENTREGA, Me.TarjetaFechaEntrega)
                .AddColumnValue(TABLA_SOCIO.CAMPANIA, Me.Campania)
                .AddColumnValue(TABLA_SOCIO.ENVIAR_MAIL, Me.EnviarMail)
            End If
            ' ---------------------------
            .WHEREstring = TABLA_SOCIO.ID & " = ?"
            .AddWHEREparam(_InternalID)

            Return .Update
        End With
    End Function

    ''' <summary>
    ''' Elimina un socio del sistema
    ''' </summary>
    ''' <param name="sqlEngine">El motor de base de datos</param>
    ''' <returns>True si la operacion fue realizada con exito, False si no</returns>
    ''' <remarks>En realidad el registro no se borra de la base de datos, se marca como "borrada"</remarks>
    Public Function Delete(ByVal sqlEngine As helix.SQLEngine, Optional id As Integer = 0, Optional motivoDeBaja As String = "", Optional hard As Boolean = False) As Boolean
        If hard Then
            With sqlEngine.Delete
                .Reset()
                .TableName = TABLA_SOCIO.TABLA_NOMBRE
                .WHEREstring = TABLA_SOCIO.ID & " = ?"
                If id <> 0 Then
                    .AddWHEREparam(id)
                Else
                    .AddWHEREparam(_InternalID)
                End If
                Return .Delete()
            End With
        Else

            With sqlEngine.Update
                .Reset()
                .TableName = TABLA_SOCIO.TABLA_NOMBRE
                .AddColumnValue(TABLA_SOCIO.MOTIVO_BAJA, motivoDeBaja)
                .AddColumnValue(TABLA_SOCIO.DELETED, True)
                .AddColumnValue(TABLA_SOCIO.MODIFICADO, Now)
                .WHEREstring = TABLA_SOCIO.ID & " = ?"
                If id <> 0 Then
                    .AddWHEREparam(id)
                Else
                    .AddWHEREparam(_InternalID)
                End If


                Return .Update
            End With
        End If

    End Function



    ''' <summary>
    ''' Elimina todos los socios de la base de datos
    ''' </summary>
    ''' <param name="sqlEngine">Motor de base de datos</param>
    ''' <param name="hard">True elimina a todos los socios definitivamente, False los marca como eliminados</param>
    ''' <returns></returns>
    Public Function DeleteAll(ByVal sqlEngine As helix.SQLEngine, ByVal hard As Boolean) As Boolean

        If hard Then
            With sqlEngine.Delete
                .Reset()
                .TableName = TABLA_SOCIO.TABLA_NOMBRE
                Return .DeleteAll
            End With
        Else
            With sqlEngine.Update
                .Reset()
                .TableName = TABLA_SOCIO.TABLA_NOMBRE
                .AddColumnValue(TABLA_SOCIO.DELETED, True)
                .AddColumnValue(TABLA_SOCIO.MODIFICADO, Now)
                .WHEREstring = TABLA_SOCIO.DELETED & " = ?"
                .AddWHEREparam(False)
                Return .Update
            End With

        End If

    End Function

    Public Function Restaurar(ByVal sqlEngine As helix.SQLEngine, Optional id As Integer = 0) As Boolean
        With sqlEngine.Update
            .Reset()
            .TableName = TABLA_SOCIO.TABLA_NOMBRE
            .AddColumnValue(TABLA_SOCIO.DELETED, False)
            .AddColumnValue(TABLA_SOCIO.MODIFICADO, Now)
            .AddColumnValue(TABLA_SOCIO.MOTIVO_BAJA, "")
            .WHEREstring = TABLA_SOCIO.ID & " = ?"

            If id <> 0 Then
                .AddWHEREparam(id)
            Else
                .AddWHEREparam(_InternalID)
            End If

            Return .Update
        End With
    End Function

    Public Function LoadMe(ByVal sqleEngine As helix.SQLEngine, ByVal myId As Integer, Optional porNumero As Integer = 0) As Boolean
        Dim ConsoleOut As New ConsoleOut

        Try
            With sqleEngine.Query
                .Reset()
                .TableName = TABLA_SOCIO.TABLA_NOMBRE
                .AddSelectColumn(TABLA_SOCIO.ALL)
                If porNumero = 0 Then
                    .SimpleSearch(TABLA_SOCIO.ID, helix.SQLEngineQuery.OperatorCriteria.Igual, myId)
                Else
                    .WHEREstring = TABLA_SOCIO.NUMERO & " = ? AND " & TABLA_SOCIO.DELETED & " = ?"
                    .AddWHEREparam(porNumero)
                    .AddWHEREparam(False)
                End If

                If .Query Then
                    If .RecordCount >= 1 Then
                        .QueryRead()
                        ' Asignaciones de valores aquí...
                        Return True
                    End If
                End If
            End With
        Catch ex As Exception
            ConsoleOut.Print("Error en LoadMe: " & ex.Message)
        End Try

        Return False
    End Function


    Public Function LoadAll(ByVal sqleEngine As helix.SQLEngine, ByRef dt As DataTable, Optional activeOnly As Boolean = True) As Boolean
        With sqleEngine.Query
            .Reset()
            .TableName = TABLA_SOCIO.TABLA_NOMBRE
            .AddSelectColumn(TABLA_SOCIO.ALL)
            If activeOnly Then
                .SimpleSearch(TABLA_SOCIO.DELETED, helix.SQLEngineQuery.OperatorCriteria.Igual, False)
            End If

            .AddOrderColumn(TABLA_SOCIO.NOMBRE, SQLEngineQuery.sortOrder.ascending)
            Return .Query(True, dt)
        End With
    End Function

    Public Function GenerarCuotas(ByVal sqle As helix.SQLEngine, ByVal periodicidad As Byte, ByVal desdeAnio As Integer, Optional generarDesdePrincipioAnio As Boolean = True) As Boolean
        Dim cantidadCuotasGenerar As Integer = 0
        Dim cantidadCuotasAnuales As Integer = 0

        Select Case periodicidad
            Case 0
                cantidadCuotasAnuales = 12
            Case 1
                cantidadCuotasAnuales = 6
            Case 2
                cantidadCuotasAnuales = 4
            Case 3
                cantidadCuotasAnuales = 3
            Case 4
                cantidadCuotasAnuales = 2
            Case 5
                cantidadCuotasAnuales = 1
        End Select

        Dim principioAnio As Date
        Date.TryParse("1/1/" & desdeAnio, principioAnio)

        Dim cuota As New CuotaSocio
        cantidadCuotasGenerar = (Math.Abs(Now.Year - desdeAnio) * cantidadCuotasAnuales) + (cuota.DeterminePeriodFromDate(Now.Date, periodicidad) + 1)

        Dim tipoSocio As New SocioTipo
        tipoSocio.sqle = sqle
        tipoSocio.LoadMe(Me.Tipo)


        Dim baseAnio As Integer = Now.Year - Math.Abs(Now.Year - desdeAnio)
        Dim controlCuota As Integer = -1

        For i = 0 To cantidadCuotasGenerar - 1
            Dim tmpCuota As New CuotaSocio

            If controlCuota > cantidadCuotasAnuales Then
                controlCuota = 0
                baseAnio += 1
            Else
                controlCuota += 1
            End If

            'If Not TieneCofre(sqle) Then
            If Not tmpCuota.CuotaExist(sqle, controlCuota, periodicidad, baseAnio, Me.InternalID) Then
                tmpCuota.anio = baseAnio
                tmpCuota.deleted = False
                tmpCuota.observaciones = "Recibo cuota " & Utils.GetNombreMes(controlCuota) & " " & baseAnio
                tmpCuota.cobradorID = 1
                tmpCuota.monto = tipoSocio.importe
                tmpCuota.PlanID = tipoSocio.id
                tmpCuota.socioID = Me.InternalID
                tmpCuota.Periodo = controlCuota
                tmpCuota.Periodicidad = periodicidad
                Debug.Print(Me.Apellido)
                Debug.Print(Utils.GetNombreMes(controlCuota) & " " & baseAnio)

                Debug.Print(Now.Date)
                Debug.Print(tmpCuota.fechaVencimiento)
                Debug.Print(Now.CompareTo(tmpCuota.fechaVencimiento))

                If Now.CompareTo(tmpCuota.fechaVencimiento) > 0 Then
                    tmpCuota.estado = CuotaSocio.ESTADO_SOCIO.MOROSO
                Else
                    tmpCuota.estado = CuotaSocio.ESTADO_SOCIO.PENDIENTE
                End If

                If tmpCuota.fechaVencimiento >= Me.FechaAceptacion Then
                    tmpCuota.Save(sqle, 0)
                End If
            End If
            'End If
        Next
        Return True
    End Function

    Public Function TieneCofre(ByVal sqle As helix.SQLEngine) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_CONTRATOS_COFRES.TABLA_NOMBRE
            .AddSelectColumn(TABLA_CONTRATOS_COFRES.ID)
            .WHEREstring = TABLA_CONTRATOS_COFRES.ES_SOCIO_ID & " = ? AND " & TABLA_CONTRATOS_COFRES.DELETED & " = ?"
            .AddWHEREparam(Me.InternalID)
            .AddWHEREparam(False)
            If .Query Then
                If .RecordCount > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        End With
        Return False
    End Function

    Public Function CheckEstado(ByVal sqle As helix.SQLEngine, Optional ByVal desde As Date = #1/1/9999#,
                                Optional ByVal hasta As Date = #1/1/9999#) As Byte

        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ID)
            .WHEREstring = TABLA_PAGO_SOCIOS.SOCIO & " = ? AND " &
                           TABLA_PAGO_SOCIOS.ESTADO & " = ? AND " &
                           TABLA_PAGO_SOCIOS.DELETED & " = ?"

            .AddWHEREparam(Me.InternalID)
            .AddWHEREparam(2)
            .AddWHEREparam(False)

            If desde.Date <> #1/1/9999# And hasta.Date <> #1/1/9999# Then
                .WHEREstring &= " AND (" & TABLA_PAGO_SOCIOS.FECHA_VENCIMIENTO & " BETWEEN ? AND ?)"
                .AddWHEREparam(desde)
                .AddWHEREparam(hasta.AddDays(1))
            End If

            If .Query() Then
                .QueryRead()
                If .RecordCount > 0 Then
                    Me.Estado = 2
                    Return 2
                End If
            End If
        End With

        With sqle.Query
            .Reset()
            .TableName = TABLA_PAGO_SOCIOS.TABLA_NOMBRE
            .AddSelectColumn(TABLA_PAGO_SOCIOS.ID)
            .WHEREstring = TABLA_PAGO_SOCIOS.SOCIO & " = ? AND " &
                           TABLA_PAGO_SOCIOS.ESTADO & " = ? AND " &
                           TABLA_PAGO_SOCIOS.DELETED & " = ?"

            .AddWHEREparam(Me.InternalID)
            .AddWHEREparam(1)
            .AddWHEREparam(False)

            If desde.Date <> #1/1/9999# And hasta.Date <> #1/1/9999# Then
                .WHEREstring &= " AND (" & TABLA_PAGO_SOCIOS.FECHA_VENCIMIENTO & " BETWEEN ? AND ?)"
                .AddWHEREparam(desde)
                .AddWHEREparam(hasta.AddDays(1))
            End If

            If .Query() Then
                .QueryRead()
                If .RecordCount > 0 Then
                    Me.Estado = 1
                    Return 1
                Else
                    Me.Estado = 0
                    Return 0
                End If
            End If
        End With

        Return 0
    End Function

    Public Function GetTotalPagarPerido(ByVal sqle As helix.SQLEngine,
                                        ByVal semestre As Byte, ByVal desdeAnio As Integer,
                                        Optional ByRef listaCuotas As List(Of Decimal) = Nothing) As Decimal
        Dim cantidadCuotasGenerar As Integer = 0
        Dim cantidadCuotasAnuales As Integer = 0

        Dim miTipo As New SocioTipo
        miTipo.sqle = sqle
        miTipo.LoadMe(Me.Tipo)

        Select Case miTipo.periodicidad
            Case 0
                cantidadCuotasAnuales = 12
            Case 1
                cantidadCuotasAnuales = 6
            Case 2
                cantidadCuotasAnuales = 4
            Case 3
                cantidadCuotasAnuales = 3
            Case 4
                cantidadCuotasAnuales = 2
            Case 5
                cantidadCuotasAnuales = 1
        End Select

        Dim cuota As New CuotaSocio
        ' cantidadCuotasGenerar = (Math.Abs(Now.Year - desdeAnio) * cantidadCuotasAnuales) + (cuota.DeterminePeriodFromDate(Now.Date, miTipo.periodicidad) + 1)

        cantidadCuotasGenerar = (Math.Abs(Now.Year - desdeAnio) * cantidadCuotasAnuales) + 6

        Dim tipoSocio As New SocioTipo
        tipoSocio.sqle = sqle
        tipoSocio.LoadMe(Me.Tipo)

        Dim total As Decimal = 0

        Dim baseAnio As Integer = Now.Year - Math.Abs(Now.Year - desdeAnio)
        Dim controlCuota As Integer = -1

        For i = 0 To cantidadCuotasGenerar - 1
            Dim tmpCuota As New CuotaSocio

            If controlCuota > cantidadCuotasAnuales Then
                controlCuota = 0
                baseAnio += 1
            Else
                controlCuota += 1
            End If

            Dim cuotaExiste As New CuotaSocio
            If Not tmpCuota.CuotaExist(sqle, controlCuota, miTipo.periodicidad, baseAnio, Me.InternalID, cuotaExiste.id) Then
                If Not IsNothing(listaCuotas) Then
                    listaCuotas.Add(miTipo.importe)
                End If
                total += miTipo.importe
            Else
                cuotaExiste.LoadMe(sqle, cuotaExiste.id)
                If cuotaExiste.estado > CuotaSocio.ESTADO_SOCIO.PENDIENTE Then
                    If Not IsNothing(listaCuotas) Then
                        listaCuotas.Add(cuotaExiste.monto)
                    End If
                    total += cuotaExiste.monto
                End If
            End If
        Next

        Return total
    End Function

    Public Shared Function GetUltimoNumeroSocio(ByVal sqle As SQLEngine) As Integer
        With sqle.Query
            .Reset()
            .TableName = TABLA_SOCIO.TABLA_NOMBRE
            .AddSelectColumn($"TOP 1 {TABLA_SOCIO.NUMERO}")
            .AddOrderColumn(TABLA_SOCIO.NUMERO, SQLEngineQuery.sortOrder.descending)
            If .Query Then
                .QueryRead()
                Return CInt(.GetQueryData(TABLA_SOCIO.NUMERO))
            Else
                Return -1
            End If
        End With
    End Function
End Class