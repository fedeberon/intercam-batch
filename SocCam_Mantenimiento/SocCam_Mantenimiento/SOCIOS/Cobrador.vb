Imports helix

Public Class Cobrador

    Public Property Sqle As New SQLEngine
    Public Property SearchResult As New List(Of Cobrador)


    ''' <summary>
    ''' Diccionario con informacion para buscar el nombre del cobrador en el combo cmbCobrador del formulario Cuota
    ''' </summary>
    ''' <remarks>Consta del key "nombre apellio" y el valor es el item index en el combo</remarks>
    Private _dictCobradorCuotaCombo As New List(Of Integer)

    Public Property ID As Integer = 0

    Public Property Nombre As String = ""

    Public Property Apellido As String = ""

    Public Property Domicilio As String = ""

    Public Property Telefono As String = ""

    Public Property Email As String = ""

    Public Property DNI As String = ""

    Public Property Sector As Integer = -1

    Public Property Comision As Decimal = -1

    Public Property ComisionFija As Decimal = 0



    ''' <summary>
    ''' Marca un cobrador como eliminado
    ''' </summary>
    ''' <param name="sqle">Motor de base de datos</param>
    ''' <param name="ID">El id de la base de datos del cobrador a eliminar</param>
    ''' <returns>TRUE si se elimino con exito, FALSE si no</returns>
    ''' <remarks></remarks>
    Public Function Delete(ByVal sqle As SQLEngine, ByVal ID As Integer) As Boolean
        With sqle.Update
            .Reset()
            .TableName = TABLA_COBRADORES.TABLA_NOMBRE
            .AddColumnValue(TABLA_COBRADORES.DELETED, True)

            .WHEREstring = TABLA_COBRADORES.ID & " = ?"
            .AddWHEREparam(ID)

            Return .Update
        End With
    End Function

    ''' <summary>
    ''' Reinicia todos los valores por defecto
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Reset()
        _dictCobradorCuotaCombo.Clear()
        ID = 0
        Nombre = ""
        Apellido = ""
        Domicilio = ""
        Telefono = ""
        Email = ""
        DNI = ""
        Sector = -1
        Comision = -1
        ComisionFija = 0
    End Sub

    Public Function Save(ByVal sqle As SQLEngine, ByVal editMode As Byte) As Integer
        Select Case editMode
            Case 0
                With sqle.Insert
                    .Reset()
                    .TableName = TABLA_COBRADORES.TABLA_NOMBRE

                    .AddColumnValue(TABLA_COBRADORES.NOMBRE, Nombre)
                    .AddColumnValue(TABLA_COBRADORES.APELLIDO, Apellido)
                    .AddColumnValue(TABLA_COBRADORES.DOMICILIO, Domicilio)
                    .AddColumnValue(TABLA_COBRADORES.TELEFONO, Telefono)
                    .AddColumnValue(TABLA_COBRADORES.MAIL, Email)
                    .AddColumnValue(TABLA_COBRADORES.DNI, DNI)
                    .AddColumnValue(TABLA_COBRADORES.ZONA, Sector)
                    .AddColumnValue(TABLA_COBRADORES.COMISION, Comision)
                    .AddColumnValue(TABLA_COBRADORES.COMISION_FIJA, ComisionFija)

                    .AddColumnValue(TABLA_COBRADORES.DELETED, False)
                    .AddColumnValue(TABLA_COBRADORES.MODIFICADO, Now)

                    Dim newIndex As Integer

                    If .Insert(newIndex) Then
                        Return newIndex     ' Si guardo bien retornar el ultimo ID
                    Else
                        Return 0            ' Si no, retornar flag de error
                    End If
                End With
            Case 1
                With sqle.Update
                    .Reset()
                    .TableName = TABLA_COBRADORES.TABLA_NOMBRE

                    .AddColumnValue(TABLA_COBRADORES.NOMBRE, Nombre)
                    .AddColumnValue(TABLA_COBRADORES.APELLIDO, Apellido)
                    .AddColumnValue(TABLA_COBRADORES.DOMICILIO, Domicilio)
                    .AddColumnValue(TABLA_COBRADORES.TELEFONO, Telefono)
                    .AddColumnValue(TABLA_COBRADORES.MAIL, Email)
                    .AddColumnValue(TABLA_COBRADORES.DNI, DNI)
                    .AddColumnValue(TABLA_COBRADORES.ZONA, Sector)
                    .AddColumnValue(TABLA_COBRADORES.COMISION, Comision)
                    .AddColumnValue(TABLA_COBRADORES.COMISION_FIJA, ComisionFija)

                    .AddColumnValue(TABLA_COBRADORES.DELETED, False)
                    .AddColumnValue(TABLA_COBRADORES.MODIFICADO, Now)

                    .WHEREstring = TABLA_COBRADORES.ID & " = ?"
                    .AddWHEREparam(ID)

                    If .Update Then
                        Return ID
                    Else
                        Return 0
                    End If
                End With
        End Select
        Return 0
    End Function

    ''' <summary>
    ''' Devuelve el db ID del cobrador seleccionado en el 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDbIDFromComboCuotaIndex(ByVal index As Integer) As Integer
        If _dictCobradorCuotaCombo.Count > 0 Then
            Return _dictCobradorCuotaCombo(index)
        Else
            Return -1
        End If
    End Function

    Public Function GetAllCobradores(ByVal sqle As SQLEngine) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_COBRADORES.TABLA_NOMBRE
            .WHEREstring = TABLA_COBRADORES.DELETED & " = ?"
            .AddWHEREparam(0)

            Return .Query
        End With
    End Function

    Public Function GetIdFromSector(ByVal sqle As SQLEngine, ByVal sectorBuscar As Integer) As Integer
        With sqle.Query
            .Reset()
            .TableName = TABLA_COBRADORES.TABLA_NOMBRE
            .WHEREstring = TABLA_COBRADORES.ZONA & " = ? AND " & TABLA_COBRADORES.DELETED & " = ?"
            .AddWHEREparam(sectorBuscar)
            .AddWHEREparam(False)

            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    Return CInt(.GetQueryData(TABLA_COBRADORES.ID))
                End If
            End If
        End With

        Return 0
    End Function


    Public Function LoadMe(ByVal sqle As SQLEngine, ByVal myId As Integer, Optional buscarPorSector As Boolean = False) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_COBRADORES.TABLA_NOMBRE
            .AddSelectColumn(TABLA_COBRADORES.ALL)
            If buscarPorSector Then
                .SimpleSearch(TABLA_COBRADORES.ZONA, SQLEngineQuery.OperatorCriteria.Igual, myId)
            Else
                .SimpleSearch(TABLA_COBRADORES.ID, SQLEngineQuery.OperatorCriteria.Igual, myId)
            End If
            If .Query() Then
                If .RecordCount > 0 Then
                    .QueryRead()
                    ID = CInt(.GetQueryData(TABLA_COBRADORES.ID))
                    Nombre = .GetQueryData(TABLA_COBRADORES.NOMBRE)
                    Apellido = .GetQueryData(TABLA_COBRADORES.APELLIDO)
                    Domicilio = .GetQueryData(TABLA_COBRADORES.DOMICILIO)
                    Telefono = .GetQueryData(TABLA_COBRADORES.TELEFONO)
                    Email = .GetQueryData(TABLA_COBRADORES.MAIL)
                    DNI = .GetQueryData(TABLA_COBRADORES.DNI)
                    Sector = .GetQueryData(TABLA_COBRADORES.ZONA)
                    Comision = .GetQueryData(TABLA_COBRADORES.COMISION)
                    ComisionFija = .GetQueryData(TABLA_COBRADORES.COMISION_FIJA)

                    Return True

                Else
                    Return False
                End If
            End If
        End With
        Return False
    End Function

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

    Public Function LoadAll(ByRef dt As DataTable) As Boolean
        With Sqle.Query
            .Reset()
            .TableName = TABLA_COBRADORES.TABLA_NOMBRE
            .AddSelectColumn(TABLA_COBRADORES.ALL)
            .SimpleSearch(TABLA_COBRADORES.DELETED, SQLEngineQuery.OperatorCriteria.Igual, False)
            Return .Query(True, dt)
        End With
    End Function

    Public Function QuickSearch(ByVal columna As String, ByVal operador As SQLEngineQuery.OperatorCriteria, ByVal value As Object) As Integer
        With Sqle.Query
            .Reset()
            .TableName = TABLA_COBRADORES.TABLA_NOMBRE
            .AddSelectColumn(TABLA_COBRADORES.ALL)
            .SimpleSearch(columna.ToString, operador, value)
            If .Query() Then
                Me.SearchResult.Clear()
                While .QueryRead
                    Dim tmp As New Cobrador
                    tmp.ID = CLng(.GetQueryData(TABLA_COBRADORES.ID))
                    tmp.Nombre = CStr(.GetQueryData(TABLA_COBRADORES.NOMBRE))
                    tmp.Apellido = CStr(.GetQueryData(TABLA_COBRADORES.APELLIDO))
                    tmp.Domicilio = CStr(.GetQueryData(TABLA_COBRADORES.DOMICILIO))
                    tmp.Telefono = CStr(.GetQueryData(TABLA_COBRADORES.TELEFONO))
                    tmp.Email = CStr(.GetQueryData(TABLA_COBRADORES.MAIL))
                    tmp.DNI = CStr(.GetQueryData(TABLA_COBRADORES.DNI))
                    tmp.Sector = CInt(.GetQueryData(TABLA_COBRADORES.ZONA))
                    tmp.Comision = CDec(.GetQueryData(TABLA_COBRADORES.COMISION))
                    tmp.ComisionFija = CDec(.GetQueryData(TABLA_COBRADORES.COMISION_FIJA))
                    SearchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function



End Class
