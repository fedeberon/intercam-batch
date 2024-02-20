Imports System.Windows.Forms
Imports helix

Public Class SocioTipo

    ''' <summary>
    ''' Motor de base de datos
    ''' </summary>
    ''' <remarks></remarks>
    Public Property sqle As New SQLEngine
    Public Property id As Integer = 0
    Public Property tipo As Integer = 0
    Public Property nombre As String = ""
    Public Property importe As Decimal = 0
    Public Property periodicidad As Integer = 0
    Public Property categoria As Byte = 0

    Public Property All As New List(Of SocioTipo)


    Private Structure dbTipoSocioIndex_Nombre
        Dim dbIndex As Integer
        Dim dbSocioTipoNombre As String
    End Structure

    Public Enum periodicidad_cuotas As Integer
        MENSUAL = 0
        BIMESTRAL = 1
        TRIMESTRAL = 2
        CUATRIMESTRAL = 3
        SEMESTRAL = 4
        ANUAL = 5
    End Enum

    Public Enum meses_x_periodo As Integer
        MENSUAL = 1
        BIMESTRAL = 2
        TRIMESTRAL = 3
        CUATRIMESTRAL = 4
        SEMESTRAL = 6
        ANUAL = 12
    End Enum

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

    Public Function getMesesPorPeriodo() As Integer
        Select Case periodicidad
            Case 0
                Return meses_x_periodo.MENSUAL
            Case 1
                Return meses_x_periodo.BIMESTRAL
            Case 2
                Return meses_x_periodo.TRIMESTRAL
            Case 3
                Return meses_x_periodo.CUATRIMESTRAL
            Case 4
                Return meses_x_periodo.SEMESTRAL
            Case 5
                Return meses_x_periodo.ANUAL

        End Select
    End Function


    ' El diccionario tiene el indice de la tabla y la cadena que representa
    ' Como se cargan en el mismo orden que el combo el indice de item es el
    ' mismo que el indice del combo
    Private _comboIndexes As New List(Of dbTipoSocioIndex_Nombre)



    ''' <summary>
    ''' Devuelve el id del tipo segun el indice que se elija
    ''' </summary>
    ''' <param name="index">Indice del item que se quiere recuperar </param>
    ''' <returns>El id del tipo de socio en la base de datos</returns>
    ''' <remarks></remarks>
    Public Function GetIDFromIndex(ByVal index As Integer) As Integer
        If index < _comboIndexes.Count Then Return _comboIndexes(index).dbIndex Else Return 0
    End Function

    ''' <summary>
    ''' Devuelve el nombre del tipo de socio segun el indice que se elija
    ''' </summary>
    ''' <param name="index">Indice del item que se quiere recuperar </param>
    ''' <returns>El nombre del tipo de socio en la base de datos</returns>
    ''' <remarks></remarks>
    Public Function GetNameFromIndex(ByVal index As Byte) As String
        If index < _comboIndexes.Count Then Return _comboIndexes(index).dbSocioTipoNombre Else Return ""
    End Function



    Public Function LoadCombo(ByRef comboToLoad As ComboBox) As Integer
        With sqle.Query
            .Reset()
            .TableName = TABLA_TIPO_SOCIO.TABLA_NOMBRE
            .AddSelectColumn(TABLA_TIPO_SOCIO.ID)
            .AddSelectColumn(TABLA_TIPO_SOCIO.NOMBRE)
            .WHEREstring = TABLA_TIPO_SOCIO.DELETED & " = ? AND " & TABLA_TIPO_SOCIO.TIPO & " = 0"
            .AddWHEREparam("0")

            If .Query Then

                comboToLoad.Items.Clear()
                Dim count As Integer = 0

                While .QueryRead
                    Dim nuevoParIndexNombre As dbTipoSocioIndex_Nombre

                    If count = 0 Then
                        comboToLoad.Items.Add("TODOS")
                        nuevoParIndexNombre.dbIndex = 0
                        nuevoParIndexNombre.dbSocioTipoNombre = "TODOS"

                        _comboIndexes.Add(nuevoParIndexNombre)
                        count += 1
                    End If

                    comboToLoad.Items.Add(.GetQueryData(1))
                    nuevoParIndexNombre.dbIndex = .GetQueryData(0)
                    nuevoParIndexNombre.dbSocioTipoNombre = .GetQueryData(1)

                    _comboIndexes.Add(nuevoParIndexNombre)
                End While
                If .RecordCount <> 0 Then comboToLoad.SelectedIndex = 0
                Return .RecordCount
            Else
                Return 0
            End If
        End With
    End Function




    Public Function Delete(ByVal dbID As Integer) As Boolean
        ' TODO: Agregar cambio en cascada si hay socios de ese tipo
        With sqle.Update
            .Reset()
            .TableName = TABLA_TIPO_SOCIO.TABLA_NOMBRE
            .AddColumnValue(TABLA_TIPO_SOCIO.DELETED, True)
            .AddColumnValue(TABLA_TIPO_SOCIO.MODIFICADO, Now())

            .WHEREstring = TABLA_TIPO_SOCIO.ID & " = ?"
            .AddWHEREparam(dbID)

            Dim result As Boolean = .Update()
            Return result
        End With
    End Function

    Public Function LoadMe(ByVal myId As Integer) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA_TIPO_SOCIO.TABLA_NOMBRE
            .AddSelectColumn(TABLA_TIPO_SOCIO.ALL)
            .SimpleSearch(TABLA_TIPO_SOCIO.ID, SQLEngineQuery.OperatorCriteria.Igual, myId)
            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    id = myId
                    tipo = .GetQueryData(TABLA_TIPO_SOCIO.TIPO)
                    nombre = .GetQueryData(TABLA_TIPO_SOCIO.NOMBRE)
                    importe = .GetQueryData(TABLA_TIPO_SOCIO.IMPORTE)
                    periodicidad = .GetQueryData(TABLA_TIPO_SOCIO.PERIODICIDAD)
                    If CStr(.GetQueryData(TABLA_TIPO_SOCIO.CATEGORIA)) <> "" Then
                        categoria = .GetQueryData(TABLA_TIPO_SOCIO.CATEGORIA)
                    Else
                        categoria = 0
                    End If
                    Return True
                End If
            End If
        End With
        Return False
    End Function

    Public Function LoadAll() As Boolean
        Dim tmpSqle As New SQLEngine
        Dim dt As New DataTable
        tmpSqle = sqle
        With sqle.Query
            .Reset()
            .TableName = TABLA_TIPO_SOCIO.TABLA_NOMBRE
            .AddSelectColumn(TABLA_TIPO_SOCIO.ALL)
            .SimpleSearch(TABLA_TIPO_SOCIO.DELETED, SQLEngineQuery.OperatorCriteria.Igual, False)
            If .Query(True, dt) Then

                Me.All.Clear()
                Dim dtr As DataTableReader
                dtr = dt.CreateDataReader


                While dtr.Read
                    Dim tmpTipo As New SocioTipo
                    tmpTipo.sqle = tmpSqle
                    tmpTipo.LoadMe(dtr(TABLA_TIPO_SOCIO.ID))
                    Me.All.Add(tmpTipo)
                End While
                Return True
            End If
        End With
        Return False
    End Function

End Class
