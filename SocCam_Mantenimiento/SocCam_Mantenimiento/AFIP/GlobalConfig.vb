Imports System.Drawing
Imports helix

Public Class GlobalConfig
    Public Property Sqle As New SQLEngine
    Public Property SearchResult As New List(Of GlobalConfig)

    Public Property Id As Long = 0
    Public Property Cuit As String = ""
    Public Property IngresosBrutos As String = ""
    Public Property PuntoVenta As Integer = 0
    Public Property FechaInicio As Integer = 0
    Public Property DomicilioComercial As String = ""
    Public Property NombreFantasia As String = ""
    Public Property FeLogo As Image = Nothing
    ''' <summary>
    ''' 0: nombre fantasia, 1:logo cargado, 2:nada
    ''' </summary>
    ''' <returns></returns>
    Public Property FeTipoLogo As Integer = 0
    Public Property Produccion As Boolean = False
    Public Property Deleted As Boolean = False
    Public Property Modificado As Date = Now
    Public Property UsuarioId As Integer = 0
    Public Property NombrePuesto As String = ""
    Public Property RazonSocialEmisor As String = "CAMARA COMERCIAL E INDUSTRIAL DE BOLIVAR"

    Public Enum Guardar
        NUEVO = 0
        EDITAR = 1
    End Enum

    Public Structure TABLA
        Const TABLA_NOMBRE As String = "GlobalConfig"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const CUIT As String = TABLA_NOMBRE & "_cuit"
        Const INGRESOS_BRUTOS As String = TABLA_NOMBRE & "_ingresosBrutos"
        Const PUNTO_VENTA As String = TABLA_NOMBRE & "_puntoVenta"
        Const FECHA_INICIO As String = TABLA_NOMBRE & "_fechaInicio"
        Const DOMICILIO_COMERCIAL As String = TABLA_NOMBRE & "_domicilioComercial"
        Const NOMBRE_FANTASIA As String = TABLA_NOMBRE & "_nombreFantasia"
        Const FE_LOGO As String = TABLA_NOMBRE & "_feLogo"
        Const FE_TIPO_LOGO As String = TABLA_NOMBRE & "_feTipoLogo"
        Const PRODUCCION As String = TABLA_NOMBRE & "_produccion"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & CUIT & ", " & INGRESOS_BRUTOS & ", " & PUNTO_VENTA & ", " & FECHA_INICIO & ", " & DOMICILIO_COMERCIAL & ", " & NOMBRE_FANTASIA & ", " & FE_LOGO & ", " & FE_TIPO_LOGO & ", " & PRODUCCION & ", " & DELETED & ", " & MODIFICADO

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
                    Cuit = CStr(.GetQueryData(TABLA.CUIT))
                    IngresosBrutos = CStr(.GetQueryData(TABLA.INGRESOS_BRUTOS))
                    PuntoVenta = CInt(.GetQueryData(TABLA.PUNTO_VENTA))
                    FechaInicio = CInt(.GetQueryData(TABLA.FECHA_INICIO))
                    DomicilioComercial = CStr(.GetQueryData(TABLA.DOMICILIO_COMERCIAL))
                    NombreFantasia = CStr(.GetQueryData(TABLA.NOMBRE_FANTASIA))
                    If .GetQueryData(TABLA.FE_LOGO).ToString.Length >= 0 Then
                        Utils.ByteToImage(FeLogo, .GetQueryData(TABLA.FE_LOGO))
                    End If
                    FeTipoLogo = CInt(.GetQueryData(TABLA.FE_TIPO_LOGO))
                    Produccion = CBool(.GetQueryData(TABLA.PRODUCCION))
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
                    .AddColumnValue(TABLA.CUIT, Cuit)
                    .AddColumnValue(TABLA.INGRESOS_BRUTOS, IngresosBrutos)
                    .AddColumnValue(TABLA.PUNTO_VENTA, PuntoVenta)
                    .AddColumnValue(TABLA.FECHA_INICIO, FechaInicio)
                    .AddColumnValue(TABLA.DOMICILIO_COMERCIAL, DomicilioComercial)
                    .AddColumnValue(TABLA.NOMBRE_FANTASIA, NombreFantasia)
                    If Not IsNothing(FeLogo) Then
                        .AddColumnValue(TABLA.FE_LOGO, Utils.ImageToByte(FeLogo))
                    End If
                    .AddColumnValue(TABLA.FE_TIPO_LOGO, FeTipoLogo)
                    .AddColumnValue(TABLA.PRODUCCION, Produccion)
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
                    .AddColumnValue(TABLA.CUIT, Cuit)
                    .AddColumnValue(TABLA.INGRESOS_BRUTOS, IngresosBrutos)
                    .AddColumnValue(TABLA.PUNTO_VENTA, PuntoVenta)
                    .AddColumnValue(TABLA.FECHA_INICIO, FechaInicio)
                    .AddColumnValue(TABLA.DOMICILIO_COMERCIAL, DomicilioComercial)
                    .AddColumnValue(TABLA.NOMBRE_FANTASIA, NombreFantasia)
                    If Not IsNothing(FeLogo) Then
                        .AddColumnValue(TABLA.FE_LOGO, Utils.ImageToByte(FeLogo))
                    End If
                    .AddColumnValue(TABLA.FE_TIPO_LOGO, FeTipoLogo)
                    .AddColumnValue(TABLA.PRODUCCION, Produccion)
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
                    Dim tmp As New GlobalConfig
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.Cuit = CStr(.GetQueryData(TABLA.CUIT))
                    tmp.IngresosBrutos = CStr(.GetQueryData(TABLA.INGRESOS_BRUTOS))
                    tmp.PuntoVenta = CInt(.GetQueryData(TABLA.PUNTO_VENTA))
                    tmp.FechaInicio = CInt(.GetQueryData(TABLA.FECHA_INICIO))
                    tmp.DomicilioComercial = CStr(.GetQueryData(TABLA.DOMICILIO_COMERCIAL))
                    tmp.NombreFantasia = CStr(.GetQueryData(TABLA.NOMBRE_FANTASIA))
                    If .GetQueryData(TABLA.FE_LOGO).ToString.Length >= 0 Then
                        Utils.ByteToImage(tmp.FeLogo, .GetQueryData(TABLA.FE_LOGO))
                    End If
                    tmp.FeTipoLogo = CInt(.GetQueryData(TABLA.FE_TIPO_LOGO))
                    tmp.Produccion = CBool(.GetQueryData(TABLA.PRODUCCION))
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




