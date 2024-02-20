Imports helix

Public Class Contacto

    Public Property sqle As New SQLEngine
    Public Property searchResult As New List(Of Contacto)

    Public Property Id As Long = 0
    Public Property NombreAMostrar As String = ""
    Public Property Nombre As String = ""
    Public Property Apellido As String = ""
    Public Property Empresa As String = ""
    Public Property Dni As Long = 0
    Public Property Cuit As Long = 0
    Public Property Mail As String = ""
    Public Property Telefono As String = ""
    Public Property Celular As String = ""
    Public Property OtroTelefono As String = ""
    Public Property Domicilio As String = ""
    Public Property Localidad As Integer = 0
    Public Property Notas As String = ""
    Public Property Deleted As Boolean = False
    Public Property Modificado As Date = Now

    Public Enum Guardar
        NUEVO = 0
        EDITAR = 1
    End Enum


    Public Structure TABLA
        Const TABLA_NOMBRE As String = "Contactos"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const NOMBRE_A_MOSTRAR As String = TABLA_NOMBRE & "_nombreAMostrar"
        Const NOMBRE As String = TABLA_NOMBRE & "_nombre"
        Const APELLIDO As String = TABLA_NOMBRE & "_apellido"
        Const EMPRESA As String = TABLA_NOMBRE & "_empresa"
        Const DNI As String = TABLA_NOMBRE & "_dni"
        Const CUIT As String = TABLA_NOMBRE & "_cuit"
        Const MAIL As String = TABLA_NOMBRE & "_mail"
        Const TELEFONO As String = TABLA_NOMBRE & "_telefono"
        Const CELULAR As String = TABLA_NOMBRE & "_celular"
        Const OTRO_TELEFONO As String = TABLA_NOMBRE & "_otroTelefono"
        Const DOMICILIO As String = TABLA_NOMBRE & "_domicilio"
        Const LOCALIDAD As String = TABLA_NOMBRE & "_localidad"
        Const NOTAS As String = TABLA_NOMBRE & "_notas"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & NOMBRE_A_MOSTRAR & ", " & NOMBRE & ", " & APELLIDO & ", " & EMPRESA & ", " & DNI & ", " & CUIT & ", " & MAIL & ", " & TELEFONO & ", " & CELULAR & ", " & OTRO_TELEFONO & ", " & DOMICILIO & ", " & LOCALIDAD & ", " & NOTAS & ", " & DELETED & ", " & MODIFICADO
    End Structure




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


    Public Function LoadMe(ByVal myID As Integer) As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .SimpleSearch(TABLA.ID, SQLEngineQuery.OperatorCriteria.Igual, myID)
            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    Id = CLng(.GetQueryData(TABLA.ID))
                    NombreAMostrar = CStr(.GetQueryData(TABLA.NOMBRE_A_MOSTRAR))
                    Nombre = CStr(.GetQueryData(TABLA.NOMBRE))
                    Apellido = CStr(.GetQueryData(TABLA.APELLIDO))
                    Empresa = CStr(.GetQueryData(TABLA.EMPRESA))
                    Dni = CLng(.GetQueryData(TABLA.DNI))
                    Cuit = CLng(.GetQueryData(TABLA.CUIT))
                    Mail = CStr(.GetQueryData(TABLA.MAIL))
                    Telefono = CStr(.GetQueryData(TABLA.TELEFONO))
                    Celular = CStr(.GetQueryData(TABLA.CELULAR))
                    OtroTelefono = CStr(.GetQueryData(TABLA.OTRO_TELEFONO))
                    Domicilio = CStr(.GetQueryData(TABLA.DOMICILIO))
                    Localidad = CInt(.GetQueryData(TABLA.LOCALIDAD))
                    Notas = CStr(.GetQueryData(TABLA.NOTAS))
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
        With sqle.Query
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
                With sqle.Insert
                    .Reset()
                    .TableName = TABLA.TABLA_NOMBRE
                    .AddColumnValue(TABLA.NOMBRE_A_MOSTRAR, NombreAMostrar)
                    .AddColumnValue(TABLA.NOMBRE, Nombre)
                    .AddColumnValue(TABLA.APELLIDO, Apellido)
                    .AddColumnValue(TABLA.EMPRESA, Empresa)
                    .AddColumnValue(TABLA.DNI, Dni)
                    .AddColumnValue(TABLA.CUIT, Cuit)
                    .AddColumnValue(TABLA.MAIL, Mail)
                    .AddColumnValue(TABLA.TELEFONO, Telefono)
                    .AddColumnValue(TABLA.CELULAR, Celular)
                    .AddColumnValue(TABLA.OTRO_TELEFONO, OtroTelefono)
                    .AddColumnValue(TABLA.DOMICILIO, Domicilio)
                    .AddColumnValue(TABLA.LOCALIDAD, Localidad)
                    .AddColumnValue(TABLA.NOTAS, Notas)
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
                With sqle.Update
                    .Reset()
                    .TableName = TABLA.TABLA_NOMBRE
                    .AddColumnValue(TABLA.NOMBRE_A_MOSTRAR, NombreAMostrar)
                    .AddColumnValue(TABLA.NOMBRE, Nombre)
                    .AddColumnValue(TABLA.APELLIDO, Apellido)
                    .AddColumnValue(TABLA.EMPRESA, Empresa)
                    .AddColumnValue(TABLA.DNI, Dni)
                    .AddColumnValue(TABLA.CUIT, Cuit)
                    .AddColumnValue(TABLA.MAIL, Mail)
                    .AddColumnValue(TABLA.TELEFONO, Telefono)
                    .AddColumnValue(TABLA.CELULAR, Celular)
                    .AddColumnValue(TABLA.OTRO_TELEFONO, OtroTelefono)
                    .AddColumnValue(TABLA.DOMICILIO, Domicilio)
                    .AddColumnValue(TABLA.LOCALIDAD, Localidad)
                    .AddColumnValue(TABLA.NOTAS, Notas)
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
            With sqle.Delete
                .Reset()
                .TableName = TABLA.TABLA_NOMBRE
                .SimpleSearch(TABLA.ID, SQLEngineUpdate.OperatorCriteria.Igual, Id)
                Return .Delete
            End With
        Else
            With sqle.Update
                .Reset()
                .TableName = TABLA.TABLA_NOMBRE
                .AddColumnValue(TABLA.DELETED, True)
                .AddColumnValue(TABLA.MODIFICADO, Now)
                .SimpleSearch(TABLA.ID, SQLEngineUpdate.OperatorCriteria.Igual, Id)
                Return .Update
            End With
        End If
    End Function

    Public Function QuickSearch(ByVal columna As String, ByVal operador As SQLEngineQuery.OperatorCriteria, ByVal value As Object, Optional ByVal orderBy As String = "") As Integer
        With sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .SimpleSearch(columna.ToString, operador, value)
            .WHEREstring &= $" AND {TABLA.DELETED} = ?"
            .AddWHEREparam(False)

            If orderBy.Length > 0 Then
                .AddOrderColumn(orderBy, SQLEngineQuery.sortOrder.ascending)
            End If
            If .Query() Then
                Me.searchResult.Clear()
                While .QueryRead
                    Dim tmp As New Contacto
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.NombreAMostrar = CStr(.GetQueryData(TABLA.NOMBRE_A_MOSTRAR))
                    tmp.Nombre = CStr(.GetQueryData(TABLA.NOMBRE))
                    tmp.Apellido = CStr(.GetQueryData(TABLA.APELLIDO))
                    tmp.Empresa = CStr(.GetQueryData(TABLA.EMPRESA))
                    tmp.Dni = CLng(.GetQueryData(TABLA.DNI))
                    tmp.Cuit = CLng(.GetQueryData(TABLA.CUIT))
                    tmp.Mail = CStr(.GetQueryData(TABLA.MAIL))
                    tmp.Telefono = CStr(.GetQueryData(TABLA.TELEFONO))
                    tmp.Celular = CStr(.GetQueryData(TABLA.CELULAR))
                    tmp.OtroTelefono = CStr(.GetQueryData(TABLA.OTRO_TELEFONO))
                    tmp.Domicilio = CStr(.GetQueryData(TABLA.DOMICILIO))
                    tmp.Localidad = CInt(.GetQueryData(TABLA.LOCALIDAD))
                    tmp.Notas = CStr(.GetQueryData(TABLA.NOTAS))
                    tmp.Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    tmp.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))
                    searchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function

    Public Function Search(ByVal query As String) As Integer
        With sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            If IsNumeric(query) Then
                .WHEREstring = $"({TABLA.DNI} = ? OR {TABLA.CUIT} = ?)"
                .AddWHEREparam(query)
                .AddWHEREparam(query)
            Else
                .WHEREstring = $"({TABLA.NOMBRE_A_MOSTRAR} LIKE ? OR {TABLA.NOMBRE} LIKE ? OR {TABLA.APELLIDO} LIKE ?)"
                .AddWHEREparam($"%{query}%")
                .AddWHEREparam($"%{query}%")
                .AddWHEREparam($"%{query}%")
            End If
            .WHEREstring &= $" AND {TABLA.DELETED} = ?"
            .AddWHEREparam(False)

            .AddOrderColumn(TABLA.NOMBRE_A_MOSTRAR, SQLEngineQuery.sortOrder.ascending)
            If .Query() Then
                Me.searchResult.Clear()
                While .QueryRead
                    Dim tmp As New Contacto
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.NombreAMostrar = CStr(.GetQueryData(TABLA.NOMBRE_A_MOSTRAR))
                    tmp.Nombre = CStr(.GetQueryData(TABLA.NOMBRE))
                    tmp.Apellido = CStr(.GetQueryData(TABLA.APELLIDO))
                    tmp.Empresa = CStr(.GetQueryData(TABLA.EMPRESA))
                    tmp.Dni = CLng(.GetQueryData(TABLA.DNI))
                    tmp.Cuit = CLng(.GetQueryData(TABLA.CUIT))
                    tmp.Mail = CStr(.GetQueryData(TABLA.MAIL))
                    tmp.Telefono = CStr(.GetQueryData(TABLA.TELEFONO))
                    tmp.Celular = CStr(.GetQueryData(TABLA.CELULAR))
                    tmp.OtroTelefono = CStr(.GetQueryData(TABLA.OTRO_TELEFONO))
                    tmp.Domicilio = CStr(.GetQueryData(TABLA.DOMICILIO))
                    tmp.Localidad = CInt(.GetQueryData(TABLA.LOCALIDAD))
                    tmp.Notas = CStr(.GetQueryData(TABLA.NOTAS))
                    tmp.Deleted = CBool(.GetQueryData(TABLA.DELETED))
                    tmp.Modificado = CDate(.GetQueryData(TABLA.MODIFICADO))
                    searchResult.Add(tmp)
                End While
                Return .RecordCount
            End If
        End With
        Return 0
    End Function

End Class


