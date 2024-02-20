Imports helix
Imports System.Security
Imports System.Security.Cryptography
Imports System.Security.Cryptography.Pkcs
Imports System.Security.Cryptography.X509Certificates

Public Class AfipCert
    Public Property sqle As New SQLEngine
    Public Property searchResult As New List(Of AfipCert)

    Public Property Id As Long = 0
    Public Property Certificado As New X509Certificate2
    Public Property Password As String = ""
    Public Property Homologacion As Boolean = False
    Public Property Deleted As Boolean = False
    Public Property Modificado As Date = Now

    Public Enum Guardar
        NUEVO = 0
        EDITAR = 1
    End Enum


    Public Structure TABLA
        Const TABLA_NOMBRE As String = "AfipCert"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const CERTIFICADO As String = TABLA_NOMBRE & "_certificado"
        Const PASSWORD As String = TABLA_NOMBRE & "_password"
        Const HOMOLOGACION As String = TABLA_NOMBRE & "_homologacion"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & CERTIFICADO & ", " & PASSWORD & ", " & HOMOLOGACION & ", " & DELETED & ", " & MODIFICADO
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
                    Password = CStr(.GetQueryData(TABLA.PASSWORD))
                    Certificado.Import(.GetQueryData(TABLA.CERTIFICADO), Password, X509KeyStorageFlags.Exportable)
                    Homologacion = CBool(.GetQueryData(TABLA.HOMOLOGACION))
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

    Public Function LoadActive() As Boolean
        With sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .WHEREstring = $"{TABLA.DELETED} = ? AND {TABLA.HOMOLOGACION} = ?"
            .AddWHEREparam(False)
            .AddWHEREparam(Homologacion)
            .AddOrderColumn(TABLA.ID, SQLEngineQuery.sortOrder.ascending)
            If .Query Then
                If .RecordCount >= 1 Then
                    .QueryRead()
                    Id = CLng(.GetQueryData(TABLA.ID))
                    Password = CStr(.GetQueryData(TABLA.PASSWORD))
                    Certificado.Import(.GetQueryData(TABLA.CERTIFICADO), Password, X509KeyStorageFlags.Exportable)
                    Homologacion = CBool(.GetQueryData(TABLA.HOMOLOGACION))
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
                    .AddColumnValue(TABLA.CERTIFICADO, Certificado.Export(X509ContentType.Pfx, Password))
                    .AddColumnValue(TABLA.PASSWORD, Password)
                    .AddColumnValue(TABLA.HOMOLOGACION, Homologacion)
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
                    .AddColumnValue(TABLA.CERTIFICADO, Certificado)
                    .AddColumnValue(TABLA.PASSWORD, Password)
                    .AddColumnValue(TABLA.HOMOLOGACION, Homologacion)
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

    Public Function DeleteAll(Optional butThis As Boolean = True) As Boolean
        With sqle.Update
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddColumnValue(TABLA.DELETED, True)
            .AddColumnValue(TABLA.MODIFICADO, Now)
            If butThis Then
                .WHEREstring = $"{TABLA.ID} <> ? AND {TABLA.HOMOLOGACION} = ?"
                .AddWHEREparam(Id)
                .AddWHEREparam(Homologacion)
            Else
                .SimpleSearch(TABLA.ID, SQLEngineUpdate.OperatorCriteria.MayorIgual, 0)
            End If
            Return .Update
        End With
    End Function

    Public Function QuickSearch(ByVal columna As String, ByVal operador As SQLEngineQuery.OperatorCriteria, ByVal value As Object) As Integer
        With sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .SimpleSearch(columna.ToString, operador, value)
            If .Query() Then
                Me.searchResult.Clear()
                While .QueryRead
                    Dim tmp As New AfipCert
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.Certificado.Import(.GetQueryData(TABLA.CERTIFICADO))
                    tmp.Password = CStr(.GetQueryData(TABLA.PASSWORD))
                    tmp.Homologacion = CBool(.GetQueryData(TABLA.HOMOLOGACION))
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

