Imports helix
Public Class AfipAuth
    Public Property sqle As New SQLEngine
    Public Property searchResult As New List(Of AfipAuth)

    Public Property Id As Long = 0
    Public Property Servicio As String = ""
    Public Property CertId As Long = 0
    Public Property GenerationTime As Long = 0
    Public Property ExpirationTime As Long = 0
    Public Property Token As String = ""
    Public Property Sign As String = ""
    Public Property Req As String = ""
    Public Property Res As String = ""
    Public Property Homologacion As Boolean = False
    Public Property Deleted As Boolean = False
    Public Property Modificado As Date = Now
    Public Enum Guardar
        NUEVO = 0
        EDITAR = 1
    End Enum


    Public Structure TABLA
        Const TABLA_NOMBRE As String = "AfipAuth"
        Const ID As String = TABLA_NOMBRE & "_id"
        Const CERT_ID As String = TABLA_NOMBRE & "_certId"
        Const SERVICIO As String = TABLA_NOMBRE & "_servicio"
        Const GENERATION_TIME As String = TABLA_NOMBRE & "_generationTime"
        Const EXPIRATION_TIME As String = TABLA_NOMBRE & "_expirationTime"
        Const TOKEN As String = TABLA_NOMBRE & "_token"
        Const SIGN As String = TABLA_NOMBRE & "_sign"
        Const REQ As String = TABLA_NOMBRE & "_req"
        Const RES As String = TABLA_NOMBRE & "_res"
        Const HOMOLOGACION As String = TABLA_NOMBRE & "_homologacion"
        Const DELETED As String = TABLA_NOMBRE & "_deleted"
        Const MODIFICADO As String = TABLA_NOMBRE & "_modificado"
        Const ALL As String = ID & ", " & CERT_ID & ", " & SERVICIO & ", " & GENERATION_TIME & ", " & EXPIRATION_TIME & ", " & TOKEN & ", " & SIGN & ", " & REQ & ", " & RES & ", " & HOMOLOGACION & ", " & DELETED & ", " & MODIFICADO
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
                    CertId = CLng(.GetQueryData(TABLA.CERT_ID))
                    Servicio = CStr(.GetQueryData(TABLA.SERVICIO))
                    GenerationTime = CLng(.GetQueryData(TABLA.GENERATION_TIME))
                    ExpirationTime = CLng(.GetQueryData(TABLA.EXPIRATION_TIME))
                    Token = CStr(.GetQueryData(TABLA.TOKEN))
                    Sign = CStr(.GetQueryData(TABLA.SIGN))
                    Req = CStr(.GetQueryData(TABLA.REQ))
                    Res = CStr(.GetQueryData(TABLA.RES))
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


    ''' <summary>
    ''' Carga sesion activa
    ''' </summary>
    ''' <param name="ServicioAuth">Nombre del servicio que corresponde la autorizacion</param>
    ''' <param name="activeTime">Ticks de la hora que se quiere obtener la sesion</param>
    ''' <returns>True si se encontro una sesion, false si no</returns>
    Public Function LoadActive(ByVal ServicioAuth As String, ByVal activeTime As Long) As Boolean
        With Sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            '.WHEREstring = TABLA.SERVICIO & " = ? AND " & TABLA.GENERATION_TIME & " <= ? AND " & TABLA.EXPIRATION_TIME & " >= ?"
            .WHEREstring = $"{TABLA.SERVICIO}  = ? AND  {TABLA.GENERATION_TIME}  <= ? AND  {TABLA.EXPIRATION_TIME}  >= ? AND {TABLA.HOMOLOGACION} = ?"
            .AddWHEREparam(ServicioAuth)
            .AddWHEREparam(activeTime)
            .AddWHEREparam(activeTime)
            .AddWHEREparam(Homologacion)

            If .Query Then
                If .RecordCount = 1 Then
                    .QueryRead()
                    Id = CLng(.GetQueryData(TABLA.ID))
                    Servicio = CStr(.GetQueryData(TABLA.SERVICIO))
                    CertId = CLng(.GetQueryData(TABLA.CERT_ID))
                    GenerationTime = CDec(.GetQueryData(TABLA.GENERATION_TIME))
                    ExpirationTime = CDec(.GetQueryData(TABLA.EXPIRATION_TIME))
                    Token = CStr(.GetQueryData(TABLA.TOKEN))
                    Sign = CStr(.GetQueryData(TABLA.SIGN))
                    Req = CStr(.GetQueryData(TABLA.REQ))
                    Res = CStr(.GetQueryData(TABLA.RES))
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
                    .AddColumnValue(TABLA.CERT_ID, CertId)
                    .AddColumnValue(TABLA.SERVICIO, Servicio)
                    .AddColumnValue(TABLA.GENERATION_TIME, GenerationTime)
                    .AddColumnValue(TABLA.EXPIRATION_TIME, ExpirationTime)
                    .AddColumnValue(TABLA.TOKEN, Token)
                    .AddColumnValue(TABLA.SIGN, Sign)
                    .AddColumnValue(TABLA.REQ, Req)
                    .AddColumnValue(TABLA.RES, Res)
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
                    .AddColumnValue(TABLA.CERT_ID, CertId)
                    .AddColumnValue(TABLA.SERVICIO, Servicio)
                    .AddColumnValue(TABLA.GENERATION_TIME, GenerationTime)
                    .AddColumnValue(TABLA.EXPIRATION_TIME, ExpirationTime)
                    .AddColumnValue(TABLA.TOKEN, Token)
                    .AddColumnValue(TABLA.SIGN, Sign)
                    .AddColumnValue(TABLA.REQ, Req)
                    .AddColumnValue(TABLA.RES, Res)
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
    Public Function QuickSearch(ByVal columna As String, ByVal operador As SQLEngineQuery.OperatorCriteria, ByVal value As Object) As Integer
        With sqle.Query
            .Reset()
            .TableName = TABLA.TABLA_NOMBRE
            .AddSelectColumn(TABLA.ALL)
            .SimpleSearch(columna.ToString, operador, value)
            If .Query() Then
                Me.searchResult.Clear()
                While .QueryRead
                    Dim tmp As New AfipAuth
                    tmp.Id = CLng(.GetQueryData(TABLA.ID))
                    tmp.CertId = CLng(.GetQueryData(TABLA.CERT_ID))
                    tmp.Servicio = CStr(.GetQueryData(TABLA.SERVICIO))
                    tmp.GenerationTime = CLng(.GetQueryData(TABLA.GENERATION_TIME))
                    tmp.ExpirationTime = CLng(.GetQueryData(TABLA.EXPIRATION_TIME))
                    tmp.Token = CStr(.GetQueryData(TABLA.TOKEN))
                    tmp.Sign = CStr(.GetQueryData(TABLA.SIGN))
                    tmp.Req = CStr(.GetQueryData(TABLA.REQ))
                    tmp.Res = CStr(.GetQueryData(TABLA.RES))
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

